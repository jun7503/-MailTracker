using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Configuration;
using ClosedXML.Excel;

namespace OutlookMailReaderGraph
{
    internal static class Program
    {
        // Safer default that works on any machine (no hard-coded OneDrive path)
        private static readonly string DefaultRoot =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                         "OutlookMailTracker");

        private class MailRow
        {
            public string DateLocal { get; set; } = string.Empty;
            public string FromName { get; set; } = string.Empty;
            public string FromAddress { get; set; } = string.Empty;
            public string CustomerCompany { get; set; } = string.Empty;
            public string CustomerWindow { get; set; } = string.Empty;
            public string Subject { get; set; } = string.Empty;
            public string IsRead { get; set; } = string.Empty;
            public string HasAttachments { get; set; } = string.Empty;
            public int AttachmentCount { get; set; }    // ← semicolon is required
            public string AttachmentPaths { get; set; } = string.Empty;
            public string MessageId { get; set; } = string.Empty;
        }

        private class SyncState
        {
            public DateTimeOffset? LastReceivedUtc { get; set; }
            public HashSet<string> ProcessedIds { get; set; } = new();
        }

        private static async Task<int> Main(string[] args)
        {
            try
            {
                var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: true)
                    .AddEnvironmentVariables()
                    .Build();

                var tenantId = config["TenantId"];
                var clientId = config["ClientId"];
                var userId = config["UserId"];

                if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId))
                {
                    Console.WriteLine("Please set TenantId and ClientId in appsettings.json");
                    return 1;
                }

                // Try to read OutputRoot from config; if missing/blank, fall back to DefaultRoot
                var outputRoot = config["OutputRoot"];
                string root = !string.IsNullOrWhiteSpace(outputRoot) ? outputRoot! : DefaultRoot;

                // Ensure the directory exists
                Directory.CreateDirectory(root);

                // All outputs go under 'root'
                string excelPath = Path.Combine(root, "MailTracker.xlsx");
                string attachDir = Path.Combine(root, "Attachments");
                string statePath = Path.Combine(root, "state.json");
                string attachDir = Path.Combine(root, "Attachments");
                string statePath = Path.Combine(root, "state.json");
                Directory.CreateDirectory(root);
                Directory.CreateDirectory(attachDir);

                var syncState = LoadState(statePath);
                DateTimeOffset? since = syncState.LastReceivedUtc; // null => first run => read all

                var credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    RedirectUri = new Uri("http://localhost")
                });

                var graph = new GraphServiceClient(credential, new[] { "User.Read", "Mail.Read" });

                int pageSize = 100; int processed = 0;
                var select = new[]
                {
                    "id","subject","from","toRecipients","ccRecipients","receivedDateTime","isRead","hasAttachments","bodyPreview"
                };

                Func<Task<MessageCollectionResponse?>> getPage = async () =>
                    await graph.Me.Messages.GetAsync(r =>
                    {
                        r.QueryParameters.Top = pageSize;
                        r.QueryParameters.Select = select;
                        r.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
                        if (since != null)
                            r.QueryParameters.Filter = $"receivedDateTime ge {since:O}";
                    });

                var rowsByTopic = new Dictionary<string, List<MailRow>>(StringComparer.OrdinalIgnoreCase);
                var page = await getPage();

                while (page != null)
                {
                    var list = page.Value;
                    if (list == null || list.Count == 0) break;

                    foreach (var m in list)
                    {
                        if (m.Id == null) continue;
                        if (syncState.ProcessedIds.Contains(m.Id)) continue;

                        var topic = DetectTopic(m);
                        var company = DeriveCompany(m.From?.EmailAddress?.Address);
                        var window = DeriveWindow(m);

                        List<string> saved = new();
                        int ac = 0;

                        if (m.HasAttachments == true)
                        {
                            var atts = await graph.Me.Messages[m.Id].Attachments.GetAsync();

                            if (atts?.Value != null)
                            {
                                foreach (var att in atts.Value)
                                {
                                    if (att is FileAttachment fa && fa.ContentBytes != null)
                                    {
                                        ac++;
                                        var safe = SanitizeFileName(fa.Name ?? "attachment");
                                        var datePart = (m.ReceivedDateTime ?? DateTimeOffset.UtcNow)
                                                        .ToLocalTime().ToString("yyyy-MM-dd");
                                        var dir = Path.Combine(attachDir, topic, datePart);
                                        Directory.CreateDirectory(dir);
                                        var p = Path.Combine(dir, safe);
                                        await File.WriteAllBytesAsync(p, fa.ContentBytes);
                                        saved.Add(p);
                                    }
                                }
                            }
                        }

                        var row = new MailRow
                        {
                            DateLocal = (m.ReceivedDateTime ?? DateTimeOffset.UtcNow).ToLocalTime().ToString("yyyy-MM-dd HH:mm"),
                            FromName = m.From?.EmailAddress?.Name ?? string.Empty,
                            FromAddress = m.From?.EmailAddress?.Address ?? string.Empty,
                            CustomerCompany = company,
                            CustomerWindow = window,
                            Subject = m.Subject ?? string.Empty,
                            IsRead = m.IsRead == true ? "Yes" : "No",
                            HasAttachments = (m.HasAttachments ?? false) ? "Yes" : "No",
                            AttachmentCount = ac,
                            AttachmentPaths = string.Join(';', saved),
                            MessageId = m.Id
                        };

                        if (!rowsByTopic.TryGetValue(topic, out var lst))
                        {
                            lst = new List<MailRow>();
                            rowsByTopic[topic] = lst;
                        }
                        lst.Add(row);

                        syncState.ProcessedIds.Add(m.Id);

                        var rcvUtc = m.ReceivedDateTime?.ToUniversalTime();
                        if (rcvUtc != null && (syncState.LastReceivedUtc == null || rcvUtc > syncState.LastReceivedUtc))
                            syncState.LastReceivedUtc = rcvUtc;

                        processed++;
                    }

                    if (!string.IsNullOrEmpty(page.OdataNextLink))
                    {
                        var nextReq = new Microsoft.Graph.Users.Item.Messages.MessagesRequestBuilder(page.OdataNextLink, graph.RequestAdapter);
                        page = await nextReq.GetAsync();
                    }
                    else break;
                }

                WriteExcel(excelPath, rowsByTopic, attachDir);

                var json = JsonSerializer.Serialize(syncState, new JsonSerializerOptions { WriteIndented = true });
                await File.WriteAllTextAsync(statePath, json, new UTF8Encoding(true));

                Console.WriteLine($"Done. Processed {processed} new messages. Updated: {excelPath}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return 2;
            }
        }

        private static void WriteExcel(string excelPath, Dictionary<string, List<MailRow>> rowsByTopic, string attachRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(excelPath) ?? ".");
            using var wb = File.Exists(excelPath) ? new XLWorkbook(excelPath) : new XLWorkbook();

            var overview = wb.Worksheets.FirstOrDefault(ws =>
                ws.Name.Equals("Overview", StringComparison.OrdinalIgnoreCase))
                ?? wb.AddWorksheet("Overview");

            foreach (var kv in rowsByTopic)
            {
                var topic = kv.Key;
                var wsName = SanitizeSheetName(topic);
                var ws = wb.Worksheets.FirstOrDefault(x =>
                    x.Name.Equals(wsName, StringComparison.OrdinalIgnoreCase))
                    ?? wb.AddWorksheet(wsName);

                if (ws.LastRowUsed() == null)
                {
                    ws.Cell(1, 1).Value = "Date/Time (Local)";
                    ws.Cell(1, 2).Value = "From (Name)";
                    ws.Cell(1, 3).Value = "From (Address)";
                    ws.Cell(1, 4).Value = "Customer Company";
                    ws.Cell(1, 5).Value = "Customer Window";
                    ws.Cell(1, 6).Value = "Subject";
                    ws.Cell(1, 7).Value = "Is Read";
                    ws.Cell(1, 8).Value = "Has Attachments";
                    ws.Cell(1, 9).Value = "Attachment Count";
                    ws.Cell(1, 10).Value = "Attachment Paths";
                    ws.Cell(1, 11).Value = "Message Id";
                    ws.Row(1).Style.Font.Bold = true;
                    ws.Columns().AdjustToContents(1, 11);
                }

                int r = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;
                foreach (var row in kv.Value)
                {
                    ws.Cell(r, 1).Value = row.DateLocal;
                    ws.Cell(r, 2).Value = row.FromName;
                    ws.Cell(r, 3).Value = row.FromAddress;
                    ws.Cell(r, 4).Value = row.CustomerCompany;
                    ws.Cell(r, 5).Value = row.CustomerWindow;
                    ws.Cell(r, 6).Value = row.Subject;
                    ws.Cell(r, 7).Value = row.IsRead;
                    ws.Cell(r, 8).Value = row.HasAttachments;
                    ws.Cell(r, 9).Value = row.AttachmentCount;
                    ws.Cell(r, 10).Value = row.AttachmentPaths;
                    ws.Cell(r, 11).Value = row.MessageId;
                    r++;
                }
            }

            var rows = new List<(string Topic, int Total, int Unread, int WithAtt, int Last7, string LatestDate, string LatestSender, string LatestSubject, string Folder)>();

            foreach (var ws in wb.Worksheets)
            {
                if (ws.Name.Equals("Overview", StringComparison.OrdinalIgnoreCase)) continue;
                if (ws.LastRowUsed() == null) continue;

                var topic = ws.Name;
                int total = ws.LastRowUsed().RowNumber() - 1;
                int unread = 0, withAtt = 0, last7 = 0;
                DateTime latest = DateTime.MinValue;
                string latestFrom = string.Empty;
                string latestSubj = string.Empty;

                for (int r = 2; r <= ws.LastRowUsed().RowNumber(); r++)
                {
                    var dateStr = ws.Cell(r, 1).GetString();
                    if (DateTime.TryParse(dateStr, out var dt) && dt > latest)
                    {
                        latest = dt;
                        latestFrom = ws.Cell(r, 2).GetString();
                        latestSubj = ws.Cell(r, 6).GetString();
                    }

                    if (ws.Cell(r, 7).GetString().Equals("No", StringComparison.OrdinalIgnoreCase)) unread++;
                    if (ws.Cell(r, 8).GetString().Equals("Yes", StringComparison.OrdinalIgnoreCase)) withAtt++;
                    if (DateTime.TryParse(dateStr, out var d2) && (DateTime.Now - d2).TotalDays <= 7) last7++;
                }

                var folder = Path.Combine(attachRoot, topic);
                rows.Add((topic, total, unread, withAtt, last7,
                          latest == DateTime.MinValue ? string.Empty : latest.ToString("yyyy-MM-dd HH:mm"),
                          latestFrom, latestSubj, folder));
            }

            overview.Clear();
            overview.Cell(1, 1).Value = "Topic";
            overview.Cell(1, 2).Value = "Topic Type";
            overview.Cell(1, 3).Value = "Total Emails";
            overview.Cell(1, 4).Value = "Unread Emails";
            overview.Cell(1, 5).Value = "With Attachments";
            overview.Cell(1, 6).Value = "Last 7 Days";
            overview.Cell(1, 7).Value = "Latest Mail Date";
            overview.Cell(1, 8).Value = "Latest Sender";
            overview.Cell(1, 9).Value = "Latest Subject";
            overview.Cell(1, 10).Value = "Topic Folder";
            overview.Row(1).Style.Font.Bold = true;

            int orow = 2;
            foreach (var t in rows.OrderByDescending(x => x.LatestDate))
            {
                overview.Cell(orow, 1).Value = t.Topic;
                overview.Cell(orow, 2).Value = "Project";
                overview.Cell(orow, 3).Value = t.Total;
                overview.Cell(orow, 4).Value = t.Unread;
                overview.Cell(orow, 5).Value = t.WithAtt;
                overview.Cell(orow, 6).Value = t.Last7;
                overview.Cell(orow, 7).Value = t.LatestDate;
                overview.Cell(orow, 8).Value = t.LatestSender;
                overview.Cell(orow, 9).Value = t.LatestSubject;
                overview.Cell(orow, 10).Value = t.Folder;
                orow++;
            }

            overview.Columns().AdjustToContents(1, 10);
            wb.SaveAs(excelPath);
        }

        private static SyncState LoadState(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var s = JsonSerializer.Deserialize<SyncState>(json);
                    if (s != null) return s;
                }
            }
            catch { }
            return new SyncState();
        }

        private static string ExtractDomain(string address)
        {
            var at = address.IndexOf('@');
            return at >= 0 && at < address.Length - 1 ? address[(at + 1)..] : string.Empty;
        }

        private static string DeriveCompany(string? addr)
        {
            var dom = ExtractDomain(addr ?? string.Empty);
            if (string.IsNullOrWhiteSpace(dom)) return string.Empty;
            var parts = dom.Split('.');
            return CultureInfo.InvariantCulture.TextInfo.ToTitleCase((parts.Length >= 2 ? parts[0] : dom).Replace('-', ' '));
        }

        private static string DeriveWindow(Message m)
        {
            var fromDom = ExtractDomain(m.From?.EmailAddress?.Address ?? string.Empty);
            var toDoms = new List<string>();
            if (m.ToRecipients != null) toDoms.AddRange(m.ToRecipients.Select(r => ExtractDomain(r.EmailAddress?.Address ?? string.Empty)));
            if (m.CcRecipients != null) toDoms.AddRange(m.CcRecipients.Select(r => ExtractDomain(r.EmailAddress?.Address ?? string.Empty)));

            var freq = toDoms
                .Where(d => !string.IsNullOrWhiteSpace(d) && !d.Equals(fromDom, StringComparison.OrdinalIgnoreCase))
                .GroupBy(d => d)
                .OrderByDescending(g => g.Count())
                .Select(g => g.Key)
                .FirstOrDefault();

            return string.IsNullOrWhiteSpace(freq)
                ? string.Empty
                : CultureInfo.InvariantCulture.TextInfo.ToTitleCase(freq.Split('.')[0]);
        }

        private static string SanitizeFileName(string name)
        {
            var pattern = "[" + Regex.Escape(new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars())) + "]";
            var safe = Regex.Replace(name, pattern, "_");
            return string.IsNullOrWhiteSpace(safe) ? "file" : safe;
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = new[] { '\\', '/', '*', '[', ']', ':', '?' }; // ← backslash must be escaped
            var s = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            if (s.Length > 31) s = s.Substring(0, 31);
            if (string.IsNullOrWhiteSpace(s)) s = "Sheet";
            return s;
        }

        private static string DetectTopic(Message m)
        {
            string subj = m.Subject ?? string.Empty;
            string preview = m.BodyPreview ?? string.Empty;
            string fromAddr = m.From?.EmailAddress?.Address ?? string.Empty;

            var patterns = new[]
            {
                new Regex(@"\[(.*?)\]"),
                new Regex(@"\((.*?)\)"),
                new Regex(@"(?i)\b(?:PRJ|PROJ|PROJECT)[:\s\-]+([A-Za-z0-9._\-]+)"),
                new Regex(@"(?i)\b(?:RFQ|PO|INV|BOM)[\-\_\s]*([A-Za-z0-9._\-]+)")
            };

            foreach (var re in patterns)
            {
                var m1 = re.Match(subj);
                if (m1.Success)
                {
                    var cand = m1.Groups.Cast<System.Text.RegularExpressions.Group>()
                        .Select(g => g.Value)
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .OrderByDescending(s => s.Length)
                        .FirstOrDefault();

                    if (!string.IsNullOrWhiteSpace(cand))
                        return CleanTopic(cand);
                }

                var m2 = re.Match(preview);
                if (m2.Success)
                {
                    var cand = m2.Groups.Cast<System.Text.RegularExpressions.Group>()
                        .Select(g => g.Value)
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .OrderByDescending(s => s.Length)
                        .FirstOrDefault();

                    if (!string.IsNullOrWhiteSpace(cand))
                        return CleanTopic(cand);
                }
            }

            var customer = DeriveCompany(fromAddr);
            if (!string.IsNullOrWhiteSpace(customer)) return customer;

            return "Uncategorized";
        }

        private static string CleanTopic(string raw)
        {
            var t = (raw ?? string.Empty).Trim();
            t = Regex.Replace(t, @"^\s*[\[\(]\s*(.*?)\s*[\]\)]\s*$", "$1"); // drop surrounding [] or ()
            t = Regex.Replace(t, @"^[#_\-:\s]+|[#_\-:\s]+$", "");          // trim punctuation
            return string.IsNullOrWhiteSpace(t) ? "Uncategorized" : t;
        }
    }
}
