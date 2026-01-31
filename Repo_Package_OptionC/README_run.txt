
# Mail Tracker — Repo Package (Option C)

## What is inside
- `appsettings.json` (filled with your TenantId/ClientId)
- `.github/workflows/build.yml` (GitHub Actions to build EXE)
- `OutlookMailReaderGraph/` (full source code)
- `RunMailTracker.bas` (Outlook macro that launches the EXE)

## How to use
1. Create a **private** repo on GitHub.
2. Upload everything from this package **as-is** at repo root (drag & drop in the web UI).
3. Go to **Actions** → **Build Mail Tracker EXE** → **Run workflow**.
4. Download the artifact **MailTracker_win-x64**.
5. Extract the ZIP, double‑click `OutlookMailReaderGraph.exe`, sign in.
6. First run builds your baseline; next runs only process new mail.

## Outlook button
- Import `RunMailTracker.bas` in Outlook (ALT+F11 → Import File).
- Edit `exePath` to your downloaded EXE location.
- Add the macro to Ribbon: File → Options → Customize Ribbon → Macros.
