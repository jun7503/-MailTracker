
Attribute VB_Name = "RunMailTracker"
Option Explicit

Public Sub RunMailTracker()
    On Error GoTo EH
    Dim exePath As String
    ' TODO: Set this to the full path of your EXE after you download from GitHub Actions
    exePath = "C:\Path\To\OutlookMailReaderGraph.exe"

    If Dir(exePath) = "" Then
        MsgBox "Mail Tracker EXE not found:
" & exePath & _
               "

Please update the path in the macro (Alt+F11) and try again.", _
               vbExclamation + vbOKOnly, "Mail Tracker"
        Exit Sub
    End If

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run Chr(34) & exePath & Chr(34), 1, False
    Exit Sub
EH:
    MsgBox "Error starting Mail Tracker: " & Err.Description, vbCritical + vbOKOnly, "Mail Tracker"
End Sub
