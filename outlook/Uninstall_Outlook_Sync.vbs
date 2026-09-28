Option Explicit
' Removes the current user's launch handler and helper. Outlook appointments remain.
Dim shell, fso, root, folder
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If MsgBox("Remove the Work Portal Outlook button handler from this PC?" & vbCrLf & _
  "Existing calendar appointments will remain.", vbYesNo + vbQuestion, "Work Portal") <> vbYes Then WScript.Quit 0
root = "HKCU\Software\Classes\workportal-outlook\"
On Error Resume Next
shell.RegDelete root & "shell\open\command\"
shell.RegDelete root & "shell\open\"
shell.RegDelete root & "shell\"
shell.RegDelete root & "URL Protocol"
shell.RegDelete root
shell.RegDelete "HKCU\Software\WorkPortalOutlook\SharedJsonPath"
shell.RegDelete "HKCU\Software\WorkPortalOutlook\"
folder = fso.BuildPath(shell.ExpandEnvironmentStrings("%APPDATA%"), "WorkPortalOutlook")
If fso.FolderExists(folder) Then fso.DeleteFolder folder, True
If Err.Number <> 0 Then
  MsgBox "Some settings could not be removed: " & Err.Description, vbExclamation, "Work Portal"
  WScript.Quit 1
End If
On Error GoTo 0
MsgBox "The helper was removed. Appointments in Outlook remain.", vbInformation, "Work Portal"
