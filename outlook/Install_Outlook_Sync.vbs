Option Explicit
' Installs the one-click handler for the current Windows user only.
' This installer and Sync_Outlook.js must be in the same directory.
Dim fso, shell, sourceFile, installDir, installedFile, jsonPath, oldPath
Dim root, wscriptExe, command, answer
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
sourceFile = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "Sync_Outlook.js")
If Not fso.FileExists(sourceFile) Then
  MsgBox "Sync_Outlook.js must be beside this installer.", vbCritical, "Work Portal"
  WScript.Quit 1
End If

oldPath = ""
On Error Resume Next
oldPath = shell.RegRead("HKCU\Software\WorkPortalOutlook\SharedJsonPath")
Err.Clear
On Error GoTo 0
jsonPath = InputBox("Enter the full path to the existing shared JSON file." & vbCrLf & _
  "Example: \\server\share\personal_work_portal_shared.json" & vbCrLf & _
  "This PC must be able to read the file when the button is pressed.", _
  "Work Portal Outlook setup", oldPath)
jsonPath = Trim(jsonPath)
If jsonPath = "" Then WScript.Quit 0
If Not fso.FileExists(jsonPath) Then
  MsgBox "The selected JSON file was not found. Nothing was installed.", vbCritical, "Work Portal"
  WScript.Quit 1
End If
If LCase(fso.GetExtensionName(jsonPath)) <> "json" Then
  MsgBox "Please select a .json file. Nothing was installed.", vbCritical, "Work Portal"
  WScript.Quit 1
End If
answer = MsgBox("Use classic Outlook? The selected shared JSON and task titles will be read by " & _
  "this PC and added to a dedicated Outlook calendar." & vbCrLf & vbCrLf & jsonPath & vbCrLf & vbCrLf & _
  "Install for this Windows user?", vbYesNo + vbQuestion, "Work Portal Outlook setup")
If answer <> vbYes Then WScript.Quit 0

On Error Resume Next
installDir = fso.BuildPath(shell.ExpandEnvironmentStrings("%APPDATA%"), "WorkPortalOutlook")
If Not fso.FolderExists(installDir) Then fso.CreateFolder installDir
installedFile = fso.BuildPath(installDir, "Sync_Outlook.js")
fso.CopyFile sourceFile, installedFile, True
If Err.Number <> 0 Then
  MsgBox "Could not copy the helper: " & Err.Description, vbCritical, "Work Portal"
  WScript.Quit 1
End If
Err.Clear
shell.RegWrite "HKCU\Software\WorkPortalOutlook\SharedJsonPath", jsonPath, "REG_SZ"
root = "HKCU\Software\Classes\workportal-outlook\"
shell.RegWrite root, "URL:Work Portal Outlook sync", "REG_SZ"
shell.RegWrite root & "URL Protocol", "", "REG_SZ"
wscriptExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\wscript.exe"
command = Chr(34) & wscriptExe & Chr(34) & " " & Chr(34) & installedFile & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
shell.RegWrite root & "shell\open\command\", command, "REG_SZ"
If Err.Number <> 0 Then
  MsgBox "Could not register the launch link: " & Err.Description, vbCritical, "Work Portal"
  WScript.Quit 1
End If
On Error GoTo 0
MsgBox "Setup completed for this Windows user." & vbCrLf & _
  "Open the portal in Edge, select yourself, and press Outlook sync." & vbCrLf & _
  "Edge may ask for permission to open the helper.", vbInformation, "Work Portal"
