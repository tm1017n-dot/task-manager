"""Generate a standalone UTF-16LE VBScript installer from the latest source tree.

The emitted file embeds every application, sample, document and Outlook helper
as base64; it needs no ZIP extractor, PowerShell, network access or admin rights.
"""
from __future__ import annotations

import base64
import hashlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "release" / "業務ポータル_一括インストーラー_latest.vbs"
SOURCES = [
    *(p for theme in ("light", "dark") for p in sorted((ROOT / "source" / theme).rglob("*")) if p.is_file()),
    *(p for p in sorted((ROOT / "outlook").iterdir()) if p.is_file()),
]


HEADER = r'''Option Explicit
' Standalone installer for the latest HTML portal. Generated: do not edit by hand.
' The shared JSON and browser-local journal are never copied or overwritten.
Dim fso, shell, stage, installDir, backupDir, oldMoved, edgePath, shortcut, answer
Dim data, localBase, outlookInstaller, rc
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
localBase = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
If localBase = "%LOCALAPPDATA%" Or Not fso.FolderExists(localBase) Then
  MsgBox "LOCALAPPDATA is unavailable.", vbCritical, "Work Portal setup"
  WScript.Quit 1
End If
installDir = fso.BuildPath(localBase, "WorkPortal")
If MsgBox("Install the latest Work Portal for this Windows user?" & vbCrLf & _
  "Location: " & installDir & vbCrLf & _
  "Existing shared JSON will not be modified." & vbCrLf & _
  "The classic Outlook helper can be configured after installation.", _
  vbYesNo + vbQuestion, "Work Portal setup") <> vbYes Then WScript.Quit 0
On Error Resume Next
stage = fso.BuildPath(localBase, "WorkPortal_stage_" & fso.GetTempName())
backupDir = fso.BuildPath(localBase, "WorkPortal_backup_" & fso.GetTempName())
oldMoved = False
fso.CreateFolder stage
Check "Create staging folder"

Sub Check(label)
  Dim detail
  If Err.Number = 0 Then Exit Sub
  detail = label & ": " & Err.Description
  Err.Clear
  AbortInstall detail
End Sub

Sub AbortInstall(detail)
  On Error Resume Next
  If stage <> "" Then If fso.FolderExists(stage) Then fso.DeleteFolder stage, True
  If oldMoved And Not fso.FolderExists(installDir) Then fso.MoveFolder backupDir, installDir
  MsgBox "Installation stopped: " & detail & vbCrLf & _
    "If a previous copy was moved, its backup is at " & backupDir, vbCritical, "Work Portal setup"
  WScript.Quit 1
End Sub

Sub EnsureFolder(folder)
  On Error Resume Next
  If fso.FolderExists(folder) Then Exit Sub
  EnsureFolder fso.GetParentFolderName(folder)
  fso.CreateFolder folder
  Check "Create " & folder
End Sub

Sub SaveEmbedded(relativePath, expectedSize, encoded)
  Dim filePath, xml, node, stream
  On Error Resume Next
  filePath = fso.BuildPath(stage, relativePath)
  EnsureFolder fso.GetParentFolderName(filePath)
  Set xml = CreateObject("MSXML2.DOMDocument.6.0")
  Check "Load decoder"
  Set node = xml.createElement("binary")
  node.dataType = "bin.base64"
  node.Text = encoded
  Check "Decode " & relativePath
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 1
  stream.Open
  stream.Write node.nodeTypedValue
  Check "Write " & relativePath
  stream.SaveToFile filePath, 2
  Check "Save " & relativePath
  stream.Close
  Check "Close " & relativePath
  If fso.GetFile(filePath).Size <> expectedSize Then AbortInstall "Incorrect size: " & relativePath
  Check "Verify " & relativePath
End Sub

' Embedded files follow.
'''


FOOTER = r'''
' All files have been decoded and checked before replacing an existing install.
If fso.FolderExists(installDir) Then
  fso.MoveFolder installDir, backupDir
  Check "Back up previous install"
  oldMoved = True
End If
fso.MoveFolder stage, installDir
Check "Activate new install"
stage = ""

edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
If Not fso.FileExists(edgePath) Then edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
If fso.FileExists(edgePath) Then
  Set shortcut = shell.CreateShortcut(fso.BuildPath(shell.SpecialFolders("Desktop"), "Work Portal.lnk"))
  shortcut.TargetPath = edgePath
  shortcut.Arguments = Chr(34) & fso.BuildPath(installDir, "source\light\index.html") & Chr(34)
  shortcut.WorkingDirectory = fso.BuildPath(installDir, "source\light")
  shortcut.Description = "Work Portal (light and dark switching available)"
  shortcut.IconLocation = fso.BuildPath(installDir, "source\light\icon.ico") & ",0"
  shortcut.Save
  Err.Clear ' A locked down Desktop must not invalidate the installed application.
End If

answer = MsgBox("Installed at " & installDir & vbCrLf & _
  "Launch source\light\index.html with Microsoft Edge." & vbCrLf & _
  "Choose the existing shared JSON in the portal. Do not use a sample for production." & vbCrLf & _
  "Would you like to configure classic Outlook now?", vbYesNo + vbQuestion, "Work Portal setup")
If answer = vbYes Then
  outlookInstaller = fso.BuildPath(installDir, "outlook\Install_Outlook_Sync.vbs")
  rc = shell.Run(Chr(34) & shell.ExpandEnvironmentStrings("%SystemRoot%") & _
    "\System32\wscript.exe" & Chr(34) & " " & Chr(34) & outlookInstaller & Chr(34), 1, True)
  If Err.Number <> 0 Then
    MsgBox "Portal installed. Outlook setup could not start: " & Err.Description, vbExclamation, "Work Portal setup"
    Err.Clear
  ElseIf rc <> 0 Then
    MsgBox "Portal installed. Outlook setup ended with code " & rc & _
      ". You can run outlook\Install_Outlook_Sync.vbs later.", vbExclamation, "Work Portal setup"
  End If
End If
If oldMoved Then
  MsgBox "Installation complete. Previous application files are preserved at:" & vbCrLf & _
    backupDir & vbCrLf & "Your shared JSON was not changed.", vbInformation, "Work Portal setup"
Else
  MsgBox "Installation complete. Your shared JSON was not changed.", vbInformation, "Work Portal setup"
End If
'''


def main() -> None:
    lines = HEADER.splitlines()
    for path in SOURCES:
        rel = path.relative_to(ROOT).as_posix().replace("/", "\\")
        payload = base64.b64encode(path.read_bytes()).decode("ascii")
        lines.extend(["", "data = \"\""])
        for offset in range(0, len(payload), 800):
            lines.append(f'data = data & "{payload[offset:offset + 800]}"')
        lines.append(f'Call SaveEmbedded("{rel}", {path.stat().st_size}, data)')
    lines.extend(FOOTER.splitlines())
    OUTPUT.write_bytes(("\ufeff" + "\r\n".join(lines) + "\r\n").encode("utf-16le"))
    sums = ROOT / "checksums" / "SHA256SUMS.txt"
    old = [line for line in sums.read_text().splitlines() if not line.endswith("  " + OUTPUT.name)]
    old.append(f"{hashlib.sha256(OUTPUT.read_bytes()).hexdigest().upper()}  {OUTPUT.name}")
    sums.write_text("\n".join(old) + "\n")
    print(f"{OUTPUT}: {len(SOURCES)} files, {OUTPUT.stat().st_size} bytes")


if __name__ == "__main__":
    main()
