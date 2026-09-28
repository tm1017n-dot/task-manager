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
  MsgBox "端末内の保存先（LOCALAPPDATA）を取得できません。", vbCritical, "業務ポータルの導入"
  WScript.Quit 1
End If
installDir = fso.BuildPath(localBase, "WorkPortal")
If MsgBox("このWindows利用者に、最新版の業務ポータルを導入しますか？" & vbCrLf & _
  "導入先：" & installDir & vbCrLf & _
  "既存の共有JSONは変更しません。" & vbCrLf & _
  "導入後、必要に応じて従来版Outlookとの連携を設定できます。", _
  vbYesNo + vbQuestion, "業務ポータルの導入") <> vbYes Then WScript.Quit 0
On Error Resume Next
stage = fso.BuildPath(localBase, "WorkPortal_stage_" & fso.GetTempName())
backupDir = fso.BuildPath(localBase, "WorkPortal_backup_" & fso.GetTempName())
oldMoved = False
fso.CreateFolder stage
Check "一時フォルダーの作成"

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
  MsgBox "導入を中断しました：" & detail & vbCrLf & _
    "以前のアプリを退避した場合、その保存先：" & backupDir, vbCritical, "業務ポータルの導入"
  WScript.Quit 1
End Sub

Sub EnsureFolder(folder)
  On Error Resume Next
  If fso.FolderExists(folder) Then Exit Sub
  EnsureFolder fso.GetParentFolderName(folder)
  fso.CreateFolder folder
  Check "フォルダーの作成：" & folder
End Sub

Sub SaveEmbedded(relativePath, expectedSize, encoded)
  Dim filePath, xml, node, stream
  On Error Resume Next
  filePath = fso.BuildPath(stage, relativePath)
  EnsureFolder fso.GetParentFolderName(filePath)
  Set xml = CreateObject("MSXML2.DOMDocument.6.0")
  Check "復元機能の準備"
  Set node = xml.createElement("binary")
  node.dataType = "bin.base64"
  node.Text = encoded
  Check "ファイルの復元：" & relativePath
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 1
  stream.Open
  stream.Write node.nodeTypedValue
  Check "ファイルの書き込み：" & relativePath
  stream.SaveToFile filePath, 2
  Check "ファイルの保存：" & relativePath
  stream.Close
  Check "ファイルの終了処理：" & relativePath
  If fso.GetFile(filePath).Size <> expectedSize Then AbortInstall "復元したファイルのサイズが正しくありません：" & relativePath
  Check "ファイルの確認：" & relativePath
End Sub

' Embedded files follow.
'''


FOOTER = r'''
' All files have been decoded and checked before replacing an existing install.
If fso.FolderExists(installDir) Then
  fso.MoveFolder installDir, backupDir
  Check "以前のアプリの退避"
  oldMoved = True
End If
fso.MoveFolder stage, installDir
Check "新しいアプリへの切替"
stage = ""

edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
If Not fso.FileExists(edgePath) Then edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
If fso.FileExists(edgePath) Then
  Set shortcut = shell.CreateShortcut(fso.BuildPath(shell.SpecialFolders("Desktop"), "業務ポータル.lnk"))
  shortcut.TargetPath = edgePath
  shortcut.Arguments = Chr(34) & fso.BuildPath(installDir, "source\light\index.html") & Chr(34)
  shortcut.WorkingDirectory = fso.BuildPath(installDir, "source\light")
  shortcut.Description = "業務ポータル（ライト／ダーク切替対応）"
  shortcut.IconLocation = fso.BuildPath(installDir, "source\light\icon.ico") & ",0"
  shortcut.Save
  Err.Clear ' A locked down Desktop must not invalidate the installed application.
End If

answer = MsgBox("アプリを次の場所に導入しました：" & vbCrLf & installDir & vbCrLf & _
  "Microsoft Edgeで source\light\index.html を開いてください。" & vbCrLf & _
  "共同利用では、画面から既存の共有JSONを選択してください。本番用にサンプルを選ばないでください。" & vbCrLf & _
  "従来版Outlookとの連携を今すぐ設定しますか？", vbYesNo + vbQuestion, "業務ポータルの導入")
If answer = vbYes Then
  outlookInstaller = fso.BuildPath(installDir, "outlook\Install_Outlook_Sync.vbs")
  rc = shell.Run(Chr(34) & shell.ExpandEnvironmentStrings("%SystemRoot%") & _
    "\System32\wscript.exe" & Chr(34) & " " & Chr(34) & outlookInstaller & Chr(34), 1, True)
  If Err.Number <> 0 Then
    MsgBox "アプリは導入済みです。Outlookの設定を開始できませんでした：" & Err.Description, vbExclamation, "業務ポータルの導入"
    Err.Clear
  ElseIf rc <> 0 Then
    MsgBox "アプリは導入済みです。Outlookの設定はエラーコード " & rc & _
      " で終了しました。後から outlook\Install_Outlook_Sync.vbs を実行できます。", vbExclamation, "業務ポータルの導入"
  End If
End If
If oldMoved Then
  MsgBox "導入が完了しました。以前のアプリは次の場所に退避しました：" & vbCrLf & _
    backupDir & vbCrLf & "共有JSONは変更していません。", vbInformation, "業務ポータルの導入"
Else
  MsgBox "導入が完了しました。共有JSONは変更していません。", vbInformation, "業務ポータルの導入"
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
