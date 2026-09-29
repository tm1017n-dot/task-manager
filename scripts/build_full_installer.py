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
VERSIONED_OUTPUT = ROOT / "release" / "業務ポータル_一括インストーラー_20260929_r4.vbs"
PREVIOUS_OUTPUTS = [ROOT / "release" / name for name in (
    "業務ポータル_一括インストーラー_20260929.vbs",
    "業務ポータル_一括インストーラー_20260929_r2.vbs",
    "業務ポータル_一括インストーラー_20260929_r3.vbs",
)]
APP = ROOT / "source" / "app"
SOURCES = [
    *(p for p in sorted(APP.rglob("*")) if p.is_file()),
    *(p for p in sorted((ROOT / "outlook").iterdir()) if p.is_file()),
]


HEADER = r'''Option Explicit
' Standalone installer for the latest HTML portal. Generated: do not edit by hand.
' The shared JSON and browser-local journal are never copied or overwritten.
Dim fso, shell, stage, installDir, backupDir, installed, backed, answer
Dim data, fileList, rel, outlookInstaller, rc, edgePath, launcherNote
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
installDir = fso.GetParentFolderName(WScript.ScriptFullName)
If MsgBox("業務ポータル一括インストーラー（2026年9月29日・第4版）" & vbCrLf & vbCrLf & _
  "このVBSと同じフォルダーに業務ポータルを導入しますか？" & vbCrLf & _
  "導入先：" & installDir & vbCrLf & _
  "同名の既存ファイルは別フォルダーに退避し、共有JSONは変更しません。" & vbCrLf & _
  "別の場所で使っていた端末内データは自動移行されません。必要なら旧画面でバックアップしてください。", _
  vbYesNo + vbQuestion, "業務ポータルの導入") <> vbYes Then WScript.Quit 0
On Error Resume Next
stage = fso.BuildPath(installDir, "業務ポータル_準備_" & fso.GetTempName())
backupDir = fso.BuildPath(installDir, "業務ポータル_旧版_" & fso.GetTempName())
Set installed = CreateObject("Scripting.Dictionary")
Set backed = CreateObject("Scripting.Dictionary")
fileList = ""
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
  Dim name, oldPath, newPath
  On Error Resume Next
  For Each name In installed.Keys
    newPath = fso.BuildPath(installDir, name)
    If fso.FileExists(newPath) Then fso.DeleteFile newPath, True
  Next
  For Each name In backed.Keys
    oldPath = fso.BuildPath(backupDir, name)
    newPath = fso.BuildPath(installDir, name)
    If fso.FileExists(oldPath) And Not fso.FileExists(newPath) Then fso.MoveFile oldPath, newPath
  Next
  If stage <> "" Then If fso.FolderExists(stage) Then fso.DeleteFolder stage, True
  MsgBox "導入を中断しました：" & detail & vbCrLf & _
    "退避済みのファイルがある場合、保存先：" & backupDir, vbCritical, "業務ポータルの導入"
  WScript.Quit 1
End Sub

Sub InstallFile(relativePath)
  Dim fromPath, toPath, oldPath
  On Error Resume Next
  fromPath = fso.BuildPath(stage, relativePath)
  toPath = fso.BuildPath(installDir, relativePath)
  EnsureFolder fso.GetParentFolderName(toPath)
  If fso.FileExists(toPath) Then
    oldPath = fso.BuildPath(backupDir, relativePath)
    EnsureFolder fso.GetParentFolderName(oldPath)
    fso.MoveFile toPath, oldPath
    Check "以前のファイルの退避：" & relativePath
    backed.Add relativePath, True
    Check "退避の記録：" & relativePath
  End If
  fso.MoveFile fromPath, toPath
  Check "新しいファイルの配置：" & relativePath
  installed.Add relativePath, True
  Check "配置の記録：" & relativePath
End Sub

Sub CreatePortalShortcut()
  Dim name, target, oldPath, link, htmlPath, iconPath
  On Error Resume Next
  name = "業務ポータル.lnk"
  target = fso.BuildPath(installDir, name)
  htmlPath = fso.BuildPath(installDir, "index.html")
  iconPath = fso.BuildPath(installDir, "icon.ico")
  If fso.FileExists(target) Then
    oldPath = fso.BuildPath(backupDir, name)
    EnsureFolder fso.GetParentFolderName(oldPath)
    fso.MoveFile target, oldPath
    Check "以前のショートカットの退避"
    backed.Add name, True
    Check "ショートカットの退避記録"
  End If
  installed.Add name, True
  Check "ショートカットの配置記録"
  edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
  If Not fso.FileExists(edgePath) Then edgePath = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
  If Not fso.FileExists(edgePath) Then edgePath = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\Edge\Application\msedge.exe"
  Set link = shell.CreateShortcut(target)
  Check "ショートカットの作成"
  If fso.FileExists(edgePath) Then
    link.TargetPath = edgePath
    link.Arguments = Chr(34) & htmlPath & Chr(34)
    launcherNote = "ショートカットはMicrosoft Edgeで開きます。"
  Else
    link.TargetPath = htmlPath
    link.Arguments = ""
    launcherNote = "Edgeを検出できなかったため、ショートカットは既定のブラウザーで開きます。共同利用ではEdgeを使ってください。"
  End If
  link.WorkingDirectory = installDir
  link.Description = "業務ポータル（ライト／ダーク切替対応）"
  link.IconLocation = iconPath & ",0"
  Check "ショートカットの設定"
  link.Save
  Check "ショートカットの保存"
  If Not fso.FileExists(target) Then AbortInstall "ショートカットを保存できませんでした。"
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
For Each rel In Split(fileList, vbLf)
  If rel <> "" Then InstallFile rel
Next
CreatePortalShortcut
fso.DeleteFolder stage, True
Check "一時ファイルの片付け"
stage = ""

answer = MsgBox("アプリを次の場所に導入しました：" & vbCrLf & installDir & vbCrLf & _
  "アイコン付きの「業務ポータル」ショートカットも同じフォルダーに作成しました。" & vbCrLf & _
  launcherNote & vbCrLf & _
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
If backed.Count > 0 Then
  MsgBox "導入が完了しました。同名の旧ファイルは次の場所に退避しました：" & vbCrLf & _
    backupDir & vbCrLf & "共有JSONは変更していません。", vbInformation, "業務ポータルの導入"
Else
  MsgBox "導入が完了しました。共有JSONは変更していません。", vbInformation, "業務ポータルの導入"
End If
'''


def main() -> None:
    lines = HEADER.splitlines()
    for path in SOURCES:
        rel = (path.relative_to(APP) if path.is_relative_to(APP) else path.relative_to(ROOT)).as_posix().replace("/", "\\")
        payload = base64.b64encode(path.read_bytes()).decode("ascii")
        lines.extend(["", "data = \"\""])
        for offset in range(0, len(payload), 800):
            lines.append(f'data = data & "{payload[offset:offset + 800]}"')
        lines.append(f'Call SaveEmbedded("{rel}", {path.stat().st_size}, data)')
        lines.append(f'fileList = fileList & "{rel}" & vbLf')
    lines.extend(FOOTER.splitlines())
    installer = ("\ufeff" + "\r\n".join(lines) + "\r\n").encode("utf-16le")
    OUTPUT.write_bytes(installer)
    VERSIONED_OUTPUT.write_bytes(installer)
    sums = ROOT / "checksums" / "SHA256SUMS.txt"
    old = [line for line in sums.read_text().splitlines()
           if not any(line.endswith("  " + path.name) for path in (OUTPUT, VERSIONED_OUTPUT, *PREVIOUS_OUTPUTS))]
    digest = hashlib.sha256(installer).hexdigest().upper()
    old.extend(f"{digest}  {path.name}" for path in (OUTPUT, VERSIONED_OUTPUT))
    sums.write_text("\n".join(old) + "\n")
    print(f"{OUTPUT}: {len(SOURCES)} files, {OUTPUT.stat().st_size} bytes")


if __name__ == "__main__":
    main()
