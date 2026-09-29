"""Check that every file embedded in the standalone VBS is byte-for-byte current."""
from __future__ import annotations

import base64
import hashlib
import re
import unittest
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
INSTALLER = ROOT / "release" / "業務ポータル_一括インストーラー_latest.vbs"
VERSIONED = ROOT / "release" / "業務ポータル_一括インストーラー_20260929_r4.vbs"
BLOCK = re.compile(
    r'data = ""\n((?:data = data & "[A-Za-z0-9+/=]+"\n)+)'
    r'Call SaveEmbedded\("([^"]+)", (\d+), data\)',
)


class FullInstallerTest(unittest.TestCase):
    def test_versioned_download_is_identical_and_has_clear_install_location(self) -> None:
        self.assertEqual(VERSIONED.read_bytes(), INSTALLER.read_bytes())
        source = VERSIONED.read_text(encoding="utf-16")
        self.assertIn("2026年9月29日・第4版", source)
        self.assertIn('installDir = fso.GetParentFolderName(WScript.ScriptFullName)', source)
        self.assertIn('"導入先：" & installDir', source)
        expected = hashlib.sha256(VERSIONED.read_bytes()).hexdigest().upper()
        self.assertIn(f"{expected}  {VERSIONED.name}",
                      (ROOT / "checksums" / "SHA256SUMS.txt").read_text())

    def test_localized_vbs_dialogs(self) -> None:
        for file in (INSTALLER, ROOT / "outlook/Install_Outlook_Sync.vbs",
                     ROOT / "outlook/Uninstall_Outlook_Sync.vbs"):
            raw = file.read_bytes()
            self.assertTrue(raw.startswith(b"\xff\xfe"), str(file))
            source = raw.decode("utf-16")
            self.assertIn("業務ポータル", source)
            self.assertNotIn('"Work Portal setup"', source)

    def test_payloads_match_sources(self) -> None:
        raw = INSTALLER.read_bytes()
        self.assertTrue(raw.startswith(b"\xff\xfe"))
        source = raw.decode("utf-16").replace("\r\n", "\n")
        found: set[str] = set()
        for match in BLOCK.finditer(source):
            lines, name, expected_size = match.groups()
            chunks = re.findall(r'data = data & "([A-Za-z0-9+/=]+)"', lines)
            self.assertTrue(all(len(chunk) <= 800 for chunk in chunks))
            payload = base64.b64decode("".join(chunks), validate=True)
            path = (ROOT / "outlook" / Path(name.replace("\\", "/")).name
                    if name.startswith("outlook\\") else
                    ROOT / "source" / "app" / Path(name.replace("\\", "/")))
            self.assertEqual(payload, path.read_bytes(), name)
            self.assertEqual(len(payload), int(expected_size), name)
            self.assertNotIn(name, found)
            found.add(name)
        source_paths = {
            str(path.relative_to(ROOT / "source" / "app")).replace("/", "\\")
            for path in (ROOT / "source" / "app").rglob("*") if path.is_file()
        } | {
            str(path.relative_to(ROOT)).replace("/", "\\")
            for path in (ROOT / "outlook").rglob("*") if path.is_file()
        }
        self.assertEqual(found, source_paths)
        self.assertIn("index.html", found)
        self.assertEqual(sum(name.endswith("index.html") for name in found), 1)
        self.assertIn("installDir = fso.GetParentFolderName(WScript.ScriptFullName)", source)
        self.assertNotIn('installDir = fso.BuildPath(localBase, "WorkPortal")', source)
        self.assertIn('Then InstallFile rel', source)

    def test_checksum(self) -> None:
        expected = hashlib.sha256(INSTALLER.read_bytes()).hexdigest().upper()
        text = (ROOT / "checksums" / "SHA256SUMS.txt").read_text()
        self.assertIn(f"{expected}  {INSTALLER.name}", text.splitlines())

    def test_shortcut_is_adjacent_and_uses_bundled_icon(self) -> None:
        source = INSTALLER.read_text(encoding="utf-16")
        self.assertIn('name = "業務ポータル.lnk"', source)
        self.assertIn('target = fso.BuildPath(installDir, name)', source)
        self.assertIn('iconPath = fso.BuildPath(installDir, "icon.ico")', source)
        self.assertIn('link.IconLocation = iconPath & ",0"', source)
        self.assertIn('link.TargetPath = edgePath', source)
        self.assertIn('link.TargetPath = htmlPath', source)
        self.assertIn('If rel <> "" Then InstallFile rel\nNext\nCreatePortalShortcut', source.replace("\r\n", "\n"))
        self.assertNotIn('shell.SpecialFolders("Desktop")', source)

    def test_full_zip_contains_one_root_html(self) -> None:
        with ZipFile(ROOT / "release/業務ポータル_latest_HTML_導入一式.zip") as archive:
            names = archive.namelist()
            self.assertIsNone(archive.testzip())
            self.assertEqual([n for n in names if n.endswith("/index.html")],
                             ["業務ポータル/index.html"])
            self.assertIn("業務ポータル/outlook/Sync_Outlook.js", names)


if __name__ == "__main__":
    unittest.main()
