"""Check that every file embedded in the standalone VBS is byte-for-byte current."""
from __future__ import annotations

import base64
import hashlib
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INSTALLER = ROOT / "release" / "業務ポータル_一括インストーラー_latest.vbs"
BLOCK = re.compile(
    r'data = ""\n((?:data = data & "[A-Za-z0-9+/=]+"\n)+)'
    r'Call SaveEmbedded\("([^"]+)", (\d+), data\)',
)


class FullInstallerTest(unittest.TestCase):
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
            path = ROOT / Path(name.replace("\\", "/"))
            self.assertEqual(payload, path.read_bytes(), name)
            self.assertEqual(len(payload), int(expected_size), name)
            self.assertNotIn(name, found)
            found.add(name)
        source_paths = {
            str(path.relative_to(ROOT)).replace("/", "\\")
            for folder in (ROOT / "source", ROOT / "outlook")
            for path in folder.rglob("*") if path.is_file()
        }
        self.assertEqual(found, source_paths)

    def test_checksum(self) -> None:
        expected = hashlib.sha256(INSTALLER.read_bytes()).hexdigest().upper()
        text = (ROOT / "checksums" / "SHA256SUMS.txt").read_text()
        self.assertIn(f"{expected}  {INSTALLER.name}", text.splitlines())


if __name__ == "__main__":
    unittest.main()
