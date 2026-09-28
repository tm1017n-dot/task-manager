"""Build the additive Outlook installer ZIP and update the latest-source manifests."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    "README.md",
    "Install_Outlook_Sync.vbs",
    "Sync_Outlook.js",
    "Uninstall_Outlook_Sync.vbs",
)
ARCHIVE = ROOT / "release" / "Outlook連携インストーラー_latest.zip"
FULL_ARCHIVE = ROOT / "release" / "業務ポータル_latest_HTML_導入一式.zip"


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    for theme in ("light", "dark"):
        base = ROOT / "source" / theme
        path = base / "manifest.json"
        manifest = json.loads(path.read_text())
        manifest["generatedAt"] = "2026-09-28"
        for name in manifest["files"]:
            item = base / name
            manifest["files"][name] = {"bytes": item.stat().st_size, "sha256": digest(item)}
        path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")

    with ZipFile(ARCHIVE, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for name in FILES:
            info = ZipInfo("Outlook連携/" + name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, (ROOT / "outlook" / name).read_bytes())

    with ZipFile(FULL_ARCHIVE, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for theme in ("light", "dark"):
            for path in sorted((ROOT / "source" / theme).rglob("*")):
                if path.is_file():
                    archive.write(path, "業務ポータル/" + str(path.relative_to(ROOT)))
        for name in FILES:
            archive.write(ROOT / "outlook" / name, "業務ポータル/outlook/" + name)

    checksum_file = ROOT / "checksums" / "SHA256SUMS.txt"
    lines = [line for line in checksum_file.read_text().splitlines()
             if line.strip() and ARCHIVE.name not in line and FULL_ARCHIVE.name not in line]
    lines.append(f"{digest(ARCHIVE).upper()}  {ARCHIVE.name}")
    lines.append(f"{digest(FULL_ARCHIVE).upper()}  {FULL_ARCHIVE.name}")
    checksum_file.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
