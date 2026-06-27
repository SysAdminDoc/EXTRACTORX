"""7-Zip discovery and command construction."""

from __future__ import annotations

import hashlib
import logging
import os
import shutil
import subprocess
import urllib.error
import urllib.request
from pathlib import Path
from uuid import uuid4

from .config import app_data_dir


log = logging.getLogger("extractorx.sevenzip")

MIN_SEVENZIP_VERSION = (26, 2)
MIN_SEVENZIP_LABEL = "26.02"


def parse_7zip_version(exe: Path) -> tuple[int, int] | None:
    try:
        result = subprocess.run(
            [str(exe)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            check=False,
        )
        combined = result.stdout + result.stderr
        import re
        match = re.search(r"7-Zip.*?\b(\d+)\.(\d+)\b", combined)
        if match:
            return int(match.group(1)), int(match.group(2))
    except (OSError, subprocess.TimeoutExpired):
        pass
    return None


def check_7zip_version(exe: Path) -> bool:
    version = parse_7zip_version(exe)
    if version is None:
        log.warning("Could not determine 7-Zip version for %s", exe)
        return True
    if version < MIN_SEVENZIP_VERSION:
        log.warning(
            "7-Zip %d.%02d at %s is below minimum %s (CVE-2025-11001, CVE-2026-48095).",
            version[0], version[1], exe, MIN_SEVENZIP_LABEL,
        )
        return False
    log.info("7-Zip %d.%02d at %s", version[0], version[1], exe)
    return True


def find_7zip(override: str | Path | None = None) -> Path | None:
    if override:
        candidate = Path(str(override)).expanduser()
        if candidate.is_file():
            return candidate
    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "7-Zip" / "7z.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "7-Zip" / "7z.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "NanaZip" / "NanaZipC.exe",
    ]
    for name in ("7z", "7z.exe", "7zzs.exe", "NanaZipC.exe"):
        found = shutil.which(name)
        if found:
            candidates.append(Path(found))
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _atomic_download(url: str, target: Path, timeout: float = 120.0, expected_sha256: str | None = None) -> None:
    """Download ``url`` to a temp sibling then move into place.

    Avoids leaving a corrupt partial file at ``target`` if the transfer is
    interrupted mid-flight. When *expected_sha256* is given, the download is
    verified before committing; a mismatch raises ``ValueError``.
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(f"{target.name}.{uuid4().hex[:8]}.part")
    request = urllib.request.Request(url, headers={"User-Agent": "ExtractorX-Bootstrap/1.0"})
    try:
        sha = hashlib.sha256() if expected_sha256 else None
        with urllib.request.urlopen(request, timeout=timeout) as response, tmp.open("wb") as handle:  # type: ignore[arg-type]
            while True:
                chunk = response.read(65536)
                if not chunk:
                    break
                handle.write(chunk)
                if sha:
                    sha.update(chunk)
        if sha and expected_sha256:
            actual = sha.hexdigest().lower()
            if actual != expected_sha256.lower():
                raise ValueError(
                    f"SHA-256 mismatch for {url}: expected {expected_sha256[:16]}..., got {actual[:16]}..."
                )
        os.replace(tmp, target)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass


def download_7zip() -> Path:
    """Fetch a bootstrap copy of 7-Zip into the app-data folder.

    Pulls the official ``7zr.exe`` (self-contained, stable URL) and then tries
    to expand the ``7z*-extra`` package on top so the user gets the full
    ``7z.exe`` command-line. The extra archive step is best-effort -- if the
    download or unpack fails, callers still receive a working ``7zr.exe``.
    """
    target_dir = app_data_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    sevenzr = target_dir / "7zr.exe"
    sevenz = target_dir / "7z.exe"
    _atomic_download(
        "https://www.7-zip.org/a/7zr.exe", sevenzr,
        expected_sha256="56b8cc9f4971cef253644fafe54063ed7fdca551d4dee0f8c6baa81b855acd72",
    )
    try:
        shutil.copyfile(sevenzr, sevenz)
    except OSError as exc:
        log.warning("Could not copy 7zr.exe to 7z.exe: %s", exc)

    extra = target_dir / "7z-extra.7z"
    try:
        _atomic_download(
            "https://www.7-zip.org/a/7z2602-extra.7z", extra,
            expected_sha256="081df9e9311dfd9c9e0e98c1c80180b99bb51e4cb24156b5f3057fe3c259d70a",
        )
        subprocess_path = sevenz if sevenz.exists() else sevenzr
        subprocess.run(
            [str(subprocess_path), "x", str(extra), f"-o{target_dir}", "-y"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            timeout=300,
            check=False,
        )
    except subprocess.TimeoutExpired:
        log.warning("7z-extra unpack timed out (keeping 7zr only).")
    except (OSError, ValueError, urllib.error.URLError) as exc:
        log.warning("Could not fetch 7z-extra package: %s (keeping 7zr only).", exc)
    finally:
        try:
            extra.unlink(missing_ok=True)
        except OSError:
            pass
    return sevenz if sevenz.exists() else sevenzr


def list_archive_contents(
    sevenzip_path: Path | str | None,
    archive: Path | str,
    password: str | None = None,
) -> list[dict[str, str]]:
    """List the contents of *archive* using ``7z l`` without extracting.

    Returns a list of dicts with keys: ``Path``, ``Size``, ``Modified``, ``Attr``.
    """
    if not sevenzip_path:
        return []
    command: list[str] = [
        str(sevenzip_path),
        "l",
        str(archive),
        "-slt",
        "-y",
    ]
    if password:
        command.append(f"-p{password}")
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    entries: list[dict[str, str]] = []
    current: dict[str, str] = {}
    for line in result.stdout.splitlines():
        line = line.strip()
        if line.startswith("Path = ") and current.get("Path"):
            entries.append(current)
            current = {}
        if " = " in line:
            key, _, value = line.partition(" = ")
            key = key.strip()
            if key in ("Path", "Size", "Modified", "Attr", "Folder"):
                current[key] = value.strip()
    if current.get("Path"):
        entries.append(current)
    archive_name_str = Path(str(archive)).name
    return [
        e for e in entries
        if e.get("Folder", "") != "+" and e.get("Path", "") != archive_name_str
    ]


def get_archive_total_size(
    sevenzip_path: Path | str | None,
    archive: Path | str,
    password: str | None = None,
) -> int | None:
    """Return the total uncompressed size of *archive* in bytes, or None."""
    entries = list_archive_contents(sevenzip_path, archive, password)
    total = 0
    for entry in entries:
        try:
            total += int(entry.get("Size", "0") or "0")
        except (ValueError, TypeError):
            continue
    return total if total > 0 else None


def overwrite_switch(mode: str) -> str:
    if mode == "Never":
        return "-aos"
    if mode == "Rename":
        return "-aou"
    return "-aoa"


_ENCODING_CODE_PAGES = {
    "UTF-8": 65001,
    "cp437": 437,
    "cp932": 932,
    "cp936": 936,
    "cp949": 949,
    "cp950": 950,
    "cp1251": 1251,
    "cp1252": 1252,
}


def build_sevenzip_command(
    sevenzip_path: Path | str | None,
    archive: Path | str,
    output: Path | str,
    overwrite_mode: str = "Always",
    exclusions: str = "",
    inclusions: str = "",
    password: str | None = None,
    test_only: bool = False,
    filename_encoding: str = "Auto",
) -> list[str]:
    """Construct the 7-Zip invocation used by the extraction service.

    Extracted into its own helper so the command layout can be verified without
    touching the filesystem or spawning 7-Zip.
    """
    if not sevenzip_path:
        raise ValueError("7-Zip path is required but was not provided. Run find_7zip() first.")
    command: list[str] = [
        str(sevenzip_path),
        "t" if test_only else "x",
        str(archive),
    ]
    if not test_only:
        command.append(f"-o{output}")
    command.extend(["-y", overwrite_switch(overwrite_mode), "-bb1", "-bsp1", "-bso1", "-bse1"])
    code_page = _ENCODING_CODE_PAGES.get(filename_encoding)
    if code_page is not None:
        command.append(f"-mcp={code_page}")
    for pattern in [part.strip() for part in str(inclusions or "").split(";") if part.strip()]:
        command.append(f"-ir!{pattern}")
    for pattern in [part.strip() for part in str(exclusions or "").split(";") if part.strip()]:
        command.append(f"-xr!{pattern}")
    if password:
        command.append(f"-p{password}")
    return command
