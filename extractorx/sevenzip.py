"""7-Zip discovery and command construction."""

from __future__ import annotations

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


def _atomic_download(url: str, target: Path, timeout: float = 120.0) -> None:
    """Download ``url`` to a temp sibling then move into place.

    Avoids leaving a corrupt partial file at ``target`` if the transfer is
    interrupted mid-flight.
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(f"{target.name}.{uuid4().hex[:8]}.part")
    request = urllib.request.Request(url, headers={"User-Agent": "ExtractorX-Bootstrap/1.0"})
    try:
        # noqa: S310 - the URL is a fixed https endpoint at 7-zip.org.
        with urllib.request.urlopen(request, timeout=timeout) as response, tmp.open("wb") as handle:  # type: ignore[arg-type]
            while True:
                chunk = response.read(65536)
                if not chunk:
                    break
                handle.write(chunk)
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
    _atomic_download("https://www.7-zip.org/a/7zr.exe", sevenzr)
    try:
        shutil.copyfile(sevenzr, sevenz)
    except OSError as exc:
        log.warning("Could not copy 7zr.exe to 7z.exe: %s", exc)

    extra = target_dir / "7z-extra.7z"
    try:
        _atomic_download("https://www.7-zip.org/a/7z2408-extra.7z", extra)
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
    command: list[str] = [
        str(sevenzip_path) if sevenzip_path else "7z",
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
