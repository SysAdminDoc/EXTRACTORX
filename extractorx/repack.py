"""Archive format conversion (repack).

Extracts an archive to a temporary directory, then re-archives its contents
in a different format using 7-Zip. Supports ZIP, 7z, and TAR output.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable

log = logging.getLogger("extractorx.repack")

REPACK_FORMATS = {
    "zip": "zip",
    "7z": "7z",
    "tar": "tar",
}

LogCallback = Callable[[str, str], None]


def repack_archive(
    sevenzip_path: Path | str,
    source: Path,
    target_format: str,
    output_dir: Path | None = None,
    password: str | None = None,
    log_cb: LogCallback | None = None,
) -> Path | None:
    """Convert *source* to *target_format* and return the output path.

    The archive is extracted to a temp directory, then re-compressed.
    Returns the path to the new archive, or ``None`` on failure.
    """
    fmt = target_format.lower().lstrip(".")
    if fmt not in REPACK_FORMATS:
        if log_cb:
            log_cb(f"Unsupported target format: {target_format}", "error")
        return None

    output_dir = output_dir or source.parent
    output_name = source.stem
    for compound in (".tar.gz", ".tar.bz2", ".tar.xz", ".tar.zst"):
        if source.name.lower().endswith(compound):
            output_name = source.name[: -len(compound)]
            break
    output_path = output_dir / f"{output_name}.{fmt}"
    counter = 1
    while output_path.exists():
        output_path = output_dir / f"{output_name} ({counter}).{fmt}"
        counter += 1

    tmp = None
    try:
        tmp = Path(tempfile.mkdtemp(prefix="extractorx_repack_"))
        if log_cb:
            log_cb(f"Extracting {source.name} for repack...", "info")

        extract_cmd: list[str] = [
            str(sevenzip_path), "x", str(source), f"-o{tmp}", "-y",
        ]
        if password:
            extract_cmd.append(f"-p{password}")
        result = subprocess.run(
            extract_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3600,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            check=False,
        )
        if result.returncode not in {0, 1}:
            if log_cb:
                tail = (result.stdout or result.stderr or "").strip().splitlines()[-1:]
                log_cb(f"Extraction failed: {tail[0] if tail else f'exit code {result.returncode}'}", "error")
            return None

        if log_cb:
            log_cb(f"Re-archiving as {fmt}...", "info")

        compress_cmd: list[str] = [
            str(sevenzip_path), "a", str(output_path), os.path.join(str(tmp), "*"), f"-t{fmt}", "-y",
        ]
        result = subprocess.run(
            compress_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=3600,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            check=False,
        )
        if result.returncode not in {0, 1}:
            if log_cb:
                tail = (result.stdout or result.stderr or "").strip().splitlines()[-1:]
                log_cb(f"Compression failed: {tail[0] if tail else f'exit code {result.returncode}'}", "error")
            try:
                output_path.unlink(missing_ok=True)
            except OSError:
                pass
            return None

        if log_cb:
            log_cb(f"Repacked to {output_path.name}", "success")
        return output_path

    except (OSError, subprocess.TimeoutExpired) as exc:
        if log_cb:
            log_cb(f"Repack failed: {exc}", "error")
        return None
    finally:
        if tmp and tmp.exists():
            shutil.rmtree(tmp, ignore_errors=True)
