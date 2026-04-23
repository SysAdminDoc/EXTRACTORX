"""Download helper for URL-based startup arguments.

Only the standard library is used so the Python port picks up no extra
runtime dependencies. Downloads go into ``app_data_dir()/downloads``.
"""

from __future__ import annotations

import re
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Callable
from uuid import uuid4

from .config import app_data_dir


LogCallback = Callable[[str, str], None]


# Cap individual downloads to 4 GiB. Larger archives are rare and giving the
# user an early error is preferable to silently filling a disk while they wait.
DEFAULT_MAX_BYTES = 4 * 1024 * 1024 * 1024
DEFAULT_TIMEOUT_SECONDS = 120
USER_AGENT = "ExtractorX/2.x (+https://github.com/SysAdminDoc/ExtractorX)"
_SAFE_NAME = re.compile(r"[^A-Za-z0-9._\- ()\[\]]+")


def download_archive(
    url: str,
    log: LogCallback | None = None,
    *,
    max_bytes: int = DEFAULT_MAX_BYTES,
    timeout: float = DEFAULT_TIMEOUT_SECONDS,
) -> Path | None:
    """Download ``url`` into the app-data downloads folder.

    Returns the path on success, ``None`` on failure. ``log`` is optional and
    invoked with ``(message, level)`` -- the caller controls surfacing. The
    download is streamed in chunks and aborts if the byte count exceeds
    ``max_bytes`` to prevent a runaway URL from filling the disk.
    """
    parsed = urllib.parse.urlsplit(url)
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        if log:
            log(f"Refusing to download unsupported URL: {url!r}", "warning")
        return None
    download_dir = app_data_dir() / "downloads"
    try:
        download_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        if log:
            log(f"Could not create download folder: {exc}", "error")
        return None
    raw_name = Path(urllib.parse.unquote(parsed.path)).name
    safe_name = _SAFE_NAME.sub("_", raw_name).strip() if raw_name else ""
    if not safe_name:
        safe_name = f"download-{uuid4().hex[:8]}"
    target = _unique(download_dir / safe_name)
    if log:
        log(f"Downloading {url}", "info")
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    bytes_written = 0
    try:
        # noqa: S310 - user-supplied URL is expected; scheme is validated above.
        with urllib.request.urlopen(request, timeout=timeout) as response:  # type: ignore[arg-type]
            try:
                declared = int(response.headers.get("Content-Length", "") or 0)
            except (TypeError, ValueError):
                declared = 0
            if declared and declared > max_bytes:
                if log:
                    log(
                        f"Download aborted: advertised size {declared} bytes exceeds limit {max_bytes}.",
                        "error",
                    )
                return None
            with target.open("wb") as handle:
                while True:
                    chunk = response.read(65536)
                    if not chunk:
                        break
                    bytes_written += len(chunk)
                    if bytes_written > max_bytes:
                        if log:
                            log(
                                f"Download aborted: exceeded {max_bytes} bytes after streaming {bytes_written}.",
                                "error",
                            )
                        raise _DownloadTooLarge()
                    handle.write(chunk)
    except _DownloadTooLarge:
        _best_effort_delete(target)
        return None
    except (OSError, ValueError) as exc:
        if log:
            log(f"Download failed: {exc}", "error")
        _best_effort_delete(target)
        return None
    if log:
        log(f"Downloaded {bytes_written} byte(s) to {target}", "success")
    return target


class _DownloadTooLarge(Exception):
    """Internal marker for the max-bytes early abort."""


def _best_effort_delete(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _unique(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = path.with_name(f"{stem} ({counter}){suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def looks_like_url(value: str) -> bool:
    parsed = urllib.parse.urlsplit(value)
    return parsed.scheme in ("http", "https") and bool(parsed.netloc)
