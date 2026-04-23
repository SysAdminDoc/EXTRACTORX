"""Folder scanning for archives."""

from __future__ import annotations

import os
from pathlib import Path
from threading import Event
from typing import Callable, Iterable

from .archive import is_non_first_volume, is_supported_archive


def scan_paths(
    roots: Iterable[Path],
    deep_detection: bool,
    should_stop: Event,
    on_path: Callable[[Path], None],
    on_log: Callable[[str, str], None] | None = None,
) -> int:
    """Walk ``roots`` and call ``on_path`` for every supported archive.

    Unlike a plain ``Path.rglob`` this uses ``os.scandir`` with per-directory
    error handling so one unreadable folder -- denied permissions, junction
    loops, disconnected network shares -- can't kill the whole scan. We also
    skip directory symlinks/junctions to avoid cycles on Windows.
    """
    count = 0
    visited: set[tuple[int, int] | str] = set()

    def report(message: str, level: str = "warning") -> None:
        if on_log:
            on_log(message, level)

    def walk(directory: Path) -> None:
        nonlocal count
        if should_stop.is_set():
            return
        try:
            key: tuple[int, int] | str
            try:
                stat_info = directory.stat()
                key = (stat_info.st_dev, stat_info.st_ino)
            except OSError:
                key = str(directory)
            if key in visited:
                return
            visited.add(key)
            with os.scandir(directory) as entries:
                for entry in entries:
                    if should_stop.is_set():
                        return
                    try:
                        is_symlink = entry.is_symlink()
                        is_dir = entry.is_dir(follow_symlinks=False)
                        is_file = entry.is_file(follow_symlinks=False)
                    except OSError as exc:
                        report(f"Skipped {entry.path}: {exc}")
                        continue
                    if is_dir:
                        if is_symlink:
                            # Avoid cycles via symlinked directories / junctions.
                            continue
                        walk(Path(entry.path))
                        continue
                    if not is_file:
                        continue
                    path = Path(entry.path)
                    try:
                        if is_non_first_volume(path):
                            continue
                        if is_supported_archive(path, deep_detection=deep_detection):
                            on_path(path)
                            count += 1
                    except OSError as exc:
                        report(f"Skipped {path}: {exc}")
        except (OSError, PermissionError) as exc:
            report(f"Skipped {directory}: {exc}")

    for root in roots:
        if should_stop.is_set():
            break
        if root.is_file():
            try:
                if is_non_first_volume(root):
                    continue
                if is_supported_archive(root, deep_detection=deep_detection):
                    on_path(root)
                    count += 1
            except OSError as exc:
                report(f"Skipped {root}: {exc}")
            continue
        if root.is_dir():
            walk(root)
    return count
