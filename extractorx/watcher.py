"""Polling watch-folder service without external dependencies."""

from __future__ import annotations

import os
from pathlib import Path
from queue import Queue
from threading import Event, Thread
from typing import Iterator

from .archive import is_non_first_volume, is_supported_archive


class WatchService:
    def __init__(self, folders: list[str], queue: Queue[Path], deep_detection: bool = True) -> None:
        self.folders = [Path(folder).expanduser() for folder in folders if str(folder).strip()]
        self.queue = queue
        self.deep_detection = deep_detection
        self._stop = Event()
        self._thread: Thread | None = None
        self._seen: dict[Path, int] = {}

    @property
    def active(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self) -> None:
        if self.active or not self.folders:
            return
        self._stop.clear()
        self._prime_seen()
        self._thread = Thread(target=self._run, name="ExtractorXWatchService", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=2)
        self._thread = None

    def _prime_seen(self) -> None:
        self._seen.clear()
        for folder in self.folders:
            if not folder.is_dir():
                continue
            for path in _walk_files(folder, self._stop):
                try:
                    key = path.resolve()
                    self._seen[key] = path.stat().st_size
                except OSError:
                    continue

    def _run(self) -> None:
        while not self._stop.is_set():
            observed: set[Path] = set()
            for folder in self.folders:
                if self._stop.is_set() or not folder.is_dir():
                    continue
                for path in _walk_files(folder, self._stop):
                    if self._stop.is_set():
                        break
                    if is_non_first_volume(path):
                        continue
                    try:
                        stat = path.stat()
                        key = path.resolve()
                    except OSError:
                        continue
                    observed.add(key)
                    last_size = self._seen.get(key)
                    if last_size == stat.st_size:
                        continue
                    self._seen[key] = stat.st_size
                    if is_supported_archive(path, self.deep_detection) and _wait_until_stable(path, self._stop):
                        self.queue.put(path)
            stale = [key for key in self._seen if key not in observed]
            for key in stale:
                self._seen.pop(key, None)
            self._stop.wait(3)


def _walk_files(folder: Path, should_stop: Event) -> Iterator[Path]:
    """Yield regular files under ``folder`` with per-directory error isolation.

    Skips directory symlinks (and Windows junctions) to avoid cycles, and
    swallows ``PermissionError`` / ``OSError`` from individual subfolders so
    the watcher never crashes on an unreadable directory.
    """
    visited: set[tuple[int, int] | str] = set()
    stack: list[Path] = [folder]
    while stack and not should_stop.is_set():
        current = stack.pop()
        try:
            info = current.stat()
            key: tuple[int, int] | str = (info.st_dev, info.st_ino)
        except OSError:
            key = str(current)
        if key in visited:
            continue
        visited.add(key)
        try:
            with os.scandir(current) as entries:
                for entry in entries:
                    if should_stop.is_set():
                        return
                    try:
                        if entry.is_symlink():
                            continue
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(Path(entry.path))
                            continue
                        if entry.is_file(follow_symlinks=False):
                            yield Path(entry.path)
                    except OSError:
                        continue
        except (OSError, PermissionError):
            continue


def _wait_until_stable(path: Path, should_stop: Event) -> bool:
    previous = -1
    for _ in range(10):
        if should_stop.is_set():
            return False
        try:
            size = path.stat().st_size
            with path.open("rb"):
                pass
        except OSError:
            should_stop.wait(0.5)
            continue
        if size == previous:
            return True
        previous = size
        should_stop.wait(0.5)
    return False
