"""Headless daemon mode for watch-folder extraction without a GUI."""

from __future__ import annotations

import logging
import signal
import sys
from pathlib import Path
from queue import Empty, Queue
from threading import Event

from .config import load_config
from .extractor import ExtractionService
from .models import OperationMessage, QueueItem, QueueStatus
from .passwords import PasswordStore
from .sevenzip import find_7zip
from .watcher import WatchService


log = logging.getLogger("extractorx.daemon")


def run_daemon() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    config = load_config()
    sevenzip_path = find_7zip(override=config.get("SevenZipOverride"))
    if not sevenzip_path:
        log.error("7-Zip not found. Install 7-Zip or set SevenZipOverride in config.json.")
        return 1

    password_store = PasswordStore()
    passwords = password_store.load()
    messages: Queue[OperationMessage] = Queue()
    watch_queue: Queue[Path] = Queue()
    stop = Event()

    folders = [str(f) for f in config.get("WatchFolders", []) if str(f).strip()]
    if not folders:
        log.error("No watch folders configured. Add folders in config.json WatchFolders.")
        return 1

    log.info("Starting ExtractorX daemon, watching %d folder(s)", len(folders))
    for folder in folders:
        log.info("  %s", folder)

    watcher = WatchService(folders, watch_queue, bool(config.get("DeepArchiveDetection", True)))
    watcher.start()

    service = ExtractionService(config, sevenzip_path, passwords, messages)

    def handle_signal(_signum: int, _frame: object) -> None:
        log.info("Shutdown requested")
        stop.set()

    signal.signal(signal.SIGINT, handle_signal)
    signal.signal(signal.SIGTERM, handle_signal)

    pending: list[QueueItem] = []
    try:
        while not stop.is_set():
            while True:
                try:
                    path = watch_queue.get_nowait()
                except Empty:
                    break
                if path.exists():
                    pending.append(QueueItem.from_path(path))
                    log.info("Detected: %s", path.name)

            if pending and not service.active:
                batch = list(pending)
                pending.clear()
                service.extract_items(batch, test_only=False)

            while True:
                try:
                    msg = messages.get_nowait()
                except Empty:
                    break
                if msg.text:
                    level = msg.level or "info"
                    getattr(log, level if level in ("info", "warning", "error") else "info", log.info)(msg.text)

            stop.wait(1)
    finally:
        watcher.stop()
        if service.active:
            service.stop()
        log.info("ExtractorX daemon stopped")

    return 0
