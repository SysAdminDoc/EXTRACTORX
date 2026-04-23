"""Shared data models for ExtractorX."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from time import time
from uuid import uuid4


class QueueStatus(str, Enum):
    QUEUED = "Queued"
    SCANNING = "Scanning"
    TESTING = "Testing"
    TEST_OK = "Test OK"
    EXTRACTING = "Extracting"
    DONE = "Done"
    FAILED = "Failed"
    SKIPPED = "Skipped"
    PASSWORD_REQUIRED = "Password Required"


@dataclass
class QueueItem:
    archive_path: Path
    output_override: str | None = None
    id: str = field(default_factory=lambda: uuid4().hex)
    status: QueueStatus = QueueStatus.QUEUED
    output_path: Path | None = None
    size_bytes: int = 0
    error: str = ""
    created_at: float = field(default_factory=time)

    @classmethod
    def from_path(cls, path: Path, output_override: str | None = None) -> "QueueItem":
        resolved = path.expanduser().resolve()
        size = resolved.stat().st_size if resolved.exists() else 0
        return cls(archive_path=resolved, output_override=output_override, size_bytes=size)


@dataclass(frozen=True)
class OperationMessage:
    type: str
    text: str = ""
    item_id: str | None = None
    level: str = "info"
    payload: dict[str, object] = field(default_factory=dict)
