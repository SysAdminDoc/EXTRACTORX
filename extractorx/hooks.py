"""Lifecycle hook dispatch.

Runs user-configured shell commands for pre-extract, post-extract, and
on-failure events. Commands use the same token syntax as external processors.
Command strings are split into argument lists for safe subprocess execution.
"""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Callable

from .postprocess import expand_processor_command


LogCallback = Callable[[str, str], None]


HOOK_KEYS = ("PreExtractCommand", "PostExtractCommand", "OnFailureCommand")


def run_hook(
    hook_name: str,
    config: dict,
    archive: Path,
    output: Path,
    log: LogCallback,
    exit_code: int = 0,
) -> None:
    """Run the configured command for ``hook_name`` if one is set.

    ``hook_name`` must be one of the keys in :data:`HOOK_KEYS`. Errors are
    surfaced through ``log`` and never propagated -- a broken hook should not
    take the extraction run down with it.
    """
    if hook_name not in HOOK_KEYS:
        return
    template = str(config.get(hook_name, "") or "").strip()
    if not template:
        return
    template = template.replace("{ExitCode}", str(exit_code))
    command = expand_processor_command(template, archive=archive, output=output)
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            timeout=1800,
            check=False,
        )
        if result.returncode == 0:
            log(f"{hook_name} completed for {archive.name}", "info")
        else:
            tail = (result.stderr or result.stdout or "").strip().splitlines()[-1:] or [
                f"exit code {result.returncode}"
            ]
            log(f"{hook_name} failed for {archive.name}: {tail[0]}", "warning")
    except (OSError, subprocess.TimeoutExpired) as exc:
        log(f"{hook_name} failed for {archive.name}: {exc}", "warning")
