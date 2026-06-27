"""Plugin SDK for custom post-processors.

Plugins are Python modules placed in the ``plugins/`` subdirectory of the
app-data folder (``%APPDATA%/ExtractorX/plugins/`` or the portable
equivalent). Each module must define a ``process(archive_path, output_path,
config)`` function that is called after successful extraction.

Plugins are loaded on demand and errors in individual plugins do not block
extraction. Plugins only execute if their SHA-256 hash appears in the
``PluginAllowlist`` config list. New or changed plugins are logged with
their hash for easy allowlisting.
"""

from __future__ import annotations

import hashlib
import importlib.util
import logging
from pathlib import Path
from typing import Any, Callable

from .config import app_data_dir


log = logging.getLogger("extractorx.plugins")

LogCallback = Callable[[str, str], None]


def plugins_dir() -> Path:
    return app_data_dir() / "plugins"


def _file_hash(path: Path) -> str:
    sha = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(65536)
            if not chunk:
                break
            sha.update(chunk)
    return sha.hexdigest()


def discover_plugins() -> list[dict[str, Any]]:
    """Find all ``.py`` files in the plugins directory.

    Returns a list of dicts with ``name``, ``path``, and ``hash`` keys.
    """
    directory = plugins_dir()
    if not directory.is_dir():
        return []
    results: list[dict[str, Any]] = []
    for path in sorted(directory.glob("*.py")):
        if path.name.startswith("_"):
            continue
        try:
            file_hash = _file_hash(path)
        except OSError:
            continue
        results.append({"name": path.stem, "path": path, "hash": file_hash})
    return results


def run_plugins(
    archive: Path,
    output: Path,
    config: dict[str, Any],
    log_cb: LogCallback | None = None,
) -> None:
    """Execute all discovered plugins after a successful extraction."""
    plugins = discover_plugins()
    if not plugins:
        return
    allowlist = set(str(h).lower() for h in (config.get("PluginAllowlist", []) or []))
    for plugin_info in plugins:
        name = plugin_info["name"]
        path = plugin_info["path"]
        file_hash = plugin_info["hash"]
        if file_hash.lower() not in allowlist:
            if log_cb:
                log_cb(
                    f"Plugin {name} skipped (not in PluginAllowlist). Hash: {file_hash}",
                    "warning",
                )
            continue
        try:
            spec = importlib.util.spec_from_file_location(f"extractorx_plugin_{name}", path)
            if spec is None or spec.loader is None:
                continue
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            process_fn = getattr(module, "process", None)
            if not callable(process_fn):
                if log_cb:
                    log_cb(f"Plugin {name}: no process() function found, skipped.", "warning")
                continue
            process_fn(archive_path=archive, output_path=output, config=config)
            if log_cb:
                log_cb(f"Plugin {name}: completed for {archive.name}", "info")
        except Exception as exc:
            log.exception("Plugin %s failed", name)
            if log_cb:
                log_cb(f"Plugin {name} failed: {exc}", "warning")
