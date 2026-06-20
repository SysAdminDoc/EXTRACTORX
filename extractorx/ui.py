"""Tkinter user interface for the Python port."""

from __future__ import annotations

import tkinter as tk
import fnmatch
import os
import re
import subprocess
import sys
import winsound
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import __version__
from .archive import format_size, is_supported_archive, resolve_output_path
from .config import (
    DEFAULT_CONFIG,
    FILENAME_ENCODINGS,
    SMART_EXTRACT_MODES,
    THEMES as THEME_NAMES,
)
from .config import app_data_dir
from .config import save_config
from .download import looks_like_url
from .extractor import ExtractionService
from .identify import identify
from .models import OperationMessage, QueueItem, QueueStatus
from .passwords import PasswordStore, classify_entropy
from . import shell_integration
from .themes import get_theme
from .watcher import WatchService
from .windows_integration import WindowsShellBridge


LOG_LINE_CAP = 2000


class ExtractorXApp:
    def __init__(
        self,
        config: dict,
        sevenzip_path: Path | None,
        password_store: PasswordStore,
        startup_items: list[QueueItem] | None = None,
        startup_scan_paths: list[Path] | None = None,
        auto_extract_startup: bool = False,
        test_startup: bool = False,
        start_minimized: bool = False,
        portable: bool = False,
    ) -> None:
        self.config = config
        self.sevenzip_path = sevenzip_path
        self.password_store = password_store
        self.portable = portable
        self.passwords = password_store.load()
        self.messages: Queue[OperationMessage] = Queue()
        self.watch_queue: Queue[Path] = Queue()
        self.items: dict[str, QueueItem] = {}
        self.startup_scan_paths = startup_scan_paths or []
        self.auto_extract_startup = auto_extract_startup
        self.test_startup = test_startup
        self.sort_column = ""
        self.sort_reverse = False
        self.password_retry_candidates: list[QueueItem] = []
        self.service = ExtractionService(config, sevenzip_path, self.passwords, self.messages)
        self.watch_service: WatchService | None = None
        self.shell_bridge: WindowsShellBridge | None = None
        self.closing = False

        self.root = tk.Tk()
        title = f"ExtractorX v{__version__}"
        if self.portable:
            title += " (portable)"
        self.root.title(title)
        self.root.geometry(f"{config['WindowWidth']}x{config['WindowHeight']}")
        left = int(config.get("WindowLeft", -1))
        top = int(config.get("WindowTop", -1))
        if left >= 0 and top >= 0:
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            if left + 120 < screen_w and top + 80 < screen_h:
                self.root.geometry(f"+{left}+{top}")
        self.root.minsize(850, 500)
        self.root.attributes("-topmost", bool(config.get("AlwaysOnTop", False)))
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.palette = get_theme(config.get("Theme"))
        self.style = ttk.Style(self.root)
        self._configure_style()
        self._build_ui()
        self.shell_bridge = WindowsShellBridge(
            self.root,
            on_files_dropped=self._handle_dropped_paths,
            on_tray_restore=self._restore_from_tray,
            on_tray_menu=self._show_tray_menu,
        )
        self.shell_bridge.enable_drag_drop()
        self.root.bind("<Unmap>", self._handle_unmap)
        self._start_watchers()
        for item in startup_items or []:
            self.add_item(item)
        self.root.after(250, self._run_startup_actions)
        if start_minimized:
            self.root.iconify()

    def run(self) -> None:
        self.root.after(100, self._pump_messages)
        self.root.mainloop()

    def _apply_theme(self, theme_name: str) -> None:
        self.palette = get_theme(theme_name)
        self._configure_style()
        p = self.palette
        if hasattr(self, "log"):
            self.log.configure(bg=p["chrome"], fg=p["text"], insertbackground=p["text"])
            self.log.tag_configure("info", foreground=p["muted"])
            self.log.tag_configure("success", foreground=p["ok"])
            self.log.tag_configure("warning", foreground=p["warn"])
            self.log.tag_configure("error", foreground=p["error"])
        if hasattr(self, "tree"):
            self.tree.tag_configure("queued", foreground=p["muted"])
            self.tree.tag_configure("working", foreground=p["accent"])
            self.tree.tag_configure("done", foreground=p["ok"])
            self.tree.tag_configure("failed", foreground=p["error"])
            self.tree.tag_configure("skipped", foreground=p["warn"])
        for menu in (
            getattr(self, "queue_menu", None),
            getattr(self, "tray_menu", None),
            getattr(self, "log_menu", None),
        ):
            if menu is not None:
                menu.configure(bg=p["surface_2"], fg=p["text"], activebackground=p["selection"], activeforeground=p["text"])

    def _configure_style(self) -> None:
        p = self.palette
        self.root.configure(bg=p["bg"])
        self.style.theme_use("clam")
        self.style.configure(".", background=p["bg"], foreground=p["text"], fieldbackground=p["surface"], bordercolor=p["border"], lightcolor=p["border"], darkcolor=p["border"])
        self.style.configure("TFrame", background=p["bg"])
        self.style.configure("Chrome.TFrame", background=p["chrome"])
        self.style.configure("Surface.TFrame", background=p["surface"])
        self.style.configure("TLabel", background=p["bg"], foreground=p["text"])
        self.style.configure("Muted.TLabel", background=p["bg"], foreground=p["muted"])
        self.style.configure("Title.TLabel", background=p["chrome"], foreground=p["text"], font=("Segoe UI", 16, "bold"))
        self.style.configure("TButton", background=p["surface_2"], foreground=p["text"], padding=(12, 8), borderwidth=1)
        self.style.map("TButton", background=[("active", p["surface"]), ("pressed", p["border"])])
        self.style.configure("Accent.TButton", background=p["accent"], foreground=p["on_accent"], padding=(14, 8), font=("Segoe UI", 10, "bold"))
        self.style.map("Accent.TButton", background=[("active", p["accent_hover"]), ("pressed", p["accent"])])
        self.style.configure("Treeview", background=p["surface_2"], foreground=p["text"], fieldbackground=p["surface_2"], rowheight=30, borderwidth=0)
        self.style.configure("Treeview.Heading", background=p["surface"], foreground=p["muted"], padding=(8, 8), font=("Segoe UI", 9, "bold"))
        self.style.map("Treeview", background=[("selected", p["selection"])], foreground=[("selected", p["text"])])
        self.style.configure("TNotebook", background=p["bg"], borderwidth=0)
        self.style.configure("TNotebook.Tab", background=p["surface"], foreground=p["muted"], padding=(12, 8))
        self.style.map("TNotebook.Tab", background=[("selected", p["surface_2"])], foreground=[("selected", p["text"])])
        self.style.configure("TCombobox", fieldbackground=p["surface_2"], background=p["surface"], foreground=p["text"])

    def _build_ui(self) -> None:
        p = self.palette
        header = ttk.Frame(self.root, style="Chrome.TFrame", padding=(18, 14))
        header.pack(fill="x")
        title = ttk.Label(header, text="ExtractorX", style="Title.TLabel")
        title.pack(side="left")
        subtitle = "7-Zip ready" if self.sevenzip_path else "7-Zip not found"
        self.status_label = ttk.Label(header, text=subtitle, style="Muted.TLabel")
        self.status_label.pack(side="right")
        self.watch_status_label = ttk.Label(header, text="", style="Muted.TLabel")
        self.watch_status_label.pack(side="right", padx=(0, 12))

        toolbar = ttk.Frame(self.root, padding=(14, 12))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Add Files", command=self._add_files).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Scan Folder", command=self._scan_folder).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Extract", style="Accent.TButton", command=self._extract).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Test", command=self._test_archives).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Stop", command=self._stop).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Clear Done", command=self._clear_done).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Clear All", command=self._clear_all).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Identify", command=self._identify_dialog).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Export Script", command=self._export_batch_script).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="About", command=self._show_about).pack(side="right", padx=(8, 0))
        ttk.Button(toolbar, text="Settings", command=self._settings).pack(side="right", padx=(0, 8))
        self.bookmarks_button = ttk.Button(toolbar, text="Bookmarks", command=self._show_bookmarks_menu)
        self.bookmarks_button.pack(side="right", padx=(0, 8))

        destination_bar = ttk.Frame(self.root, padding=(14, 0, 14, 12))
        destination_bar.pack(fill="x")
        ttk.Label(destination_bar, text="Output").pack(side="left", padx=(0, 8))
        self.output_template_var = tk.StringVar(value=str(self.config.get("OutputPath", "")))
        ttk.Entry(destination_bar, textvariable=self.output_template_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(destination_bar, text="Browse", command=self._browse_main_output).pack(side="left")

        body = ttk.PanedWindow(self.root, orient="vertical")
        body.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        queue_frame = ttk.Frame(body, style="Surface.TFrame", padding=1)
        self.tree = ttk.Treeview(queue_frame, columns=("archive", "destination", "status", "size"), show="headings", selectmode="extended")
        self.tree.heading("archive", text="Archive", command=lambda: self._sort_queue("archive"))
        self.tree.heading("destination", text="Destination", command=lambda: self._sort_queue("destination"))
        self.tree.heading("status", text="Status", command=lambda: self._sort_queue("status"))
        self.tree.heading("size", text="Size", command=lambda: self._sort_queue("size"))
        self.tree.column("archive", width=330, anchor="w")
        self.tree.column("destination", width=430, anchor="w")
        self.tree.column("status", width=130, anchor="w")
        self.tree.column("size", width=90, anchor="e")
        y_scroll = ttk.Scrollbar(queue_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self._update_footer())
        self.tree.bind("<Button-3>", self._show_queue_menu)
        self.tree.bind("<Double-Button-1>", self._on_queue_double_click)
        self.tree.tag_configure("queued", foreground=p["muted"])
        self.tree.tag_configure("working", foreground=p["accent"])
        self.tree.tag_configure("done", foreground=p["ok"])
        self.tree.tag_configure("failed", foreground=p["error"])
        self.tree.tag_configure("skipped", foreground=p["warn"])
        self.tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        body.add(queue_frame, weight=4)

        log_frame = ttk.Frame(body, style="Surface.TFrame", padding=1)
        self.log = tk.Text(log_frame, height=9, bg=p["chrome"], fg=p["text"], insertbackground=p["text"], relief="flat", wrap="word", font=("Consolas", 9))
        self.log.tag_configure("info", foreground=p["muted"])
        self.log.tag_configure("success", foreground=p["ok"])
        self.log.tag_configure("warning", foreground=p["warn"])
        self.log.tag_configure("error", foreground=p["error"])
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log_menu = tk.Menu(self.root, tearoff=False, bg=p["surface_2"], fg=p["text"], activebackground=p["selection"], activeforeground=p["text"])
        self.log_menu.add_command(label="Copy", command=self._copy_log_selection)
        self.log_menu.add_command(label="Clear Log", command=self._clear_log)
        self.log_menu.add_separator()
        self.log_menu.add_command(label="Open Log Folder", command=self._open_log_folder)
        self.log.bind("<Button-3>", self._show_log_menu)
        body.add(log_frame, weight=1)

        footer = ttk.Frame(self.root, padding=(14, 0, 14, 12))
        footer.pack(fill="x")
        self.footer_label = ttk.Label(footer, text="Ready", style="Muted.TLabel")
        self.footer_label.pack(side="left")
        self.progress = ttk.Progressbar(footer, mode="determinate", maximum=100, value=0, length=180)
        self.progress.pack(side="right")

        self.queue_menu = tk.Menu(self.root, tearoff=False, bg=p["surface_2"], fg=p["text"], activebackground=p["selection"], activeforeground=p["text"])
        self.queue_menu.add_command(label="Extract Selected", command=self._extract)
        self.queue_menu.add_command(label="Test Selected", command=self._test_archives)
        self.queue_menu.add_command(label="Retry Selected", command=self._retry_selected)
        self.queue_menu.add_separator()
        self.queue_menu.add_command(label="Set Destination...", command=self._set_selected_destination)
        self.queue_menu.add_command(label="Open Destination", command=self._open_selected_destination)
        self.queue_menu.add_command(label="Reveal Archive", command=self._reveal_selected_archive)
        self.queue_menu.add_command(label="Copy Archive Path", command=self._copy_selected_archive_path)
        self.queue_menu.add_command(label="Copy Destination Path", command=self._copy_selected_destination_path)
        self.queue_menu.add_separator()
        self.queue_menu.add_command(label="Remove Selected", command=self._remove_selected)

        self.tray_menu = tk.Menu(self.root, tearoff=False, bg=p["surface_2"], fg=p["text"], activebackground=p["selection"], activeforeground=p["text"])
        self.tray_menu.add_command(label="Show ExtractorX", command=self._restore_from_tray)
        self.tray_menu.add_separator()
        self.tray_menu.add_command(label="Extract Queued", command=self._extract)
        self.tray_menu.add_command(label="Test Queued", command=self._test_archives)
        self.tray_menu.add_command(label="Stop", command=self._stop)
        self.tray_menu.add_command(label="Clear Done", command=self._clear_done)
        self.tray_menu.add_separator()
        self.tray_menu.add_command(label="Exit", command=self._on_close)

    def add_item(self, item: QueueItem) -> None:
        if any(existing.archive_path == item.archive_path for existing in self.items.values()):
            self._log(f"Already queued: {item.archive_path.name}", "warning")
            return
        item.output_path = Path(item.output_override).expanduser() if item.output_override else resolve_output_path(self._current_output_template(), item.archive_path)
        self.items[item.id] = item
        self.tree.insert("", "end", iid=item.id, values=(str(item.archive_path), str(item.output_path), item.status.value, format_size(item.size_bytes)), tags=(self._status_tag(item.status),))
        self._update_footer()

    def _current_output_template(self) -> str:
        value = self.output_template_var.get().strip()
        self.config["OutputPath"] = value
        return value

    def _browse_main_output(self) -> None:
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_template_var.set(folder)
            self.config["OutputPath"] = folder
            for item in self.items.values():
                if not item.output_override and item.status in {QueueStatus.QUEUED, QueueStatus.FAILED, QueueStatus.TEST_OK}:
                    item.output_path = resolve_output_path(folder, item.archive_path)
                    self._update_item(item)

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(title="Select archives")
        for raw in files:
            self._add_archive_path(Path(raw), apply_drop_filter=False)

    def _scan_folder(self) -> None:
        folder = filedialog.askdirectory(title="Scan folder for archives")
        if not folder:
            return
        path = Path(folder).expanduser()
        if _is_drive_root(path):
            if not messagebox.askyesno(
                "Scan drive root?",
                f"You picked {path}. Scanning an entire drive can take a long time and pick up many unrelated files. Continue?",
                parent=self.root,
            ):
                return
        if not self.service.scan_paths([path]):
            self._log("A scan or extraction is already running.", "warning")

    def _add_archive_path(self, path: Path, apply_drop_filter: bool = False) -> bool:
        if apply_drop_filter and not self._passes_drag_filter(path):
            self._log(f"Skipped by drag/drop filter: {path.name}", "warning")
            return False
        if is_supported_archive(path, bool(self.config.get("DeepArchiveDetection", True))):
            self.add_item(QueueItem.from_path(path))
            return True
        self._log(f"Skipped unsupported file: {path.name}", "warning")
        return False

    def _handle_dropped_paths(self, paths: list[str]) -> None:
        added = 0
        folders: list[Path] = []
        for raw in paths:
            path = Path(raw)
            if path.is_dir():
                folders.append(path)
            elif path.is_file() and self._add_archive_path(path, apply_drop_filter=True):
                added += 1
        if folders:
            if not self.service.scan_paths(folders):
                self._log("Dropped folder queued for scanning, but current work is still running.", "warning")
            else:
                self._log(f"Scanning {len(folders)} dropped folder(s).", "success")
        if added and bool(self.config.get("AutoExtractOnDrop", False)) and not self.service.active:
            self._extract()

    def _passes_drag_filter(self, path: Path) -> bool:
        mode = str(self.config.get("DragDropFilterType", "None"))
        masks = [mask.strip() for mask in str(self.config.get("DragDropFilterMask", "")).split(";") if mask.strip()]
        if mode == "None" or not masks:
            return True
        matched = any(fnmatch.fnmatch(path.name, mask) for mask in masks)
        if mode == "Inclusion":
            return matched
        if mode == "Exclusion":
            return not matched
        return True

    _ELIGIBLE_STATUSES = frozenset(
        {QueueStatus.QUEUED, QueueStatus.FAILED, QueueStatus.TEST_OK, QueueStatus.PASSWORD_REQUIRED}
    )

    def _eligible_items(self) -> list[QueueItem]:
        selected = self.tree.selection()
        if selected:
            return [
                self.items[item_id]
                for item_id in selected
                if item_id in self.items and self.items[item_id].status in self._ELIGIBLE_STATUSES
            ]
        return [item for item in self.items.values() if item.status in self._ELIGIBLE_STATUSES]

    def _extract(self) -> None:
        self.config["OutputPath"] = self._current_output_template()
        items = self._eligible_items()
        self._start_items(items, test_only=False)

    def _test_archives(self) -> None:
        self.config["OutputPath"] = self._current_output_template()
        items = self._eligible_items()
        self._start_items(items, test_only=True)

    def _start_items(self, items: list[QueueItem], test_only: bool) -> None:
        if not items:
            self._log("No queued archives are ready.", "warning")
            return
        if self.service.active:
            self._log("A scan or extraction is already running.", "warning")
            return
        self.progress.configure(value=0)
        for item in items:
            item.status = QueueStatus.TESTING if test_only else QueueStatus.EXTRACTING
            self._update_item(item)
        if not self.service.extract_items(items, test_only=test_only):
            for item in items:
                item.status = QueueStatus.QUEUED
                self._update_item(item)
            self._log("A scan or extraction is already running.", "warning")

    def _stop(self) -> None:
        if not self.service.active:
            self._log("Nothing to stop.", "info")
            return
        self.service.stop()
        self._log("Stop requested.", "warning")
        self.status_label.configure(text="Stopping...")

    def _clear_done(self) -> None:
        for item_id, item in list(self.items.items()):
            if item.status in {QueueStatus.DONE, QueueStatus.SKIPPED, QueueStatus.TEST_OK}:
                self.tree.delete(item_id)
                del self.items[item_id]
        self._update_footer()

    def _clear_all(self) -> None:
        if self.service.active:
            self._log("Stop active work before clearing the queue.", "warning")
            return
        for item_id in list(self.items):
            self.tree.delete(item_id)
        self.items.clear()
        self.progress.configure(value=0)
        self._update_footer()

    def _settings(self) -> None:
        SettingsDialog(self)

    def _identify_dialog(self) -> None:
        paths = filedialog.askopenfilenames(title="Identify files", parent=self.root)
        if not paths:
            return
        lines = []
        for raw in paths:
            result = identify(Path(raw).expanduser())
            status = "OK" if result.supported else "NOT SUPPORTED"
            lines.append(f"{result.path.name}\n  format: {result.format}\n  status: {status}\n  {result.reason}")
        messagebox.showinfo("Identify", "\n\n".join(lines), parent=self.root)

    def _export_batch_script(self) -> None:
        if not self.items:
            self._log("Queue is empty -- nothing to export.", "warning")
            return
        target = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export batch script",
            defaultextension=".cmd",
            filetypes=(("Windows batch", "*.cmd"), ("Shell", "*.sh"), ("All", "*.*")),
        )
        if not target:
            return
        entrypoint = Path(sys.executable)
        script = Path(__file__).resolve().parents[1] / "ExtractorX.py"
        lines = ["@echo off", f"REM ExtractorX v{__version__} batch export", ""]
        for item in self.items.values():
            parts = [_cmd_quote(str(entrypoint)), _cmd_quote(str(script))]
            if item.output_override:
                parts.extend(["--target", _cmd_quote(item.output_override)])
            parts.append("--auto-extract")
            parts.append(_cmd_quote(str(item.archive_path)))
            lines.append(" ".join(parts))
        try:
            Path(target).write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
            self._log(f"Exported {len(self.items)} item(s) to {target}", "success")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc), parent=self.root)

    def _show_bookmarks_menu(self) -> None:
        menu = tk.Menu(self.root, tearoff=False, bg=self.palette["surface_2"], fg=self.palette["text"], activebackground=self.palette["selection"], activeforeground=self.palette["text"])
        bookmarks = list(self.config.get("Bookmarks", []) or [])
        if not bookmarks:
            menu.add_command(label="(no bookmarks)", state="disabled")
        else:
            for bookmark in bookmarks:
                label = str(bookmark.get("Label", "")) or str(bookmark.get("Path", ""))
                path = str(bookmark.get("Path", ""))
                if path:
                    menu.add_command(label=label, command=lambda target=path: self._open_bookmark(target))
        menu.add_separator()
        menu.add_command(label="Add current output as bookmark...", command=self._add_output_bookmark)
        menu.add_command(label="Manage bookmarks in Settings...", command=self._settings)
        x = self.bookmarks_button.winfo_rootx()
        y = self.bookmarks_button.winfo_rooty() + self.bookmarks_button.winfo_height()
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _open_bookmark(self, target: str) -> None:
        path = Path(target).expanduser()
        if path.is_dir():
            if not self.service.scan_paths([path]):
                self._log("A scan or extraction is already running.", "warning")
        elif path.is_file():
            self._add_archive_path(path)

    def _add_output_bookmark(self) -> None:
        value = simpledialog.askstring("Add bookmark", "Path (file or folder):", parent=self.root)
        if not value:
            return
        label = simpledialog.askstring("Add bookmark", "Label:", parent=self.root) or value
        bookmarks = list(self.config.get("Bookmarks", []) or [])
        bookmarks.append({"Label": label, "Path": value})
        self.config["Bookmarks"] = bookmarks
        save_config(self.config)
        self._log("Bookmark added.", "success")

    def _show_about(self) -> None:
        sevenzip = str(self.sevenzip_path) if self.sevenzip_path else "not found (will download on first extract)"
        messagebox.showinfo(
            "About ExtractorX",
            "ExtractorX Python\n\nBulk archive extraction for Windows with queueing, monitoring, "
            "nested extraction, password cycling, and 7-Zip integration.\n\n"
            f"7-Zip: {sevenzip}\n"
            f"Config: {app_data_dir()}\n\n"
            "Inspired by ExtractNow by Nathan Moinvaziri. Powered by 7-Zip.",
            parent=self.root,
        )

    def _run_startup_actions(self) -> None:
        if self.startup_scan_paths:
            if not self.service.scan_paths(self.startup_scan_paths):
                self._log("Could not start startup scan because work is already running.", "warning")
        elif self.test_startup and self.items:
            self._test_archives()
        elif self.auto_extract_startup and self.items:
            self._extract()

    def _start_watchers(self) -> None:
        if self.watch_service:
            self.watch_service.stop()
        folders = [str(folder) for folder in self.config.get("WatchFolders", []) if str(folder).strip()]
        self.watch_service = WatchService(folders, self.watch_queue, bool(self.config.get("DeepArchiveDetection", True)))
        self.watch_service.start()
        if hasattr(self, "watch_status_label"):
            self.watch_status_label.configure(text=f"Watching {len(folders)} folder(s)" if folders else "")

    def _pump_messages(self) -> None:
        if self.closing:
            return
        self._pump_watch_queue()
        while True:
            try:
                message = self.messages.get_nowait()
            except Empty:
                break
            try:
                self._handle_message(message)
            except tk.TclError:
                # Widgets may be torn down mid-message when the user closes the window.
                return
        if not self.closing:
            self.root.after(100, self._pump_messages)

    def _pump_watch_queue(self) -> None:
        detected: list[QueueItem] = []
        while True:
            try:
                path = self.watch_queue.get_nowait()
            except Empty:
                break
            if not path.exists():
                continue
            if any(existing.archive_path == path for existing in self.items.values()):
                continue
            if not is_supported_archive(path, bool(self.config.get("DeepArchiveDetection", True))):
                self._log(f"Skipped watched file (not an archive): {path.name}", "warning")
                continue
            item = QueueItem.from_path(path)
            self.add_item(item)
            self._log(f"Detected watched archive: {path.name}", "success")
            detected.append(item)
        if not detected or not bool(self.config.get("WatchAutoExtract", True)):
            return
        if self.service.active:
            self._log("Watched archive(s) queued; current work is still running.", "warning")
            return
        for item in detected:
            item.status = QueueStatus.EXTRACTING
            self._update_item(item)
        if not self.service.extract_items(detected, test_only=False):
            for item in detected:
                item.status = QueueStatus.QUEUED
                self._update_item(item)

    def _handle_message(self, message: OperationMessage) -> None:
        if message.type == "sevenzip_ready":
            path_payload = str(message.payload.get("path", "") or "")
            if path_payload:
                self.sevenzip_path = Path(path_payload)
                self.service.sevenzip_path = self.sevenzip_path
            self.status_label.configure(text="7-Zip ready")
        if message.type == "archive_found":
            raw = str(message.payload.get("path", "") or "")
            if raw:
                candidate = Path(raw)
                if candidate.exists() and not any(
                    existing.archive_path == candidate for existing in self.items.values()
                ):
                    self.add_item(QueueItem.from_path(candidate))
            return
        if message.item_id and message.item_id in self.items:
            item = self.items[message.item_id]
            if message.type == "item_started":
                item.status = QueueStatus.TESTING if bool(message.payload.get("test_only")) else QueueStatus.EXTRACTING
            elif message.type == "item_done":
                item.status = QueueStatus.TEST_OK if bool(message.payload.get("test_only")) else QueueStatus.DONE
            elif message.type == "item_failed":
                item.status = QueueStatus.FAILED
                item.error = message.text
                if bool(self.config.get("PromptOnExhaustion", False)) and _looks_like_password_failure(message.text):
                    self.password_retry_candidates.append(item)
            elif message.type == "item_skipped":
                item.status = QueueStatus.SKIPPED
            self._update_item(item)
        if message.text:
            level = "success" if message.type.endswith("_done") else message.level
            self._log(message.text, level)
        if message.type == "progress":
            total = int(message.payload.get("total", 0) or 0)
            current = int(message.payload.get("current", 0) or 0)
            self.progress.configure(value=(current / total * 100) if total else 0)
            self.status_label.configure(text=message.text)
            if self.shell_bridge:
                self.shell_bridge.update_tray_tip(f"ExtractorX - {message.text}")
        if message.type == "extract_done":
            self.progress.configure(value=100)
            self.status_label.configure(text="7-Zip ready" if self.sevenzip_path else "7-Zip not found")
            self._play_completion_sound()
            if self.password_retry_candidates:
                self.root.after(250, self._prompt_password_retry)
            if bool(self.config.get("ClearListOnComplete", False)):
                self._clear_done()
            if bool(self.config.get("CloseOnCompleteAlways", False)) or (
                bool(self.config.get("CloseOnComplete", False))
                and all(item.status in {QueueStatus.DONE, QueueStatus.SKIPPED} for item in self.items.values())
            ):
                self.root.after(800, self._close_after_complete)
        self._update_footer()

    def _update_item(self, item: QueueItem) -> None:
        self.tree.item(item.id, values=(str(item.archive_path), str(item.output_path or ""), item.status.value, format_size(item.size_bytes)))
        self.tree.item(item.id, tags=(self._status_tag(item.status),))

    @staticmethod
    def _status_tag(status: QueueStatus) -> str:
        if status in {QueueStatus.EXTRACTING, QueueStatus.TESTING, QueueStatus.SCANNING}:
            return "working"
        if status in {QueueStatus.DONE, QueueStatus.TEST_OK}:
            return "done"
        if status == QueueStatus.FAILED:
            return "failed"
        if status in {QueueStatus.SKIPPED, QueueStatus.PASSWORD_REQUIRED}:
            return "skipped"
        return "queued"

    def _update_footer(self) -> None:
        counts = {status: 0 for status in QueueStatus}
        for item in self.items.values():
            counts[item.status] += 1
        selected = [self.items[item_id] for item_id in self.tree.selection() if item_id in self.items]
        if selected:
            selected_size = sum(item.size_bytes for item in selected)
            prefix = f"{len(selected)} selected, {format_size(selected_size)} | "
        else:
            prefix = ""
        self.footer_label.configure(text=f"{prefix}{len(self.items)} item(s) | {counts[QueueStatus.QUEUED]} queued | {counts[QueueStatus.DONE]} done | {counts[QueueStatus.TEST_OK]} tested | {counts[QueueStatus.FAILED]} failed")

    def _log(self, text: str, level: str = "info") -> None:
        self.log.insert("end", text + "\n", level if level in {"info", "success", "warning", "error"} else "info")
        line_count = int(self.log.index("end-1c").split(".")[0])
        if line_count > LOG_LINE_CAP:
            self.log.delete("1.0", f"{line_count - LOG_LINE_CAP}.0")
        self.log.see("end")
        if bool(self.config.get("LogHistory", True)):
            try:
                log_dir = app_data_dir() / "logs"
                log_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%H:%M:%S")
                log_path = log_dir / f"{datetime.now():%Y-%m-%d}.log"
                with log_path.open("a", encoding="utf-8") as handle:
                    handle.write(f"[{stamp}] [{level.upper()}] {text}\n")
            except OSError:
                pass

    def _sort_queue(self, column: str) -> None:
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False

        def key_for(item_id: str) -> object:
            item = self.items[item_id]
            if column == "archive":
                return item.archive_path.name.lower()
            if column == "destination":
                return str(item.output_path or "").lower()
            if column == "status":
                return item.status.value
            if column == "size":
                return item.size_bytes
            return item.created_at

        for index, item_id in enumerate(sorted(self.items, key=key_for, reverse=self.sort_reverse)):
            self.tree.move(item_id, "", index)

    def _show_queue_menu(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        if row_id and row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
        try:
            self.queue_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.queue_menu.grab_release()

    def _selected_items(self) -> list[QueueItem]:
        return [self.items[item_id] for item_id in self.tree.selection() if item_id in self.items]

    def _remove_selected(self) -> None:
        for item in self._selected_items():
            self.tree.delete(item.id)
            self.items.pop(item.id, None)
        self._update_footer()

    def _retry_selected(self) -> None:
        items = self._selected_items() or [
            item
            for item in self.items.values()
            if item.status in {QueueStatus.FAILED, QueueStatus.PASSWORD_REQUIRED}
        ]
        if not items:
            self._log("Nothing to retry.", "warning")
            return
        skipped = 0
        retriable: list[QueueItem] = []
        for item in items:
            if item.status in {QueueStatus.EXTRACTING, QueueStatus.TESTING}:
                skipped += 1
                continue
            item.status = QueueStatus.QUEUED
            item.error = ""
            self._update_item(item)
            retriable.append(item)
        if skipped:
            self._log(f"Skipped {skipped} item(s) already running.", "warning")
        if not retriable:
            return
        self._start_items(retriable, test_only=False)

    def _set_selected_destination(self) -> None:
        selected = self._selected_items()
        if not selected:
            self._log("Select at least one archive first.", "warning")
            return
        folder = filedialog.askdirectory(title="Choose destination folder")
        if not folder:
            return
        for item in selected:
            if item.status in {QueueStatus.EXTRACTING, QueueStatus.TESTING}:
                continue
            item.output_override = folder
            item.output_path = resolve_output_path(folder, item.archive_path)
            self._update_item(item)
        self._log(f"Destination override set for {len(selected)} item(s).", "success")

    def _open_selected_destination(self) -> None:
        for item in self._selected_items()[:1]:
            target = item.output_path
            if target and target.exists():
                _open_path(target)
            else:
                self._log("Destination does not exist yet.", "warning")

    def _reveal_selected_archive(self) -> None:
        for item in self._selected_items()[:1]:
            if item.archive_path.exists():
                _reveal_path(item.archive_path)
            else:
                self._log("Archive no longer exists.", "warning")

    def _copy_selected_archive_path(self) -> None:
        paths = [str(item.archive_path) for item in self._selected_items()]
        if paths:
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(paths))
            self._log("Copied archive path.", "success")

    def _copy_selected_destination_path(self) -> None:
        paths = [str(item.output_path) for item in self._selected_items() if item.output_path]
        if paths:
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(paths))
            self._log("Copied destination path.", "success")

    def _play_completion_sound(self) -> None:
        if not bool(self.config.get("SoundsEnabled", True)):
            return
        failed = any(item.status == QueueStatus.FAILED for item in self.items.values())
        try:
            winsound.MessageBeep(winsound.MB_ICONHAND if failed else winsound.MB_ICONASTERISK)
        except RuntimeError:
            pass

    def _prompt_password_retry(self) -> None:
        if self.service.active:
            self.root.after(300, self._prompt_password_retry)
            return
        seen: set[str] = set()
        candidates: list[QueueItem] = []
        for item in self.password_retry_candidates:
            if item.id not in seen:
                seen.add(item.id)
                candidates.append(item)
        self.password_retry_candidates.clear()
        if not candidates:
            return
        candidates = [item for item in candidates if item.id in self.items and item.status == QueueStatus.FAILED]
        if not candidates:
            return
        password = _ask_password_with_entropy(
            self.root,
            "Password required",
            f"{len(candidates)} archive(s) may need another password. Enter a password to retry them:",
        )
        if not password:
            return
        if password not in self.passwords:
            self.passwords.append(password)
            self.password_store.save(self.passwords)
        for item in candidates:
            item.status = QueueStatus.QUEUED
            self._update_item(item)
        self._start_items(candidates, test_only=False)

    def _on_close(self) -> None:
        self.closing = True
        if self.service.active:
            if not messagebox.askyesno("Stop active work?", "Extraction or scanning is still running. Stop it and close ExtractorX?"):
                self.closing = False
                return
            self.service.stop()
        self._close_without_prompt()

    def _close_after_complete(self) -> None:
        if self.service.active:
            self.root.after(500, self._close_after_complete)
            return
        self._close_without_prompt()

    def _close_without_prompt(self) -> None:
        if getattr(self, "_destroyed", False):
            return
        self.closing = True
        try:
            self.config["OutputPath"] = self._current_output_template()
        except tk.TclError:
            pass
        if self.watch_service:
            try:
                self.watch_service.stop()
            except Exception:  # noqa: BLE001 - best-effort shutdown
                pass
        if self.shell_bridge:
            try:
                self.shell_bridge.dispose()
            except Exception:  # noqa: BLE001
                pass
        try:
            geometry = self.root.winfo_geometry()
        except tk.TclError:
            geometry = ""
        match = re.match(r"(?P<width>\d+)x(?P<height>\d+)\+(?P<left>-?\d+)\+(?P<top>-?\d+)", geometry)
        if match:
            self.config["WindowWidth"] = int(match.group("width"))
            self.config["WindowHeight"] = int(match.group("height"))
            self.config["WindowLeft"] = int(match.group("left"))
            self.config["WindowTop"] = int(match.group("top"))
        else:
            try:
                self.config["WindowWidth"] = self.root.winfo_width()
                self.config["WindowHeight"] = self.root.winfo_height()
            except tk.TclError:
                pass
        try:
            save_config(self.config)
        except OSError as exc:
            # Surface via logger only -- the user is already closing the app.
            import logging

            logging.getLogger("extractorx.ui").warning("Could not save config on close: %s", exc)
        self._destroyed = True
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _handle_unmap(self, _event: tk.Event) -> None:
        if self.closing or not bool(self.config.get("MinimizeToTray", True)):
            return
        if self.root.state() == "iconic":
            self.root.after(20, self._minimize_to_tray)

    def _minimize_to_tray(self) -> None:
        if self.closing or not self.shell_bridge:
            return
        self.shell_bridge.add_tray_icon("ExtractorX")
        self.root.withdraw()

    def _restore_from_tray(self) -> None:
        if self.shell_bridge:
            self.shell_bridge.remove_tray_icon()
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()

    def _on_queue_double_click(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        if not row_id or row_id not in self.items:
            return
        item = self.items[row_id]
        if item.output_path and item.status in {QueueStatus.DONE, QueueStatus.TEST_OK} and item.output_path.exists():
            _open_path(item.output_path)
        elif item.archive_path.exists():
            _reveal_path(item.archive_path)

    def _show_log_menu(self, event: tk.Event) -> None:
        try:
            self.log_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.log_menu.grab_release()

    def _copy_log_selection(self) -> None:
        try:
            text = self.log.get("sel.first", "sel.last")
        except tk.TclError:
            text = self.log.get("1.0", "end").strip()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def _open_log_folder(self) -> None:
        path = app_data_dir() / "logs"
        path.mkdir(parents=True, exist_ok=True)
        _open_path(path)

    def _show_tray_menu(self) -> None:
        try:
            self.tray_menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            self.tray_menu.grab_release()


class SettingsDialog:
    def __init__(self, app: ExtractorXApp) -> None:
        self.app = app
        self.config = dict(app.config)
        self.window = tk.Toplevel(app.root)
        self.window.title("Settings")
        self.window.transient(app.root)
        self.window.grab_set()
        self.window.geometry("620x560")
        self.window.configure(bg=app.palette["bg"])
        self._build()

    def _build(self) -> None:
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill="both", expand=True, padx=14, pady=14)

        general = ttk.Frame(notebook, padding=16)
        notebook.add(general, text="General")
        ttk.Label(general, text="Theme").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.theme_var = tk.StringVar(value=str(self.config.get("Theme", "Midnight")))
        ttk.Combobox(general, textvariable=self.theme_var, values=THEME_NAMES, state="readonly", width=24).grid(row=0, column=1, sticky="w", pady=(0, 8))
        self.always_var = tk.BooleanVar(value=bool(self.config.get("AlwaysOnTop", False)))
        self.minimize_tray_var = tk.BooleanVar(value=bool(self.config.get("MinimizeToTray", True)))
        self.log_history_var = tk.BooleanVar(value=bool(self.config.get("LogHistory", True)))
        self.watch_auto_var = tk.BooleanVar(value=bool(self.config.get("WatchAutoExtract", True)))
        self.nested_var = tk.BooleanVar(value=bool(self.config.get("NestedExtraction", True)))
        self.sounds_var = tk.BooleanVar(value=bool(self.config.get("SoundsEnabled", True)))
        self.auto_drop_var = tk.BooleanVar(value=bool(self.config.get("AutoExtractOnDrop", False)))
        ttk.Checkbutton(general, text="Always on top", variable=self.always_var).grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Minimize to tray", variable=self.minimize_tray_var).grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Save log history", variable=self.log_history_var).grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Auto-extract watched archives", variable=self.watch_auto_var).grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Extract nested archives", variable=self.nested_var).grid(row=5, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Play completion sounds", variable=self.sounds_var).grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(general, text="Auto-extract dropped archives", variable=self.auto_drop_var).grid(row=7, column=0, columnspan=2, sticky="w")

        destination = ttk.Frame(notebook, padding=16)
        notebook.add(destination, text="Destination")
        ttk.Label(destination, text="Output path template").pack(anchor="w")
        self.output_var = tk.StringVar(value=str(self.config.get("OutputPath", "")))
        ttk.Entry(destination, textvariable=self.output_var).pack(fill="x", pady=(6, 10))
        ttk.Label(
            destination,
            text=(
                "Tokens: {ArchiveFolder}, {ArchiveName}, {ArchiveNameUnique}, {ArchiveExtension}, "
                "{ArchiveFileName}, {ArchiveFolderName}, {ArchivePath}, {Desktop}, {UserProfile}, "
                "{Program Files}, {Windows}, {Date}, {Time}, {Guid}, {Env:NAME}"
            ),
            style="Muted.TLabel",
            wraplength=540,
            justify="left",
        ).pack(anchor="w")
        ttk.Label(destination, text="Overwrite mode").pack(anchor="w", pady=(16, 4))
        self.overwrite_var = tk.StringVar(value=str(self.config.get("OverwriteMode", "Always")))
        ttk.Combobox(destination, textvariable=self.overwrite_var, values=("Always", "Never", "Rename"), state="readonly", width=20).pack(anchor="w")
        ttk.Label(destination, text="Smart Extract").pack(anchor="w", pady=(16, 4))
        self.smart_extract_var = tk.StringVar(value=str(self.config.get("SmartExtract", "Auto")))
        ttk.Combobox(destination, textvariable=self.smart_extract_var, values=SMART_EXTRACT_MODES, state="readonly", width=20).pack(anchor="w")
        ttk.Label(destination, text="Auto: flatten duplicate archive-name folder. AlwaysWrap: keep 7-Zip layout as-is. NeverWrap: unwrap any single child.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(2, 8))
        ttk.Label(destination, text="Archive filename encoding").pack(anchor="w", pady=(16, 4))
        self.filename_encoding_var = tk.StringVar(value=str(self.config.get("FilenameEncoding", "Auto")))
        ttk.Combobox(destination, textvariable=self.filename_encoding_var, values=FILENAME_ENCODINGS, state="readonly", width=20).pack(anchor="w")
        ttk.Label(destination, text="Select a code page when extracting archives from other locales (cp932 Japanese, cp936 Simplified Chinese, cp1251 Cyrillic, etc.).", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(2, 4))

        process = ttk.Frame(notebook, padding=16)
        notebook.add(process, text="Process")
        self.nested_depth_var = tk.IntVar(value=int(self.config.get("NestedMaxDepth", 5)))
        self.nested_post_var = tk.BooleanVar(value=bool(self.config.get("NestedApplyPostAction", False)))
        self.delete_after_var = tk.BooleanVar(value=bool(self.config.get("DeleteAfterExtract", False)))
        self.post_action_var = tk.StringVar(value=str(self.config.get("PostAction", "None")))
        self.post_folder_var = tk.StringVar(value=str(self.config.get("PostActionFolder", "")))
        self.open_dest_var = tk.BooleanVar(value=bool(self.config.get("OpenDestAfterExtract", False)))
        self.remove_dupe_var = tk.BooleanVar(value=bool(self.config.get("RemoveDuplicateFolder", True)))
        self.rename_single_var = tk.BooleanVar(value=bool(self.config.get("RenameSingleFile", False)))
        self.delete_broken_var = tk.BooleanVar(value=bool(self.config.get("DeleteBrokenFiles", False)))
        self.clear_done_var = tk.BooleanVar(value=bool(self.config.get("ClearListOnComplete", False)))
        self.close_done_var = tk.BooleanVar(value=bool(self.config.get("CloseOnComplete", False)))
        ttk.Label(process, text="Nested extraction depth").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Spinbox(process, from_=1, to=50, textvariable=self.nested_depth_var, width=8).grid(row=0, column=1, sticky="w", pady=(0, 8))
        ttk.Checkbutton(process, text="Apply source cleanup to nested archives", variable=self.nested_post_var).grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Delete source archive after extraction", variable=self.delete_after_var).grid(row=2, column=0, columnspan=3, sticky="w")
        ttk.Label(process, text="Post action").grid(row=3, column=0, sticky="w", pady=(12, 8))
        ttk.Combobox(process, textvariable=self.post_action_var, values=("None", "Recycle", "MoveToFolder", "Delete"), state="readonly", width=18).grid(row=3, column=1, sticky="w", pady=(12, 8))
        ttk.Label(process, text="Move folder").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(process, textvariable=self.post_folder_var).grid(row=4, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(process, text="Browse", command=self._browse_post_folder).grid(row=4, column=2, sticky="w", padx=(8, 0), pady=(0, 8))
        ttk.Checkbutton(process, text="Open destination after extraction", variable=self.open_dest_var).grid(row=5, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Remove duplicate archive-name folder", variable=self.remove_dupe_var).grid(row=6, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Rename a single extracted file to the archive name", variable=self.rename_single_var).grid(row=7, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Delete incomplete output after failure", variable=self.delete_broken_var).grid(row=8, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Clear successful items when a batch completes", variable=self.clear_done_var).grid(row=9, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Checkbutton(process, text="Close when the full batch succeeds", variable=self.close_done_var).grid(row=10, column=0, columnspan=3, sticky="w")
        process.columnconfigure(1, weight=1)

        monitor = ttk.Frame(notebook, padding=16)
        notebook.add(monitor, text="Monitor")
        self.watch_list = tk.Listbox(monitor, height=9, bg=self.app.palette["surface_2"], fg=self.app.palette["text"], selectbackground=self.app.palette["selection"], relief="flat")
        self.watch_list.pack(fill="both", expand=True)
        for folder in self.config.get("WatchFolders", []):
            self.watch_list.insert("end", folder)
        watch_buttons = ttk.Frame(monitor)
        watch_buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(watch_buttons, text="Add Folder", command=self._add_watch_folder).pack(side="left", padx=(0, 8))
        ttk.Button(watch_buttons, text="Remove", command=self._remove_watch_folder).pack(side="left")

        passwords = ttk.Frame(notebook, padding=16)
        notebook.add(passwords, text="Passwords")
        self.use_passwords_var = tk.BooleanVar(value=bool(self.config.get("UsePasswordList", True)))
        self.prompt_password_var = tk.BooleanVar(value=bool(self.config.get("PromptOnExhaustion", False)))
        self.assume_one_password_var = tk.BooleanVar(value=bool(self.config.get("AssumeOnePassword", True)))
        self.use_sidecars_var = tk.BooleanVar(value=bool(self.config.get("UsePasswordSidecars", True)))
        self.hash_probe_var = tk.BooleanVar(value=bool(self.config.get("HashModePasswordProbe", True)))
        self.wordlist_var = tk.BooleanVar(value=bool(self.config.get("WordlistGeneration", False)))
        self.wordlist_max_var = tk.IntVar(value=int(self.config.get("WordlistMaxAttempts", 500)))
        self.password_timeout_var = tk.IntVar(value=int(self.config.get("PasswordTimeout", 45)))
        ttk.Checkbutton(passwords, text="Use saved password list", variable=self.use_passwords_var).pack(anchor="w")
        ttk.Checkbutton(passwords, text="Read sidecar password files (archive.pwd.txt, passwords.txt)", variable=self.use_sidecars_var).pack(anchor="w")
        ttk.Checkbutton(passwords, text="Prompt for another password when saved passwords fail", variable=self.prompt_password_var).pack(anchor="w")
        ttk.Checkbutton(passwords, text="Assume one successful password applies to the batch", variable=self.assume_one_password_var).pack(anchor="w")
        ttk.Checkbutton(passwords, text="Fast test-mode password probe (7z t)", variable=self.hash_probe_var).pack(anchor="w")
        ttk.Checkbutton(passwords, text="Generate wordlist variants (case, leet, date suffixes)", variable=self.wordlist_var).pack(anchor="w")
        wordlist_max_frame = ttk.Frame(passwords)
        wordlist_max_frame.pack(fill="x", pady=(0, 4))
        ttk.Label(wordlist_max_frame, text="Max wordlist candidates").pack(side="left")
        ttk.Spinbox(wordlist_max_frame, from_=10, to=10000, textvariable=self.wordlist_max_var, width=8).pack(side="left", padx=(8, 0))
        timeout_frame = ttk.Frame(passwords)
        timeout_frame.pack(fill="x", pady=(8, 4))
        ttk.Label(timeout_frame, text="Password retry timeout").pack(side="left")
        ttk.Spinbox(timeout_frame, from_=5, to=600, textvariable=self.password_timeout_var, width=8).pack(side="left", padx=(8, 0))
        skip_frame = ttk.Frame(passwords)
        skip_frame.pack(fill="x", pady=(0, 8))
        self.skip_passwords_var = tk.IntVar(value=int(self.config.get("SkipAfterFailedPasswords", 0)))
        ttk.Label(skip_frame, text="Give up after N failed password attempts (0 = unlimited)").pack(side="left")
        ttk.Spinbox(skip_frame, from_=0, to=9999, textvariable=self.skip_passwords_var, width=8).pack(side="left", padx=(8, 0))
        self.password_list = tk.Listbox(passwords, height=9, bg=self.app.palette["surface_2"], fg=self.app.palette["text"], selectbackground=self.app.palette["selection"], relief="flat")
        self.password_list.pack(fill="both", expand=True)
        for password in self.app.passwords:
            self.password_list.insert("end", _mask_password(password))
        pass_buttons = ttk.Frame(passwords)
        pass_buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(pass_buttons, text="Add Password", command=self._add_password).pack(side="left", padx=(0, 8))
        ttk.Button(pass_buttons, text="Import File", command=self._import_passwords).pack(side="left", padx=(0, 8))
        ttk.Button(pass_buttons, text="Remove", command=self._remove_password).pack(side="left")

        files = ttk.Frame(notebook, padding=16)
        notebook.add(files, text="Files")
        ttk.Label(files, text="Extraction exclusions").pack(anchor="w")
        ttk.Label(files, text="Semicolon-separated masks passed to 7-Zip, for example Thumbs.db;desktop.ini", style="Muted.TLabel").pack(anchor="w", pady=(2, 8))
        self.exclusions_var = tk.StringVar(value=str(self.config.get("FileExclusions", "")))
        ttk.Entry(files, textvariable=self.exclusions_var).pack(fill="x")
        ttk.Label(files, text="Extraction include masks").pack(anchor="w", pady=(14, 4))
        ttk.Label(files, text="Only files matching these masks will be extracted. Empty = extract everything.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(0, 8))
        self.include_masks_var = tk.StringVar(value=str(self.config.get("IncludeMasks", "")))
        ttk.Entry(files, textvariable=self.include_masks_var).pack(fill="x")
        ttk.Label(files, text="Archive handler allowlist").pack(anchor="w", pady=(14, 4))
        ttk.Label(files, text="Semicolon-separated extensions. Empty = allow all. Example: zip;7z;rar", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(0, 8))
        self.handler_allowlist_var = tk.StringVar(value=";".join(str(ext) for ext in self.config.get("HandlerAllowlist", []) or []))
        ttk.Entry(files, textvariable=self.handler_allowlist_var).pack(fill="x")
        ttk.Label(files, text="Drag/drop filter").pack(anchor="w", pady=(18, 4))
        self.drop_filter_var = tk.StringVar(value=str(self.config.get("DragDropFilterType", "None")))
        ttk.Combobox(files, textvariable=self.drop_filter_var, values=("None", "Inclusion", "Exclusion"), state="readonly", width=18).pack(anchor="w")
        ttk.Label(files, text="Filter masks").pack(anchor="w", pady=(12, 4))
        self.drop_mask_var = tk.StringVar(value=str(self.config.get("DragDropFilterMask", "")))
        ttk.Entry(files, textvariable=self.drop_mask_var).pack(fill="x")
        ttk.Label(files, text="Semicolon-separated masks, for example *.zip;*.7z", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        explorer = ttk.Frame(notebook, padding=16)
        notebook.add(explorer, text="Explorer")
        self.ctx_enabled_var = tk.BooleanVar(value=bool(self.config.get("CtxEnabled", False)))
        self.ctx_grouped_var = tk.BooleanVar(value=bool(self.config.get("CtxGrouped", True)))
        self.ctx_here_var = tk.BooleanVar(value=bool(self.config.get("CtxExtractHere", True)))
        self.ctx_folder_var = tk.BooleanVar(value=bool(self.config.get("CtxExtractToFolder", True)))
        self.ctx_enqueue_var = tk.BooleanVar(value=bool(self.config.get("CtxEnqueue", True)))
        self.ctx_search_var = tk.BooleanVar(value=bool(self.config.get("CtxSearchArchives", True)))
        ttk.Checkbutton(explorer, text="Add ExtractorX to Explorer context menus", variable=self.ctx_enabled_var).pack(anchor="w")
        ttk.Checkbutton(explorer, text="Group actions under an ExtractorX submenu", variable=self.ctx_grouped_var).pack(anchor="w")
        ttk.Checkbutton(explorer, text="Extract here", variable=self.ctx_here_var).pack(anchor="w")
        ttk.Checkbutton(explorer, text="Extract to folder", variable=self.ctx_folder_var).pack(anchor="w")
        ttk.Checkbutton(explorer, text="Add to ExtractorX", variable=self.ctx_enqueue_var).pack(anchor="w")
        ttk.Checkbutton(explorer, text="Search folders for archives", variable=self.ctx_search_var).pack(anchor="w")
        explorer_buttons = ttk.Frame(explorer)
        explorer_buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(explorer_buttons, text="Install Menus", command=self._install_context_menu).pack(side="left", padx=(0, 8))
        ttk.Button(explorer_buttons, text="Remove Menus", command=self._remove_context_menu).pack(side="left")
        ttk.Label(explorer, text="File associations", style="Muted.TLabel").pack(anchor="w", pady=(18, 4))
        self.file_assoc_var = tk.StringVar(value=";".join(str(ext) for ext in self.config.get("FileAssociations", [])))
        ttk.Entry(explorer, textvariable=self.file_assoc_var).pack(fill="x")
        ttk.Label(explorer, text="Semicolon-separated extensions, for example .zip;.7z;.rar", style="Muted.TLabel").pack(anchor="w", pady=(4, 8))
        assoc_buttons = ttk.Frame(explorer)
        assoc_buttons.pack(fill="x")
        ttk.Button(assoc_buttons, text="Register Associations", command=self._install_file_associations).pack(side="left", padx=(0, 8))
        ttk.Button(assoc_buttons, text="Remove Associations", command=self._remove_file_associations).pack(side="left")

        advanced = ttk.Frame(notebook, padding=16)
        notebook.add(advanced, text="Advanced")
        ttk.Label(advanced, text="Thread priority").pack(anchor="w")
        self.thread_priority_var = tk.StringVar(value=str(self.config.get("ThreadPriority", "Normal")))
        ttk.Combobox(advanced, textvariable=self.thread_priority_var, values=("Low", "BelowNormal", "Normal", "AboveNormal", "High"), state="readonly", width=18).pack(anchor="w", pady=(4, 12))
        parallel_frame = ttk.Frame(advanced)
        parallel_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(parallel_frame, text="Parallel extractions").pack(side="left")
        self.parallel_var = tk.IntVar(value=int(self.config.get("MaxParallelExtractions", 1)))
        ttk.Spinbox(parallel_frame, from_=1, to=8, textvariable=self.parallel_var, width=6).pack(side="left", padx=(8, 0))
        ttk.Label(parallel_frame, text="(1 = serial, higher = concurrent 7-Zip processes)", style="Muted.TLabel").pack(side="left", padx=(8, 0))
        bomb_frame = ttk.Frame(advanced)
        bomb_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(bomb_frame, text="Max decompression ratio").pack(side="left")
        self.decomp_ratio_var = tk.IntVar(value=int(self.config.get("MaxDecompressionRatio", 1000)))
        ttk.Spinbox(bomb_frame, from_=0, to=100000, textvariable=self.decomp_ratio_var, width=8).pack(side="left", padx=(8, 0))
        ttk.Label(bomb_frame, text="(0 = unlimited; warns on suspected zip bombs)", style="Muted.TLabel").pack(side="left", padx=(8, 0))
        backend_frame = ttk.Frame(advanced)
        backend_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(backend_frame, text="7-Zip override").pack(side="left")
        self.sevenzip_override_var = tk.StringVar(value=str(self.config.get("SevenZipOverride", "")))
        ttk.Entry(backend_frame, textvariable=self.sevenzip_override_var).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(backend_frame, text="Browse", command=self._browse_sevenzip).pack(side="left")
        ttk.Label(advanced, text="Leave empty to auto-detect. Point at 7-Zip ZS's 7zzs.exe or NanaZipC.exe for alternate backends.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(0, 12))
        ttk.Label(advanced, text="Lifecycle hooks (shell, tokens: {ArchivePath} {Output} {ArchiveName} {ExitCode})").pack(anchor="w", pady=(0, 4))
        self.pre_hook_var = tk.StringVar(value=str(self.config.get("PreExtractCommand", "")))
        self.post_hook_var = tk.StringVar(value=str(self.config.get("PostExtractCommand", "")))
        self.fail_hook_var = tk.StringVar(value=str(self.config.get("OnFailureCommand", "")))
        for label_text, var in (
            ("Pre-extract", self.pre_hook_var),
            ("Post-extract", self.post_hook_var),
            ("On-failure", self.fail_hook_var),
        ):
            row = ttk.Frame(advanced)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=label_text, width=14).pack(side="left")
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        ttk.Label(advanced, text="External processors").pack(anchor="w")
        ttk.Label(
            advanced,
            text=(
                "Use Extension|Command, one per line. Tokens: {ArchivePath}, {Destination}, "
                "{Output}, {ArchiveName}. Tokens are auto-quoted; do not wrap them in quotes."
            ),
            style="Muted.TLabel",
            wraplength=540,
            justify="left",
        ).pack(anchor="w", pady=(2, 8))
        self.external_text = tk.Text(advanced, height=9, bg=self.app.palette["surface_2"], fg=self.app.palette["text"], insertbackground=self.app.palette["text"], relief="flat", wrap="none")
        self.external_text.pack(fill="both", expand=True)
        for processor in self.config.get("ExternalProcessors", []):
            self.external_text.insert("end", f"{processor.get('Extension', '')}|{processor.get('Command', '')}\n")

        bookmarks_tab = ttk.Frame(notebook, padding=16)
        notebook.add(bookmarks_tab, text="Bookmarks")
        ttk.Label(bookmarks_tab, text="Bookmarks").pack(anchor="w")
        ttk.Label(bookmarks_tab, text="One per line as Label|Path. The toolbar Bookmarks menu lists these.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(2, 8))
        self.bookmarks_text = tk.Text(bookmarks_tab, height=10, bg=self.app.palette["surface_2"], fg=self.app.palette["text"], insertbackground=self.app.palette["text"], relief="flat", wrap="none")
        self.bookmarks_text.pack(fill="both", expand=True)
        for bookmark in self.config.get("Bookmarks", []) or []:
            label = str(bookmark.get("Label", "")).strip()
            path = str(bookmark.get("Path", "")).strip()
            if label and path:
                self.bookmarks_text.insert("end", f"{label}|{path}\n")

        footer = ttk.Frame(self.window, padding=(14, 0, 14, 14))
        footer.pack(fill="x")
        ttk.Button(footer, text="Cancel", command=self.window.destroy).pack(side="right")
        ttk.Button(footer, text="Save Settings", style="Accent.TButton", command=self._save).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="Reset to Defaults", command=self._reset_defaults).pack(side="left")

    def _add_watch_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.window, title="Choose watch folder")
        if folder:
            self.watch_list.insert("end", folder)

    def _remove_watch_folder(self) -> None:
        for index in reversed(self.watch_list.curselection()):
            self.watch_list.delete(index)

    def _browse_post_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.window, title="Choose post-action folder")
        if folder:
            self.post_folder_var.set(folder)

    def _browse_sevenzip(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.window,
            title="Choose 7-Zip executable",
            filetypes=(("Executables", "*.exe"), ("All files", "*.*")),
        )
        if path:
            self.sevenzip_override_var.set(path)

    def _read_bookmarks(self) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        for line in self.bookmarks_text.get("1.0", "end").splitlines():
            if "|" not in line:
                continue
            label, path = line.split("|", 1)
            label = label.strip()
            path = path.strip()
            if label and path:
                entries.append({"Label": label, "Path": path})
        return entries

    def _install_context_menu(self) -> None:
        self._capture_context_config()
        try:
            shell_integration.install_context_menu(self.config)
            self.ctx_enabled_var.set(True)
            messagebox.showinfo("Explorer menus installed", "ExtractorX was added to your Explorer context menus.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("Explorer menus failed", str(exc), parent=self.window)

    def _remove_context_menu(self) -> None:
        try:
            shell_integration.uninstall_context_menu()
            self.ctx_enabled_var.set(False)
            messagebox.showinfo("Explorer menus removed", "ExtractorX Explorer context menus were removed.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("Explorer menus failed", str(exc), parent=self.window)

    def _install_file_associations(self) -> None:
        extensions = self._read_file_associations()
        if not extensions:
            messagebox.showwarning("File associations", "Enter at least one extension before registering file associations.", parent=self.window)
            return
        try:
            shell_integration.install_file_associations(extensions)
            messagebox.showinfo("File associations registered", "Selected archive extensions now open with ExtractorX.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("File associations failed", str(exc), parent=self.window)

    def _remove_file_associations(self) -> None:
        extensions = self._read_file_associations()
        if not extensions:
            messagebox.showwarning("File associations", "Enter the extensions to remove.", parent=self.window)
            return
        try:
            shell_integration.uninstall_file_associations(extensions)
            messagebox.showinfo("File associations removed", "Selected ExtractorX file associations were removed.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("File associations failed", str(exc), parent=self.window)

    def _add_password(self) -> None:
        password = _ask_password_with_entropy(self.window, "Add Password", "Password")
        if password:
            self.app.passwords.append(password)
            self.password_list.insert("end", _mask_password(password))

    def _remove_password(self) -> None:
        for index in reversed(self.password_list.curselection()):
            self.password_list.delete(index)
            if 0 <= index < len(self.app.passwords):
                del self.app.passwords[index]

    def _import_passwords(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.window,
            title="Import password list",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*")),
        )
        if not path:
            return
        try:
            text = Path(path).read_text(encoding="utf-8-sig", errors="replace")
        except OSError as exc:
            messagebox.showerror("Import passwords", f"Could not read file: {exc}", parent=self.window)
            return
        imported = 0
        for line in text.splitlines():
            password = line.strip()
            if not password or password in self.app.passwords:
                continue
            self.app.passwords.append(password)
            self.password_list.insert("end", _mask_password(password))
            imported += 1
        messagebox.showinfo(
            "Import passwords",
            f"Imported {imported} password(s). Click Save Settings to persist.",
            parent=self.window,
        )

    def _save(self) -> None:
        self.config["Theme"] = self.theme_var.get()
        self.config["AlwaysOnTop"] = self.always_var.get()
        self.config["MinimizeToTray"] = self.minimize_tray_var.get()
        self.config["LogHistory"] = self.log_history_var.get()
        self.config["WatchAutoExtract"] = self.watch_auto_var.get()
        self.config["NestedExtraction"] = self.nested_var.get()
        self.config["SoundsEnabled"] = self.sounds_var.get()
        self.config["AutoExtractOnDrop"] = self.auto_drop_var.get()
        self.config["UsePasswordList"] = self.use_passwords_var.get()
        self.config["PromptOnExhaustion"] = self.prompt_password_var.get()
        self.config["AssumeOnePassword"] = self.assume_one_password_var.get()
        self.config["UsePasswordSidecars"] = self.use_sidecars_var.get()
        self.config["HashModePasswordProbe"] = self.hash_probe_var.get()
        self.config["WordlistGeneration"] = self.wordlist_var.get()
        self.config["WordlistMaxAttempts"] = int(self.wordlist_max_var.get() or 500)
        self.config["PasswordTimeout"] = self.password_timeout_var.get()
        self.config["NestedMaxDepth"] = self.nested_depth_var.get()
        self.config["NestedApplyPostAction"] = self.nested_post_var.get()
        self.config["OutputPath"] = self.output_var.get().strip()
        self.config["OverwriteMode"] = self.overwrite_var.get()
        self.config["DeleteAfterExtract"] = self.delete_after_var.get()
        self.config["PostAction"] = self.post_action_var.get()
        self.config["PostActionFolder"] = self.post_folder_var.get()
        self.config["OpenDestAfterExtract"] = self.open_dest_var.get()
        self.config["RemoveDuplicateFolder"] = self.remove_dupe_var.get()
        self.config["RenameSingleFile"] = self.rename_single_var.get()
        self.config["DeleteBrokenFiles"] = self.delete_broken_var.get()
        self.config["ClearListOnComplete"] = self.clear_done_var.get()
        self.config["CloseOnComplete"] = self.close_done_var.get()
        self.config["FileExclusions"] = self.exclusions_var.get()
        self.config["DragDropFilterType"] = self.drop_filter_var.get()
        self.config["DragDropFilterMask"] = self.drop_mask_var.get()
        self.config["ThreadPriority"] = self.thread_priority_var.get()
        self.config["WatchFolders"] = list(self.watch_list.get(0, "end"))
        self.config["SmartExtract"] = self.smart_extract_var.get()
        self.config["FilenameEncoding"] = self.filename_encoding_var.get()
        self.config["IncludeMasks"] = self.include_masks_var.get().strip()
        self.config["SkipAfterFailedPasswords"] = int(self.skip_passwords_var.get() or 0)
        self.config["MaxParallelExtractions"] = int(self.parallel_var.get() or 1)
        self.config["MaxDecompressionRatio"] = int(self.decomp_ratio_var.get() or 1000)
        self.config["SevenZipOverride"] = self.sevenzip_override_var.get().strip()
        self.config["PreExtractCommand"] = self.pre_hook_var.get().strip()
        self.config["PostExtractCommand"] = self.post_hook_var.get().strip()
        self.config["OnFailureCommand"] = self.fail_hook_var.get().strip()
        self.config["HandlerAllowlist"] = [
            part.strip().lower().lstrip(".")
            for part in re.split(r"[;,\s]+", self.handler_allowlist_var.get())
            if part.strip()
        ]
        self.config["Bookmarks"] = self._read_bookmarks()
        self._capture_context_config()
        previous_assocs = list(self.app.config.get("FileAssociations", []))
        self.config["FileAssociations"] = self._read_file_associations()
        self.config["ExternalProcessors"] = self._read_external_processors()
        self.app.config.clear()
        self.app.config.update(save_config(self.config))
        removed_assocs = [ext for ext in previous_assocs if ext not in self.app.config.get("FileAssociations", [])]
        self.app.output_template_var.set(str(self.app.config.get("OutputPath", "")))
        self.app.password_store.save(self.app.passwords)
        self.app.root.attributes("-topmost", bool(self.app.config.get("AlwaysOnTop", False)))
        self.app._apply_theme(str(self.app.config.get("Theme", "Midnight")))
        from .sevenzip import find_7zip as _find_7zip

        resolved = _find_7zip(override=self.app.config.get("SevenZipOverride"))
        if resolved is not None:
            self.app.sevenzip_path = resolved
            self.app.service.sevenzip_path = resolved
        self.app._start_watchers()
        try:
            if self.app.config.get("CtxEnabled"):
                shell_integration.install_context_menu(self.app.config)
            else:
                shell_integration.uninstall_context_menu()
            if removed_assocs:
                shell_integration.uninstall_file_associations(removed_assocs)
            if self.app.config.get("FileAssociations"):
                shell_integration.install_file_associations(list(self.app.config.get("FileAssociations", [])))
        except Exception as exc:
            messagebox.showwarning("Explorer menus", f"Settings were saved, but Explorer menus could not be updated: {exc}", parent=self.window)
        messagebox.showinfo("Settings saved", "Settings were saved.", parent=self.window)
        self.window.destroy()

    def _reset_defaults(self) -> None:
        if not messagebox.askyesno(
            "Reset to Defaults",
            "Restore all settings to their factory defaults? Window size, password list, and watch folders will be preserved.",
            parent=self.window,
        ):
            return
        preserved = {
            "WindowWidth": self.app.config.get("WindowWidth"),
            "WindowHeight": self.app.config.get("WindowHeight"),
            "WindowLeft": self.app.config.get("WindowLeft"),
            "WindowTop": self.app.config.get("WindowTop"),
            "WatchFolders": list(self.app.config.get("WatchFolders", [])),
        }
        fresh = dict(DEFAULT_CONFIG)
        fresh.update({key: value for key, value in preserved.items() if value is not None})
        self.app.config.clear()
        self.app.config.update(save_config(fresh))
        self.app.output_template_var.set(str(self.app.config.get("OutputPath", "")))
        self.app._apply_theme(str(self.app.config.get("Theme", "Midnight")))
        self.app._start_watchers()
        messagebox.showinfo("Settings reset", "Defaults restored. Reopen Settings to review.", parent=self.window)
        self.window.destroy()

    def _capture_context_config(self) -> None:
        self.config["CtxEnabled"] = self.ctx_enabled_var.get()
        self.config["CtxGrouped"] = self.ctx_grouped_var.get()
        self.config["CtxExtractHere"] = self.ctx_here_var.get()
        self.config["CtxExtractToFolder"] = self.ctx_folder_var.get()
        self.config["CtxEnqueue"] = self.ctx_enqueue_var.get()
        self.config["CtxSearchArchives"] = self.ctx_search_var.get()

    def _read_external_processors(self) -> list[dict[str, str]]:
        processors: list[dict[str, str]] = []
        for line in self.external_text.get("1.0", "end").splitlines():
            if "|" not in line:
                continue
            extension, command = line.split("|", 1)
            extension = extension.strip().lstrip(".")
            command = command.strip()
            if extension and command:
                processors.append({"Extension": extension, "Command": command})
        return processors

    def _read_file_associations(self) -> list[str]:
        values = re.split(r"[;,\s]+", self.file_assoc_var.get())
        extensions: list[str] = []
        for value in values:
            value = value.strip().lower()
            if not value:
                continue
            if not value.startswith("."):
                value = "." + value
            if value not in extensions:
                extensions.append(value)
        return extensions


def _mask_password(password: str) -> str:
    if len(password) <= 4:
        return "*" * len(password)
    return password[:2] + "*" * (len(password) - 4) + password[-2:]


def _ask_password_with_entropy(parent: tk.Misc, title: str, prompt: str) -> str | None:
    window = tk.Toplevel(parent)
    window.title(title)
    window.transient(parent)
    window.grab_set()
    window.resizable(False, False)
    result: dict[str, str | None] = {"value": None}
    ttk.Label(window, text=prompt).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(16, 4))
    var = tk.StringVar()
    entry = ttk.Entry(window, textvariable=var, show="*", width=32)
    entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16)
    meter = ttk.Label(window, text="Strength: weak (0 bits)")
    meter.grid(row=2, column=0, columnspan=2, sticky="w", padx=16, pady=(4, 8))

    def update_meter(*_: object) -> None:
        label, bits = classify_entropy(var.get())
        meter.configure(text=f"Strength: {label} ({bits:.0f} bits)")

    var.trace_add("write", update_meter)

    def commit() -> None:
        result["value"] = var.get()
        window.destroy()

    buttons = ttk.Frame(window)
    buttons.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 16))
    ttk.Button(buttons, text="Cancel", command=window.destroy).pack(side="right")
    ttk.Button(buttons, text="OK", command=commit).pack(side="right", padx=(0, 8))
    window.bind("<Return>", lambda _event: commit())
    window.bind("<Escape>", lambda _event: window.destroy())
    entry.focus_set()
    parent.wait_window(window)
    return result["value"]


def _looks_like_password_failure(text: str) -> bool:
    lowered = text.lower()
    return any(token in lowered for token in ("password", "encrypted", "wrong password", "data error"))


def _open_path(path: Path) -> None:
    if path.exists():
        if hasattr(os, "startfile"):
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(path)])


def _reveal_path(path: Path) -> None:
    if not path.exists():
        return
    if os.name == "nt":
        # explorer /select, needs a single quoted argument. Using a list avoids
        # shell interpretation but explorer.exe expects the comma prefix to
        # sit next to the path, which Popen's list form preserves.
        subprocess.Popen(["explorer.exe", f"/select,{path}"])
    else:
        _open_path(path.parent)


def _cmd_quote(value: str) -> str:
    """Quote ``value`` for use in a Windows ``cmd.exe`` batch script.

    Wraps in double quotes and escapes embedded quotes. Caret-escaped meta
    characters inside the value are left alone -- users who embed them
    intentionally (rare) can post-process the output.
    """
    return '"' + value.replace('"', '""') + '"'


def _is_drive_root(path: Path) -> bool:
    """Return True for paths like ``C:\\``, ``/``, or a UNC share root."""
    try:
        resolved = path.resolve(strict=False)
    except OSError:
        resolved = path
    return resolved == resolved.parent
