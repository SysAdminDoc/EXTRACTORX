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
from .repack import REPACK_FORMATS, repack_archive
from .settings_ui import SettingsDialog, _ask_password_with_entropy, _mask_password
from .sevenzip import list_archive_contents
from .config import app_data_dir
from .config import save_config
from .download import check_for_updates, looks_like_url
from .extractor import ExtractionService
from .identify import identify
from .models import OperationMessage, QueueItem, QueueStatus
from .passwords import PasswordStore
from . import shell_integration
from .themes import get_theme
from .watcher import WatchService
from .windows_integration import (
    TaskbarProgress,
    TBPF_ERROR,
    TBPF_NORMAL,
    WindowsShellBridge,
    detect_system_theme,
    set_dark_titlebar,
)


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
        if not config.get("_theme_set_by_user"):
            system_theme = detect_system_theme()
            if system_theme == "light" and config.get("Theme") == "Midnight":
                config["Theme"] = "White"
        self.palette = get_theme(config.get("Theme"))
        self.style = ttk.Style(self.root)
        self._configure_style()
        self._build_ui()
        self.root.update_idletasks()
        is_dark = config.get("Theme") not in ("White",)
        set_dark_titlebar(int(self.root.wm_frame(), 16), dark=is_dark)
        self.shell_bridge = WindowsShellBridge(
            self.root,
            on_files_dropped=self._handle_dropped_paths,
            on_tray_restore=self._restore_from_tray,
            on_tray_menu=self._show_tray_menu,
        )
        self.shell_bridge.enable_drag_drop()
        self.root.update_idletasks()
        self.taskbar_progress = TaskbarProgress(int(self.root.wm_frame(), 16))
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
        try:
            is_dark = theme_name not in ("White",)
            set_dark_titlebar(int(self.root.wm_frame(), 16), dark=is_dark)
        except (tk.TclError, ValueError):
            pass
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
        ttk.Button(toolbar, text="Diff", command=self._diff_archives).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Search", command=self._search_archives).pack(side="left", padx=(0, 8))
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
        self.tree.bind("<Delete>", lambda _e: self._remove_selected())
        self.tree.bind("<Return>", lambda _e: self._extract())
        self.root.bind("<Control-o>", lambda _e: self._add_files())
        self.root.bind("<Control-e>", lambda _e: self._extract())
        self.root.bind("<Control-s>", lambda _e: self._export_batch_script())
        self.root.bind("<Escape>", lambda _e: self._stop_extraction())
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
        self.log_menu.add_command(label="Export History...", command=self._export_history)
        self.log_menu.add_command(label="Export Log...", command=self._export_log)
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
        self.queue_menu.add_command(label="Preview Contents", command=self._preview_selected_contents)
        convert_menu = tk.Menu(self.queue_menu, tearoff=False, bg=p["surface_2"], fg=p["text"], activebackground=p["selection"], activeforeground=p["text"])
        for fmt_label in ("ZIP", "7z", "TAR"):
            convert_menu.add_command(label=fmt_label, command=lambda f=fmt_label.lower(): self._convert_selected(f))
        self.queue_menu.add_cascade(label="Convert to...", menu=convert_menu)
        self.queue_menu.add_command(label="Hex View", command=self._hex_view_selected)
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
        from .sevenzip import build_sevenzip_command
        from .archive import resolve_output_path
        lines = ["@echo off", f"REM ExtractorX v{__version__} — 7-Zip batch export", ""]
        for item in self.items.values():
            output = (
                Path(item.output_override).expanduser()
                if item.output_override
                else resolve_output_path(str(self.config.get("OutputPath", "")), item.archive_path)
            )
            try:
                cmd = build_sevenzip_command(
                    sevenzip_path=self.sevenzip_path,
                    archive=item.archive_path,
                    output=output,
                    overwrite_mode=str(self.config.get("OverwriteMode", "Always")),
                    exclusions=str(self.config.get("FileExclusions", "")),
                    inclusions=str(self.config.get("IncludeMasks", "")),
                    filename_encoding=str(self.config.get("FilenameEncoding", "Auto")),
                )
            except ValueError:
                cmd = ["7z", "x", str(item.archive_path), f"-o{output}", "-y"]
            lines.append(" ".join(_cmd_quote(arg) for arg in cmd))
        try:
            Path(target).write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
            self._log(f"Exported {len(self.items)} item(s) as 7z commands to {target}", "success")
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

    def _diff_archives(self) -> None:
        paths = filedialog.askopenfilenames(title="Select two archives to compare", parent=self.root)
        if not paths or len(paths) != 2:
            if paths and len(paths) != 2:
                self._log("Select exactly two archives to compare.", "warning")
            return
        if not self.sevenzip_path:
            self._log("7-Zip not found.", "error")
            return
        self._log(f"Comparing {Path(paths[0]).name} vs {Path(paths[1]).name}...", "info")
        entries_a = list_archive_contents(self.sevenzip_path, paths[0])
        entries_b = list_archive_contents(self.sevenzip_path, paths[1])
        set_a = {e["Path"]: e.get("Size", "") for e in entries_a}
        set_b = {e["Path"]: e.get("Size", "") for e in entries_b}
        all_paths = sorted(set(set_a) | set(set_b))
        window = tk.Toplevel(self.root)
        window.title(f"Diff: {Path(paths[0]).name} vs {Path(paths[1]).name}")
        window.transient(self.root)
        window.geometry("750x500")
        window.configure(bg=self.palette["bg"])
        tree = ttk.Treeview(window, columns=("path", "status", "size_a", "size_b"), show="headings")
        tree.heading("path", text="Path")
        tree.heading("status", text="Status")
        tree.heading("size_a", text=Path(paths[0]).name)
        tree.heading("size_b", text=Path(paths[1]).name)
        tree.column("path", width=350, anchor="w")
        tree.column("status", width=100, anchor="w")
        tree.column("size_a", width=100, anchor="e")
        tree.column("size_b", width=100, anchor="e")
        scroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        tree.tag_configure("added", foreground=self.palette["ok"])
        tree.tag_configure("removed", foreground=self.palette["error"])
        tree.tag_configure("changed", foreground=self.palette["warn"])
        added = removed = changed = same = 0
        for p in all_paths:
            in_a = p in set_a
            in_b = p in set_b
            if in_a and not in_b:
                tree.insert("", "end", values=(p, "Removed", set_a[p], ""), tags=("removed",))
                removed += 1
            elif in_b and not in_a:
                tree.insert("", "end", values=(p, "Added", "", set_b[p]), tags=("added",))
                added += 1
            elif set_a[p] != set_b[p]:
                tree.insert("", "end", values=(p, "Changed", set_a[p], set_b[p]), tags=("changed",))
                changed += 1
            else:
                same += 1
        footer = ttk.Frame(window, padding=(8, 4))
        footer.pack(fill="x")
        ttk.Label(footer, text=f"{added} added, {removed} removed, {changed} changed, {same} unchanged", style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="Close", command=window.destroy).pack(side="right")
        self._log(f"Diff: {added} added, {removed} removed, {changed} changed.", "info")

    def _search_archives(self) -> None:
        if not self.items:
            self._log("Queue is empty — nothing to search.", "warning")
            return
        pattern = simpledialog.askstring("Search Archives", "Filename pattern (e.g. *.txt, readme*):", parent=self.root)
        if not pattern:
            return
        self._log(f"Searching queued archives for '{pattern}'...", "info")
        import threading
        def do_search() -> None:
            matches = 0
            for item in list(self.items.values()):
                if not item.archive_path.exists():
                    continue
                entries = list_archive_contents(self.sevenzip_path, item.archive_path)
                for entry in entries:
                    entry_path = entry.get("Path", "")
                    if fnmatch.fnmatch(entry_path.lower(), pattern.lower()) or fnmatch.fnmatch(Path(entry_path).name.lower(), pattern.lower()):
                        self.root.after(0, lambda a=item.archive_path.name, p=entry_path: self._log(f"  {a}: {p}", "success"))
                        matches += 1
            self.root.after(0, lambda: self._log(f"Search complete: {matches} match(es) found.", "info"))
        threading.Thread(target=do_search, name="ExtractorXSearch", daemon=True).start()

    def _show_about(self) -> None:
        sevenzip = str(self.sevenzip_path) if self.sevenzip_path else "not found (will download on first extract)"
        messagebox.showinfo(
            "About ExtractorX",
            f"ExtractorX v{__version__}\n\nBulk archive extraction for Windows with queueing, monitoring, "
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
        import threading
        threading.Thread(
            target=lambda: check_for_updates(__version__, log=lambda text, level="info": self.root.after(0, lambda: self._log(text, level))),
            name="ExtractorXUpdateCheck",
            daemon=True,
        ).start()

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

    def _watch_rule_output(self, archive_path: Path) -> str | None:
        for rule in self.config.get("WatchFolderRules", []) or []:
            folder = str(rule.get("Folder", "")).strip()
            if not folder:
                continue
            try:
                rule_folder = Path(folder).resolve()
                archive_parent = archive_path.parent.resolve()
                if archive_parent == rule_folder or str(archive_parent).startswith(str(rule_folder) + os.sep):
                    output_template = rule.get("OutputPath", "")
                    if output_template:
                        return str(resolve_output_path(output_template, archive_path))
            except OSError:
                continue
        return None

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
            output_override = self._watch_rule_output(path)
            item = QueueItem.from_path(path, output_override=output_override)
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
        if message.type == "sub_progress":
            pct = int(message.payload.get("percent", 0) or 0)
            self.progress.configure(value=pct)
            self.taskbar_progress.set_state(TBPF_NORMAL)
            self.taskbar_progress.set_progress(pct, 100)
            return
        if message.type == "progress":
            total = int(message.payload.get("total", 0) or 0)
            current = int(message.payload.get("current", 0) or 0)
            self.progress.configure(value=(current / total * 100) if total else 0)
            self.status_label.configure(text=message.text)
            if self.shell_bridge:
                self.shell_bridge.update_tray_tip(f"ExtractorX - {message.text}")
            self.taskbar_progress.set_state(TBPF_NORMAL)
            self.taskbar_progress.set_progress(current, total)
        if message.type == "extract_done":
            self.progress.configure(value=100)
            self.status_label.configure(text="7-Zip ready" if self.sevenzip_path else "7-Zip not found")
            failed = any(item.status == QueueStatus.FAILED for item in self.items.values())
            if failed:
                self.taskbar_progress.set_state(TBPF_ERROR)
                self.taskbar_progress.set_progress(1, 1)
            else:
                self.taskbar_progress.clear()
            self._play_completion_sound()
            if bool(message.payload.get("test_only")):
                self.root.after(200, self._show_test_results)
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
        stamp = datetime.now().strftime("%H:%M:%S")
        tag = level if level in {"info", "success", "warning", "error"} else "info"
        self.log.insert("end", f"[{stamp}] {text}\n", tag)
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

    def _convert_selected(self, target_format: str) -> None:
        items = self._selected_items()[:1]
        if not items:
            self._log("Select an archive to convert.", "warning")
            return
        item = items[0]
        if not item.archive_path.exists():
            self._log("Archive no longer exists.", "warning")
            return
        if not self.sevenzip_path:
            self._log("7-Zip not found — cannot convert.", "error")
            return
        self._log(f"Converting {item.archive_path.name} to {target_format}...", "info")
        import threading
        def do_convert() -> None:
            result = repack_archive(
                self.sevenzip_path,
                item.archive_path,
                target_format,
                log_cb=lambda text, level="info": self.root.after(0, lambda: self._log(text, level)),
            )
            if result:
                self.root.after(0, lambda: self.add_item(QueueItem.from_path(result)))
        threading.Thread(target=do_convert, name="ExtractorXRepack", daemon=True).start()

    def _preview_selected_contents(self) -> None:
        items = self._selected_items()[:1]
        if not items:
            self._log("Select an archive to preview.", "warning")
            return
        item = items[0]
        if not item.archive_path.exists():
            self._log("Archive no longer exists.", "warning")
            return
        entries = list_archive_contents(self.sevenzip_path, item.archive_path)
        if not entries:
            self._log(f"No contents found in {item.archive_path.name}.", "warning")
            return
        window = tk.Toplevel(self.root)
        window.title(f"Contents: {item.archive_path.name}")
        window.transient(self.root)
        window.geometry("700x500")
        window.configure(bg=self.palette["bg"])
        tree = ttk.Treeview(window, columns=("path", "size", "modified"), show="headings", selectmode="extended")
        tree.heading("path", text="Path")
        tree.heading("size", text="Size")
        tree.heading("modified", text="Modified")
        tree.column("path", width=420, anchor="w")
        tree.column("size", width=100, anchor="e")
        tree.column("modified", width=160, anchor="w")
        scroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        def populate(flat: bool = False) -> None:
            tree.delete(*tree.get_children())
            display_entries = sorted(entries, key=lambda e: e.get("Path", "").lower()) if flat else entries
            for entry in display_entries:
                entry_path = entry.get("Path", "")
                display = Path(entry_path).name if flat else entry_path
                size_str = entry.get("Size", "")
                try:
                    size_str = format_size(int(size_str)) if size_str else ""
                except ValueError:
                    pass
                tree.insert("", "end", values=(display, size_str, entry.get("Modified", "")))

        populate()
        controls = ttk.Frame(window)
        controls.pack(fill="x", padx=8, pady=(0, 8))
        flat_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(controls, text="Flat view", variable=flat_var, command=lambda: populate(flat_var.get())).pack(side="left")
        ttk.Label(controls, text=f"{len(entries)} item(s)", style="Muted.TLabel").pack(side="right")

        def extract_selected_files() -> None:
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("No selection", "Select files to extract.", parent=window)
                return
            dest = filedialog.askdirectory(title="Extract selected files to...", parent=window)
            if not dest:
                return
            paths_to_extract = []
            for iid in selected:
                vals = tree.item(iid, "values")
                if vals:
                    paths_to_extract.append(vals[0])
            if not paths_to_extract or not self.sevenzip_path:
                return
            import subprocess as sp
            cmd = [str(self.sevenzip_path), "x", str(item.archive_path), f"-o{dest}", "-y"]
            for p in paths_to_extract:
                cmd.append(f"-i!{p}")
            try:
                result = sp.run(
                    cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                    timeout=3600,
                    creationflags=sp.CREATE_NO_WINDOW if hasattr(sp, "CREATE_NO_WINDOW") else 0,
                    check=False,
                )
                if result.returncode in {0, 1}:
                    self._log(f"Extracted {len(paths_to_extract)} file(s) to {dest}", "success")
                else:
                    self._log(f"Selective extraction failed (exit {result.returncode})", "error")
            except (OSError, sp.TimeoutExpired) as exc:
                self._log(f"Selective extraction failed: {exc}", "error")

        ttk.Button(controls, text="Extract Selected", command=extract_selected_files).pack(side="right", padx=(8, 0))
        self._log(f"Previewed {len(entries)} item(s) in {item.archive_path.name}.", "info")

    def _hex_view_selected(self) -> None:
        items = self._selected_items()[:1]
        if not items:
            self._log("Select an archive for hex view.", "warning")
            return
        self._show_hex_view(items[0].archive_path)

    def _show_hex_view(self, file_path: Path) -> None:
        if not file_path.exists():
            self._log("File no longer exists.", "warning")
            return
        try:
            data = file_path.read_bytes()[:4096]
        except OSError as exc:
            self._log(f"Could not read file: {exc}", "error")
            return
        window = tk.Toplevel(self.root)
        window.title(f"Hex: {file_path.name}")
        window.transient(self.root)
        window.geometry("720x450")
        window.configure(bg=self.palette["bg"])
        text = tk.Text(window, bg=self.palette["chrome"], fg=self.palette["text"],
                       insertbackground=self.palette["text"], relief="flat",
                       wrap="none", font=("Consolas", 9))
        scroll = ttk.Scrollbar(window, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        for offset in range(0, len(data), 16):
            chunk = data[offset:offset + 16]
            hex_part = " ".join(f"{b:02X}" for b in chunk).ljust(48)
            ascii_part = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
            text.insert("end", f"{offset:08X}  {hex_part}  {ascii_part}\n")
        if len(data) >= 4096:
            text.insert("end", f"\n... (showing first 4096 bytes of {file_path.stat().st_size})")
        text.configure(state="disabled")

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

    def _show_test_results(self) -> None:
        tested = [
            item for item in self.items.values()
            if item.status in {QueueStatus.TEST_OK, QueueStatus.FAILED} and item.test_detail
        ]
        if not tested:
            return
        window = tk.Toplevel(self.root)
        window.title("Test Results")
        window.transient(self.root)
        window.geometry("650x400")
        window.configure(bg=self.palette["bg"])
        tree = ttk.Treeview(window, columns=("archive", "result", "detail"), show="headings", selectmode="extended")
        tree.heading("archive", text="Archive")
        tree.heading("result", text="Result")
        tree.heading("detail", text="Detail")
        tree.column("archive", width=250, anchor="w")
        tree.column("result", width=80, anchor="w")
        tree.column("detail", width=300, anchor="w")
        scroll = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        ok_count = 0
        fail_count = 0
        for item in tested:
            result = "OK" if item.status == QueueStatus.TEST_OK else "FAIL"
            if item.status == QueueStatus.TEST_OK:
                ok_count += 1
            else:
                fail_count += 1
            tag = "done" if item.status == QueueStatus.TEST_OK else "failed"
            tree.insert("", "end", values=(item.archive_path.name, result, item.test_detail), tags=(tag,))
        tree.tag_configure("done", foreground=self.palette["ok"])
        tree.tag_configure("failed", foreground=self.palette["error"])
        footer = ttk.Frame(window, padding=(8, 4))
        footer.pack(fill="x")
        ttk.Label(footer, text=f"{ok_count} OK, {fail_count} failed", style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="Close", command=window.destroy).pack(side="right")

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

    def _export_history(self) -> None:
        if not self.items:
            self._log("No items to export.", "warning")
            return
        target = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export extraction history",
            defaultextension=".csv",
            filetypes=(("CSV", "*.csv"), ("JSON", "*.json"), ("All", "*.*")),
        )
        if not target:
            return
        import csv, json
        items_data = [
            {
                "archive": str(item.archive_path),
                "destination": str(item.output_path or ""),
                "status": item.status.value,
                "size": item.size_bytes,
                "error": item.error,
            }
            for item in self.items.values()
        ]
        try:
            if target.endswith(".json"):
                Path(target).write_text(json.dumps(items_data, indent=2), encoding="utf-8")
            else:
                with open(target, "w", newline="", encoding="utf-8") as f:
                    writer = csv.DictWriter(f, fieldnames=["archive", "destination", "status", "size", "error"])
                    writer.writeheader()
                    writer.writerows(items_data)
            self._log(f"Exported {len(items_data)} item(s) to {target}", "success")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc), parent=self.root)

    def _export_log(self) -> None:
        content = self.log.get("1.0", "end").strip()
        if not content:
            self._log("Log is empty.", "warning")
            return
        target = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export log",
            defaultextension=".txt",
            filetypes=(("Text files", "*.txt"), ("CSV", "*.csv"), ("All", "*.*")),
        )
        if not target:
            return
        try:
            if target.endswith(".csv"):
                import csv
                with open(target, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["timestamp", "level", "message"])
                    for line in content.splitlines():
                        if line.startswith("[") and "] " in line:
                            ts = line[1:line.index("]")]
                            rest = line[line.index("] ") + 2:]
                            writer.writerow([ts, "", rest])
                        else:
                            writer.writerow(["", "", line])
            else:
                Path(target).write_text(content, encoding="utf-8")
            self._log(f"Log exported to {target}", "success")
        except OSError as exc:
            messagebox.showerror("Export failed", str(exc), parent=self.root)

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
