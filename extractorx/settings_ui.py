"""Settings dialog extracted from ui.py to keep the main module focused."""

from __future__ import annotations

import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import (
    DEFAULT_CONFIG,
    FILENAME_ENCODINGS,
    SMART_EXTRACT_MODES,
    THEMES as THEME_NAMES,
    save_config,
)
from .passwords import classify_entropy
from . import shell_integration

TYPE_CHECKING = False
if TYPE_CHECKING:
    from .ui import ExtractorXApp


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
        ttk.Label(destination, text="7-Zip output path mode").pack(anchor="w", pady=(16, 4))
        self.output_path_mode_var = tk.StringVar(value=str(self.config.get("OutputPathMode", "")))
        ttk.Combobox(destination, textvariable=self.output_path_mode_var, values=("", "d", "c", "r"), state="readonly", width=20).pack(anchor="w")
        ttk.Label(destination, text="d=subfolder, c=current folder, r=archive folder. Empty uses ExtractorX output template.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(2, 8))
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
        self.motw_var = tk.BooleanVar(value=bool(self.config.get("PropagateMotw", True)))
        self.secure_delete_var = tk.BooleanVar(value=bool(self.config.get("SecureDelete", False)))
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
        ttk.Checkbutton(process, text="Propagate Mark-of-the-Web to extracted files", variable=self.motw_var).grid(row=9, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Secure delete (overwrite with zeros before removing)", variable=self.secure_delete_var).grid(row=10, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(process, text="Clear successful items when a batch completes", variable=self.clear_done_var).grid(row=11, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Checkbutton(process, text="Close when the full batch succeeds", variable=self.close_done_var).grid(row=12, column=0, columnspan=3, sticky="w")
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
        current_assocs = set(str(ext).lower().lstrip(".") for ext in self.config.get("FileAssociations", []))
        assoc_grid = ttk.Frame(explorer)
        assoc_grid.pack(fill="x", pady=(0, 4))
        self._assoc_vars: dict[str, tk.BooleanVar] = {}
        common_exts = ["zip", "7z", "rar", "tar", "gz", "bz2", "xz", "zst", "iso", "cab"]
        for i, ext in enumerate(common_exts):
            var = tk.BooleanVar(value=ext in current_assocs)
            self._assoc_vars[ext] = var
            ttk.Checkbutton(assoc_grid, text=f".{ext}", variable=var, command=self._sync_assoc_text).grid(row=i // 5, column=i % 5, sticky="w", padx=4)
        self.file_assoc_var = tk.StringVar(value=";".join(str(ext) for ext in self.config.get("FileAssociations", [])))
        ttk.Entry(explorer, textvariable=self.file_assoc_var).pack(fill="x")
        ttk.Label(explorer, text="Toggle common formats above, or edit the full list directly", style="Muted.TLabel").pack(anchor="w", pady=(4, 8))
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
        mem_frame = ttk.Frame(advanced)
        mem_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(mem_frame, text="Max memory for RAR (GB)").pack(side="left")
        self.max_memory_var = tk.IntVar(value=int(self.config.get("MaxMemoryGB", 0)))
        ttk.Spinbox(mem_frame, from_=0, to=64, textvariable=self.max_memory_var, width=4).pack(side="left", padx=(8, 0))
        ttk.Label(mem_frame, text="(0 = unlimited; caps 7-Zip RAM via -smemx)", style="Muted.TLabel").pack(side="left", padx=(8, 0))
        retry_frame = ttk.Frame(advanced)
        retry_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(retry_frame, text="Retry failed extractions").pack(side="left")
        self.retry_count_var = tk.IntVar(value=int(self.config.get("RetryCount", 0)))
        ttk.Spinbox(retry_frame, from_=0, to=10, textvariable=self.retry_count_var, width=4).pack(side="left", padx=(8, 0))
        ttk.Label(retry_frame, text="times, delay").pack(side="left", padx=(8, 0))
        self.retry_delay_var = tk.IntVar(value=int(self.config.get("RetryDelaySeconds", 30)))
        ttk.Spinbox(retry_frame, from_=5, to=600, textvariable=self.retry_delay_var, width=5).pack(side="left", padx=(4, 0))
        ttk.Label(retry_frame, text="sec (0 retries = disabled)", style="Muted.TLabel").pack(side="left", padx=(8, 0))
        backend_frame = ttk.Frame(advanced)
        backend_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(backend_frame, text="7-Zip override").pack(side="left")
        self.sevenzip_override_var = tk.StringVar(value=str(self.config.get("SevenZipOverride", "")))
        ttk.Entry(backend_frame, textvariable=self.sevenzip_override_var).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(backend_frame, text="Browse", command=self._browse_sevenzip).pack(side="left")
        ttk.Label(advanced, text="Leave empty to auto-detect. Point at 7-Zip ZS's 7zzs.exe or NanaZipC.exe for alternate backends.", style="Muted.TLabel", wraplength=540, justify="left").pack(anchor="w", pady=(0, 12))
        self.block_outdated_var = tk.BooleanVar(value=bool(self.config.get("BlockOutdated7Zip", True)))
        ttk.Checkbutton(advanced, text="Block extraction when 7-Zip is below minimum version (security)", variable=self.block_outdated_var).pack(anchor="w", pady=(0, 12))
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

    def _sync_assoc_text(self) -> None:
        current = set(
            v.strip().lower().lstrip(".")
            for v in re.split(r"[;,\s]+", self.file_assoc_var.get())
            if v.strip()
        )
        for ext, var in self._assoc_vars.items():
            if var.get():
                current.add(ext)
            else:
                current.discard(ext)
        self.file_assoc_var.set(";".join(f".{e}" for e in sorted(current) if e))

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
        self.config["OutputPathMode"] = self.output_path_mode_var.get()
        self.config["DeleteAfterExtract"] = self.delete_after_var.get()
        self.config["PostAction"] = self.post_action_var.get()
        self.config["PostActionFolder"] = self.post_folder_var.get()
        self.config["OpenDestAfterExtract"] = self.open_dest_var.get()
        self.config["RemoveDuplicateFolder"] = self.remove_dupe_var.get()
        self.config["RenameSingleFile"] = self.rename_single_var.get()
        self.config["DeleteBrokenFiles"] = self.delete_broken_var.get()
        self.config["PropagateMotw"] = self.motw_var.get()
        self.config["SecureDelete"] = self.secure_delete_var.get()
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
        self.config["MaxMemoryGB"] = int(self.max_memory_var.get() or 0)
        self.config["RetryCount"] = int(self.retry_count_var.get() or 0)
        self.config["RetryDelaySeconds"] = int(self.retry_delay_var.get() or 30)
        self.config["SevenZipOverride"] = self.sevenzip_override_var.get().strip()
        self.config["BlockOutdated7Zip"] = self.block_outdated_var.get()
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
