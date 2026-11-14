#!/usr/bin/env python3
"""
A small desktop app for turning PowerPoint decks into PDFs.

Drag files (or folders) onto the window and they start converting.
Built on top of :mod:`ppt_to_pdf`, so it uses the same PowerPoint /
LibreOffice backends.

Run with::

    python gui.py
"""

from __future__ import annotations

import os
import platform
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ppt_to_pdf import (
    SUPPORTED_SUFFIXES,
    convert,
    describe_backend,
    find_libreoffice,
)

try:  # Drag & drop is optional; the app still works with the buttons.
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - depends on environment
    DND_FILES = None
    TkinterDnD = None


APP_NAME = "PowerPoint → PDF"

BG = "#f4f5f7"
CARD = "#ffffff"
BORDER = "#e4e7ec"
TEXT = "#101828"
MUTED = "#667085"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
ACCENT_SOFT = "#eff4ff"
OK = "#067647"
ERR = "#d92d20"
WARN = "#b54708"
ROW_ALT = "#fafbfc"
SELECT = "#dbe6fe"


@dataclass
class Item:
    """One file in the queue."""

    path: Path
    iid: str
    status: str = "queued"
    output: Path | None = None
    detail: str = ""
    error: str = ""
    seconds: float = 0.0


STATUS_TEXT = {
    "queued": "Queued",
    "working": "Converting…",
    "done": "Done",
    "error": "Failed",
    "skipped": "Already exists",
}

STATUS_COLOR = {
    "queued": MUTED,
    "working": ACCENT,
    "done": OK,
    "error": ERR,
    "skipped": WARN,
}


def human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def pick_family() -> str:
    available = set(tkfont.families())
    candidates = {
        "Windows": ("Segoe UI", "Calibri", "Arial"),
        "Darwin": ("SF Pro Text", "Helvetica Neue", "Helvetica"),
    }.get(platform.system(), ("Inter", "DejaVu Sans", "Noto Sans", "Arial"))
    for name in candidates:
        if name in available:
            return name
    return "TkDefaultFont"


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.items: dict[str, Item] = {}
        self.counter = 0
        self.out_dir: Path | None = None
        self.skip_existing = tk.BooleanVar(value=True)
        self.autostart = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Drop presentations to get started.")
        self.outdir_var = tk.StringVar(value="Next to each file")
        self.events: queue.Queue[tuple] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel = threading.Event()
        self.running = False
        self._drop_hot = False

        self._setup_window()
        self._setup_fonts()
        self._setup_style()
        self._build_ui()
        self._bind_keys()
        self._poll()

    # ------------------------------------------------------------------ setup

    def _setup_window(self) -> None:
        self.root.title(APP_NAME)
        self.root.configure(bg=BG)
        width, height = 900, 660
        self.root.minsize(760, 540)
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 3)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_fonts(self) -> None:
        family = pick_family()
        mono = "Consolas" if "Consolas" in tkfont.families() else "Courier New"
        self.font_base = (family, 10)
        self.font_small = (family, 9)
        self.font_bold = (family, 10, "bold")
        self.font_title = (family, 17, "bold")
        self.font_dz = (family, 12, "bold")
        self.font_icon = (family, 22)
        self.font_mono = (mono, 9)

    def _setup_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(".", font=self.font_base, background=CARD, foreground=TEXT)
        style.configure("TFrame", background=CARD)
        style.configure("BG.TFrame", background=BG)
        style.configure("TLabel", background=CARD, foreground=TEXT)
        style.configure("Title.TLabel", font=self.font_title, background=CARD, foreground=TEXT)
        style.configure("Sub.TLabel", font=self.font_small, background=CARD, foreground=MUTED)
        style.configure("Muted.TLabel", font=self.font_small, background=BG, foreground=MUTED)
        style.configure("Status.TLabel", font=self.font_small, background=BG, foreground=MUTED)
        style.configure("Danger.TLabel", font=self.font_small, background=BG, foreground=ERR)
        style.configure("Section.TLabel", font=self.font_bold, background=CARD, foreground=TEXT)

        style.configure(
            "TButton",
            font=self.font_base,
            padding=(12, 7),
            background="#ffffff",
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor="#ffffff",
            darkcolor="#ffffff",
            focuscolor=CARD,
            relief="flat",
        )
        style.map(
            "TButton",
            background=[("pressed", "#eef0f3"), ("active", "#f4f6f8"), ("disabled", "#f7f8fa")],
            foreground=[("disabled", "#98a2b3")],
            bordercolor=[("active", "#cfd4dc")],
        )

        style.configure(
            "Accent.TButton",
            font=self.font_bold,
            padding=(16, 9),
            background=ACCENT,
            foreground="#ffffff",
            bordercolor=ACCENT,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            focuscolor=ACCENT,
        )
        style.map(
            "Accent.TButton",
            background=[("pressed", ACCENT_DARK), ("active", ACCENT_DARK), ("disabled", "#b6c8f5")],
            foreground=[("disabled", "#f2f6ff")],
            bordercolor=[("disabled", "#b6c8f5")],
        )

        style.configure(
            "Stop.TButton",
            font=self.font_bold,
            padding=(16, 9),
            background="#ffffff",
            foreground=ERR,
            bordercolor="#f3b3ae",
            focuscolor="#ffffff",
        )
        style.map(
            "Stop.TButton",
            background=[("active", "#fef3f2"), ("pressed", "#fee4e2")],
        )

        style.configure(
            "TCheckbutton",
            background=CARD,
            foreground=TEXT,
            font=self.font_small,
            focuscolor=CARD,
        )
        style.map("TCheckbutton", background=[("active", CARD)])

        style.configure("TSeparator", background=BORDER)

        style.configure(
            "Treeview",
            font=self.font_base,
            rowheight=30,
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
        )
        style.map(
            "Treeview",
            background=[("selected", SELECT)],
            foreground=[("selected", TEXT)],
        )
        style.configure(
            "Treeview.Heading",
            font=self.font_small,
            background=CARD,
            foreground=MUTED,
            relief="flat",
            padding=(6, 6),
        )
        style.map("Treeview.Heading", background=[("active", CARD)])

        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#eaecf0",
            background=ACCENT,
            bordercolor="#eaecf0",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=6,
        )

    # --------------------------------------------------------------------- ui

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="BG.TFrame", padding=(18, 16, 18, 12))
        outer.pack(fill="both", expand=True)

        card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        inner = tk.Frame(card, bg=CARD, padx=20, pady=16)
        inner.pack(fill="both", expand=True)

        # Header -------------------------------------------------------
        header = tk.Frame(inner, bg=CARD)
        header.pack(fill="x")
        titles = tk.Frame(header, bg=CARD)
        titles.pack(side="left", anchor="w")
        ttk.Label(titles, text="Deck to PDF", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            titles,
            text="Convert .pptx, .ppt and .odp files without opening PowerPoint.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        backend = describe_backend("auto")
        chip = tk.Label(
            header,
            text=f"⚙  {backend}",
            font=self.font_small,
            bg=ACCENT_SOFT,
            fg=ACCENT,
            padx=10,
            pady=5,
        )
        chip.pack(side="right", anchor="ne")
        self.backend_label = chip

        ttk.Separator(inner).pack(fill="x", pady=(14, 14))

        # Drop zone ----------------------------------------------------
        self.drop_canvas = tk.Canvas(
            inner,
            height=172,
            bg="#fcfcfd",
            highlightthickness=0,
            bd=0,
        )
        self.drop_canvas.pack(fill="x")
        self.drop_canvas.bind("<Configure>", self._paint_dropzone)

        self._dz_border = "#d0d5dd"
        self.dz_icon = self.drop_canvas.create_text(
            0, 0, text="⇣", font=self.font_icon, fill=ACCENT, tags="dz"
        )
        self.dz_main = self.drop_canvas.create_text(
            0, 0, text="Drop presentations here", font=self.font_dz, fill=TEXT, tags="dz"
        )
        self.dz_sub = self.drop_canvas.create_text(
            0, 0,
            text=".ppt  .pptx  .pptm  .ppsx  .odp  —  or use the buttons below",
            font=self.font_small, fill=MUTED, tags="dz",
        )

        self.add_btn = ttk.Button(
            self.drop_canvas, text="Add files…", command=self.add_files, cursor="hand2"
        )
        self.add_folder_btn = ttk.Button(
            self.drop_canvas, text="Add folder…", command=self.add_folder, cursor="hand2"
        )
        self.dz_add_win = self.drop_canvas.create_window(0, 0, window=self.add_btn, tags="dz")
        self.dz_folder_win = self.drop_canvas.create_window(
            0, 0, window=self.add_folder_btn, tags="dz"
        )

        # File list ----------------------------------------------------
        list_header = tk.Frame(inner, bg=CARD)
        list_header.pack(fill="x", pady=(16, 6))
        self.count_label = ttk.Label(list_header, text="Files", style="Section.TLabel")
        self.count_label.pack(side="left")
        self.remove_btn = ttk.Button(
            list_header, text="Remove", command=self.remove_selected, state="disabled"
        )
        self.remove_btn.pack(side="right")
        self.clear_btn = ttk.Button(list_header, text="Clear", command=self.clear_all)
        self.clear_btn.pack(side="right", padx=(0, 8))

        table_wrap = tk.Frame(inner, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        table_wrap.pack(fill="both", expand=True)
        columns = ("name", "status", "detail")
        self.tree = ttk.Treeview(
            table_wrap,
            columns=columns,
            show="headings",
            selectmode="extended",
            height=7,
        )
        self.tree.heading("name", text="File", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.heading("detail", text="Details", anchor="w")
        self.tree.column("name", width=290, minwidth=160, stretch=False, anchor="w")
        self.tree.column("status", width=120, minwidth=100, stretch=False, anchor="w")
        self.tree.column("detail", width=380, minwidth=140, stretch=True, anchor="w")
        self.tree.tag_configure("odd", background=ROW_ALT)
        self.tree.tag_configure("done", foreground=OK)
        self.tree.tag_configure("error", foreground=ERR)
        self.tree.tag_configure("working", foreground=ACCENT)
        self.tree.tag_configure("skipped", foreground=WARN)

        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", self._on_double_click)
        self.tree.bind("<Delete>", lambda _e: self.remove_selected())
        self.tree.bind("<Button-3>", self._on_right_click)

        self.empty_hint = tk.Label(
            inner,
            text="No files yet — drop a deck above.",
            font=self.font_small,
            bg=CARD,
            fg=MUTED,
        )
        self.empty_hint.pack(pady=(6, 0))

        # Destination --------------------------------------------------
        dest = tk.Frame(inner, bg=CARD)
        dest.pack(fill="x", pady=(14, 0))
        ttk.Label(dest, text="Save to", style="Sub.TLabel").pack(side="left")
        self.outdir_label = tk.Label(
            dest,
            textvariable=self.outdir_var,
            font=self.font_small,
            bg=BG,
            fg=TEXT,
            padx=8,
            pady=4,
            anchor="w",
        )
        self.outdir_label.pack(side="left", padx=(10, 8), fill="x", expand=True)
        ttk.Button(dest, text="Change…", command=self.choose_outdir).pack(side="right")
        self.open_dir_btn = ttk.Button(
            dest, text="Open folder", command=self.open_output_folder
        )
        self.open_dir_btn.pack(side="right", padx=(0, 8))

        # Footer -------------------------------------------------------
        footer = tk.Frame(inner, bg=CARD)
        footer.pack(fill="x", pady=(14, 0))

        self.progress = ttk.Progressbar(
            footer, style="Accent.Horizontal.TProgressbar", mode="determinate"
        )
        self.progress.pack(fill="x", pady=(0, 10))

        actions = tk.Frame(footer, bg=CARD)
        actions.pack(fill="x")
        self.convert_btn = ttk.Button(
            actions,
            text="Convert all",
            style="Accent.TButton",
            command=self.start_or_stop,
            cursor="hand2",
        )
        self.convert_btn.pack(side="right")
        self.autostart_check = ttk.Checkbutton(
            actions,
            text="Start automatically when files are added",
            variable=self.autostart,
        )
        self.autostart_check.pack(side="right", padx=(0, 14))

        left = tk.Frame(footer, bg=CARD)
        left.pack(side="left", anchor="w")
        self.status_label = tk.Label(
            left, textvariable=self.status_var, font=self.font_small, bg=CARD, fg=MUTED, anchor="w"
        )
        self.status_label.pack(anchor="w")

        self.skip_check = ttk.Checkbutton(
            footer,
            text="Skip files that already have a PDF",
            variable=self.skip_existing,
        )
        self.skip_check.pack(anchor="w", pady=(10, 0))

        self._register_dnd()
        if DND_FILES is None:
            self.dz_sub.itemconfigure(
                text="Drag & drop unavailable (pip install tkinterdnd2) — use the buttons below"
            )

    def _register_dnd(self) -> None:
        if DND_FILES is None:
            return
        for widget in (self.root, self.drop_canvas, self.tree):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
                widget.dnd_bind("<<DropEnter>>", self._on_drop_enter)
                widget.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            except Exception:
                continue

    def _bind_keys(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.add_files())
        self.root.bind("<Control-O>", lambda _e: self.add_files())
        self.root.bind("<Control-Return>", lambda _e: self.start_or_stop())

    # --------------------------------------------------------------- dropzone

    def _paint_dropzone(self, _event: object | None = None) -> None:
        c = self.drop_canvas
        w = max(c.winfo_width(), 10)
        h = max(c.winfo_height(), 10)
        c.delete("chrome")
        c.create_rectangle(
            2,
            2,
            w - 2,
            h - 2,
            dash=(7, 5),
            outline=self._dz_border,
            width=2,
            tags="chrome",
        )
        c.tag_lower("chrome")
        c.coords(self.dz_icon, w / 2, h * 0.30)
        c.coords(self.dz_main, w / 2, h * 0.53)
        c.coords(self.dz_sub, w / 2, h * 0.70)
        c.coords(self.dz_add_win, w / 2 - 68, h * 0.87)
        c.coords(self.dz_folder_win, w / 2 + 68, h * 0.87)

    def _set_drop_hot(self, hot: bool) -> None:
        if hot == self._drop_hot:
            return
        self._drop_hot = hot
        self._dz_border = ACCENT if hot else "#d0d5dd"
        self.drop_canvas.configure(bg=ACCENT_SOFT if hot else "#fcfcfd")
        self.drop_canvas.itemconfigure(
            self.dz_main,
            text="Release to add" if hot else "Drop presentations here",
        )
        self._paint_dropzone()

    def _on_drop_enter(self, _event: object | None = None) -> None:
        self._set_drop_hot(True)

    def _on_drop_leave(self, _event: object | None = None) -> None:
        self._set_drop_hot(False)

    def _on_drop(self, event) -> None:  # noqa: ANN001 - tkinter event
        self._set_drop_hot(False)
        try:
            raw = self.root.tk.splitlist(event.data)
        except Exception:
            raw = [event.data]
        self._add_paths([Path(p) for p in raw], autostart=True)

    # ------------------------------------------------------------------- files

    def add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Choose presentations",
            filetypes=[
                ("Presentations", " ".join(f"*{s}" for s in sorted(SUPPORTED_SUFFIXES))),
                ("PowerPoint", "*.ppt *.pptx *.pptm"),
                ("All files", "*.*"),
            ],
        )
        if files:
            self._add_paths([Path(f) for f in files])

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose a folder of presentations")
        if folder:
            self._add_paths([Path(folder)])

    def _add_paths(
        self, paths: list[str | os.PathLike[str]], autostart: bool = False
    ) -> None:
        added = 0
        skipped = 0
        for raw in paths:
            path = Path(raw).expanduser()
            candidates = (
                [p for p in sorted(path.iterdir()) if p.is_file()]
                if path.is_dir()
                else [path]
            )
            for candidate in candidates:
                if candidate.suffix.lower() not in SUPPORTED_SUFFIXES:
                    continue
                resolved = candidate.resolve()
                if any(item.path == resolved for item in self.items.values()):
                    skipped += 1
                    continue
                self._insert_item(resolved)
                added += 1

        if added:
            msg = f"Added {added} file{'s' if added != 1 else ''}."
            if skipped:
                msg += f" {skipped} already in the list."
            self._set_status(msg)
            self._layout_rows()
            if autostart and self.autostart.get() and not self.running:
                self.start()
        elif skipped:
            self._set_status("Those files are already in the list.")
        else:
            self._set_status("No supported presentations found there.")

    def _insert_item(self, path: Path) -> Item:
        self.counter += 1
        iid = f"item{self.counter}"
        item = Item(path=path, iid=iid)
        try:
            item.detail = f"{human_size(path.stat().st_size)}"
        except OSError:
            item.detail = ""
        self.items[iid] = item
        self.tree.insert(
            "",
            "end",
            iid=iid,
            values=(path.name, STATUS_TEXT["queued"], item.detail),
            tags=("odd",) if self.counter % 2 == 0 else (),
        )
        return item

    def _layout_rows(self) -> None:
        count = len(self.items)
        self.count_label.configure(text=f"Files ({count})" if count else "Files")
        if count:
            self.empty_hint.pack_forget()
        else:
            self.empty_hint.pack(pady=(6, 0))
        self.clear_btn.state(["!disabled"] if count else ["disabled"])
        self.convert_btn.state(["!disabled"] if count else ["disabled"])

    def _refresh_row(self, item: Item) -> None:
        tags: tuple[str, ...] = ()
        if item.status in ("done", "error", "working", "skipped"):
            tags = (item.status,)
        elif self.counter % 2 == 0:
            tags = ("odd",)
        self.tree.item(
            item.iid,
            values=(item.path.name, STATUS_TEXT[item.status], item.detail),
            tags=tags,
        )

    def remove_selected(self) -> None:
        if self.running:
            self._set_status("Finish or stop the current conversion first.")
            return
        for iid in self.tree.selection():
            item = self.items.pop(iid, None)
            if item is not None:
                self.tree.delete(iid)
                self._set_status(f"Removed {item.path.name}.")
        self._layout_rows()

    def clear_all(self) -> None:
        if self.running:
            self._set_status("Finish or stop the current conversion first.")
            return
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self.progress["value"] = 0
        self._set_status("Cleared. Drop presentations to get started.")
        self._layout_rows()

    def _on_select(self, _event: object | None = None) -> None:
        self.remove_btn.state(["!disabled"] if self.tree.selection() else ["disabled"])

    def _on_double_click(self, _event: object | None = None) -> None:
        for iid in self.tree.selection():
            item = self.items.get(iid)
            if item is None:
                continue
            if item.status == "done" and item.output and item.output.is_file():
                self._open_path(item.output)
                return
            if item.status == "error":
                messagebox.showerror("Conversion failed", item.error or "Unknown error")
                return
            if item.path.is_file():
                self._open_path(item.path)
                return

    def _on_right_click(self, event) -> None:  # noqa: ANN001 - tkinter event
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        item = self.items[iid]

        menu = tk.Menu(self.root, tearoff=0, font=self.font_small)
        if item.status == "done" and item.output:
            menu.add_command(label="Open PDF", command=lambda: self._open_path(item.output))
        menu.add_command(
            label="Show in folder", command=lambda: self._reveal(item.output or item.path)
        )
        if item.error:
            menu.add_separator()
            menu.add_command(
                label="Copy error message", command=lambda: self._copy(item.error)
            )
            menu.add_command(
                label="Show error…",
                command=lambda: messagebox.showerror("Conversion failed", item.error),
            )
        menu.add_separator()
        menu.add_command(label="Remove from list", command=self.remove_selected)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------ output

    def choose_outdir(self) -> None:
        folder = filedialog.askdirectory(
            title="Where should the PDFs go?",
            initialdir=str(self.out_dir) if self.out_dir else None,
        )
        if not folder:
            return
        self.out_dir = Path(folder).resolve()
        text = str(self.out_dir)
        if len(text) > 72:
            text = "…" + text[-70:]
        self.outdir_var.set(text)
        self.outdir_label.configure(bg=ACCENT_SOFT, fg=TEXT)
        self._set_status(f"PDFs will be saved to {self.out_dir}")

    def open_output_folder(self) -> None:
        target = self.out_dir
        if target is None:
            done = [i.output for i in self.items.values() if i.output]
            if done:
                target = done[-1].parent
            elif self.items:
                target = next(iter(self.items.values())).path.parent
        if target and target.is_dir():
            self._reveal(target)
        else:
            self._set_status("The output folder does not exist yet.")

    # -------------------------------------------------------------- conversion

    def start_or_stop(self) -> None:
        if self.running:
            self.cancel.set()
            self._set_status("Stopping after the current file…")
            self.convert_btn.configure(text="Stopping…", state="disabled")
            return
        self.start()

    def start(self) -> None:
        pending = [
            item
            for item in self.items.values()
            if item.status in ("queued", "error", "skipped")
        ]
        if not pending:
            if self.items:
                self._set_status("Everything is already converted.")
            return
        if self.out_dir is None:
            self.out_dir = None  # explicit: PDFs go next to each source

        self.cancel.clear()
        self.running = True
        self.progress.configure(maximum=len(pending), value=0)
        self.convert_btn.configure(text="Stop", style="Stop.TButton", state="!disabled")
        # Read the Tk variables here (main thread); the worker must not touch Tk.
        self.worker = threading.Thread(
            target=self._work,
            args=(pending, self.skip_existing.get()),
            daemon=True,
        )
        self.worker.start()
        self._set_status(f"Converting {len(pending)} file{'s' if len(pending) != 1 else ''}…")

    def _destination_for(self, item: Item) -> Path | None:
        if self.out_dir is None:
            return None
        return self.out_dir / f"{item.path.stem}.pdf"

    def _work(self, pending: list[Item], skip_existing: bool) -> None:
        completed = 0
        failed = 0
        skipped = 0
        for item in pending:
            if self.cancel.is_set():
                break
            destination = self._destination_for(item)
            if skip_existing and destination and destination.is_file():
                self.events.put(("skipped", item.iid))
                skipped += 1
                completed += 1
                continue

            self.events.put(("start", item.iid))
            started = time.monotonic()
            try:
                output = convert(
                    item.path,
                    destination,
                    method="auto",
                    timeout=300.0,
                )
                self.events.put(
                    ("done", item.iid, output, time.monotonic() - started)
                )
                completed += 1
            except Exception as e:  # noqa: BLE001 - surfaced in the UI
                self.events.put(
                    ("error", item.iid, str(e), time.monotonic() - started)
                )
                failed += 1
                completed += 1

        self.events.put(("finished", completed, failed, skipped, self.cancel.is_set()))

    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "start":
            item = self.items.get(event[1])
            if item:
                item.status = "working"
                item.detail = "Converting…"
                self._refresh_row(item)
                self.tree.see(item.iid)
                self._set_status(f"Converting {item.path.name}…")
        elif kind == "done":
            _, iid, output, seconds = event
            item = self.items.get(iid)
            if item:
                item.status = "done"
                item.output = output
                item.seconds = seconds
                try:
                    size = human_size(output.stat().st_size)
                except OSError:
                    size = "PDF"
                item.detail = f"{output.name} · {size} · {seconds:.1f}s"
                self._refresh_row(item)
            self._bump_progress()
            self._set_status(f"Saved {Path(output).name}")
        elif kind == "error":
            _, iid, message, seconds = event
            item = self.items.get(iid)
            if item:
                item.status = "error"
                item.error = message
                item.detail = message.splitlines()[0][:110]
                self._refresh_row(item)
            self._bump_progress()
            self._set_status(f"Failed: {item.path.name if item else ''}")
        elif kind == "skipped":
            item = self.items.get(event[1])
            if item:
                item.status = "skipped"
                item.output = self._destination_for(item)
                item.detail = "PDF already exists — not overwritten"
                self._refresh_row(item)
            self._bump_progress()
        elif kind == "finished":
            _, completed, failed, skipped, cancelled = event
            self._finish(completed, failed, skipped, cancelled)

    def _bump_progress(self) -> None:
        current = self.progress["value"]
        self.progress["value"] = min(current + 1, self.progress["maximum"])

    def _finish(self, completed: int, failed: int, skipped: int, cancelled: bool) -> None:
        self.running = False
        self.convert_btn.configure(text="Convert all", style="Accent.TButton", state="!disabled")
        parts = [f"{completed - failed} converted"]
        if skipped:
            parts.append(f"{skipped} skipped")
        if failed:
            parts.append(f"{failed} failed")
        summary = "  ·  ".join(parts)
        if cancelled:
            summary = f"Stopped — {summary}"
        self._set_status(summary)
        self.status_label.configure(fg=ERR if failed else MUTED)
        if failed and not cancelled:
            messagebox.showwarning(
                "Some files could not be converted",
                f"{failed} of {completed} file(s) failed.\n\n"
                "Double-click a failed row to see the details.",
            )

    # ------------------------------------------------------------------ helper

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.status_label.configure(fg=ERR if text.lower().startswith(("fail", "error")) else MUTED)

    def _copy(self, text: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status("Error message copied to clipboard.")

    def _open_path(self, path: Path) -> None:
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:  # noqa: BLE001
            self._set_status(f"Could not open: {e}")

    def _reveal(self, path: Path) -> None:
        try:
            if platform.system() == "Windows":
                if path.is_dir():
                    os.startfile(str(path))  # type: ignore[attr-defined]
                else:
                    subprocess.Popen(["explorer", "/select,", str(path)])
            elif platform.system() == "Darwin":
                subprocess.Popen(
                    ["open", "-R", str(path)] if path.is_file() else ["open", str(path)]
                )
            else:
                subprocess.Popen(["xdg-open", str(path if path.is_dir() else path.parent)])
        except Exception as e:  # noqa: BLE001
            self._set_status(f"Could not open folder: {e}")


def create_root() -> tk.Tk:
    if TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def main() -> int:
    if not find_libreoffice() and platform.system() != "Windows":
        print(
            "Warning: no conversion backend found. Install LibreOffice.",
            file=sys.stderr,
        )

    root = create_root()
    ConverterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())