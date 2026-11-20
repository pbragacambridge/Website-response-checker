"""
URL Checker - desktop app.

Type a URL, or drop an Excel file of URLs (column A) onto the window, and watch each site being
checked live in a real browser. Results are saved as "<file name> - results.xlsx" next to the file.
"""
import asyncio
import ctypes
import os
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import checker

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # drag and drop is a nice-to-have - clicking to choose a file still works
    TkinterDnD = None
try:
    import sv_ttk
except Exception:
    sv_ttk = None

APP = "URL Checker"
HERE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
SPINNER = "◐◓◑◒"
ICONS = {"ok": "✓", "warn": "⚠", "fail": "✗", "bad": "✗", "queued": "•", "stopped": "–"}
PALETTE = {
    "light": {"bg": "#fafafa", "panel": "#ffffff", "line": "#e0e0e0", "border": "#b4b4b4", "text": "#1b1b1b",
              "muted": "#616161", "accent": "#005fb8", "ok": "#0f7b0f", "warn": "#9d5d00", "bad": "#c42b1c"},
    "dark": {"bg": "#1c1c1c", "panel": "#232323", "line": "#3a3a3a", "border": "#6a6a6a", "text": "#f3f3f3",
             "muted": "#a8a8a8", "accent": "#60cdff", "ok": "#6ccb5f", "warn": "#fce100", "bad": "#ff99a4"},
}


def windows_dark_mode():
    if os.environ.get("URLCHECKER_THEME") in ("light", "dark"):
        return os.environ["URLCHECKER_THEME"] == "dark"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def clock(seconds):
    m, s = divmod(int(seconds), 60)
    return f"{m}:{s:02d}"


class Engine:
    """Runs the browser and the checks on a background thread, so the window never freezes.
    Everything that happens is put on `events` for the window to show."""

    def __init__(self, events):
        self.events = events
        self.checker = checker.Checker()
        self.loop = asyncio.new_event_loop()
        self.ready = asyncio.Event()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self._start())
        self.loop.run_forever()

    async def _start(self):
        try:
            await self.checker.start()
            self.events.put(("browser", self.checker.name))
        except Exception as e:
            self.events.put(("browser_error", f"{type(e).__name__}: {e}"))
        self.ready.set()

    def submit(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    async def check(self, key, host):
        await self.ready.wait()
        self.events.put(("start", key))
        steps = checker.Steps(lambda items: self.events.put(("steps", key, items)))
        try:
            result = await self.checker.check(host, steps)
        except Exception as e:  # never let one URL stop the run
            result = f"Failed - {e}" if not self.checker.browser else f"Failed - {type(e).__name__}"
        self.events.put(("done", key, result))

    async def check_list(self, keys_and_hosts):
        limit = asyncio.Semaphore(checker.PARALLEL)

        async def one(key, host):
            async with limit:
                await self.check(key, host)

        await asyncio.gather(*(one(k, h) for k, h in keys_and_hosts))

    def close(self):
        try:
            self.submit(self.checker.stop()).result(timeout=8)
        except Exception:
            pass
        self.loop.call_soon_threadsafe(self.loop.stop)


@dataclass
class Row:
    host: str
    number: str
    state: str = "queued"   # queued, running, ok, warn, bad or stopped
    result: str = ""
    steps: list = field(default_factory=list)
    started: float = 0.0


class App:
    def __init__(self, root, dnd, first_file=None):
        self.root = root
        self.events = queue.Queue()
        self.engine = Engine(self.events)
        self.rows = {}           # table row id -> Row
        self.running = set()     # table row ids being checked right now
        self.singles = 0         # one-off URL checks so far
        self.batch = 0           # Excel runs so far (stale events from older runs are ignored)
        self.batch_future = None
        self.book = None
        self.results = {}
        self.total = self.done = 0
        self.started = self.finished = None
        self.saved_path = None
        self.unsaved = False     # a finished run whose results haven't been saved yet
        self.file_name = ""
        self.spin = 0
        self.placeholder = False

        theme = "dark" if windows_dark_mode() and sv_ttk else "light"
        self.c = dict(PALETTE[theme])
        if sv_ttk:
            sv_ttk.set_theme(theme)
        else:  # plain Windows look
            self.c["bg"] = ttk.Style().lookup(".", "background") or self.c["bg"]
        self.scale = root.winfo_fpixels("1i") / 96
        self._build(dnd)
        root.after(50, self._poll)
        root.after(150, self._tick)
        if first_file:  # the file was dropped onto the app's icon
            root.after(300, lambda: self.load_file(first_file))

    # ---------- layout ----------

    def _build(self, dnd):
        c, px = self.c, lambda n: int(n * self.scale)
        outer = ttk.Frame(self.root, padding=(px(20), px(14), px(20), px(14)))
        outer.pack(fill="both", expand=True)

        head = ttk.Frame(outer)
        head.pack(fill="x")
        ttk.Label(head, text=APP, font=("Segoe UI Semibold", 18)).pack(side="left")
        self.browser_label = ttk.Label(head, text="Starting browser…", foreground=c["muted"])
        self.browser_label.pack(side="right")
        ttk.Label(outer, text="Opens each site in a real browser and records what a visitor sees.",
                  foreground=c["muted"]).pack(anchor="w", pady=(0, px(12)))

        one = ttk.Frame(outer)
        one.pack(fill="x")
        self.url = ttk.Entry(one, font=("Segoe UI", 11))
        self.url.pack(side="left", fill="x", expand=True)
        self.url.bind("<Return>", lambda e: self.check_one())
        self.url.bind("<FocusIn>", self._hide_placeholder)
        self.url.bind("<FocusOut>", self._show_placeholder)
        self._show_placeholder()
        ttk.Button(one, text="Check URL", command=self.check_one,
                   style="Accent.TButton" if sv_ttk else "TButton").pack(side="left", padx=(px(8), 0))

        self.drop = tk.Canvas(outer, height=px(78), highlightthickness=0, bd=0, bg=c["bg"], cursor="hand2")
        self.drop.pack(fill="x", pady=(px(12), 0))
        self.drop.bind("<Configure>", lambda e: self._draw_drop())
        self.drop.bind("<Button-1>", lambda e: self.browse())

        prog = ttk.Frame(outer)
        prog.pack(fill="x", pady=(px(12), 0))
        self.bar = ttk.Progressbar(prog, mode="determinate", maximum=1, value=0)
        self.bar.pack(side="left", fill="x", expand=True)
        self.count = ttk.Label(prog, text="No file yet", foreground=c["muted"])
        self.count.pack(side="left", padx=(px(12), 0))
        self.stop_btn = ttk.Button(prog, text="Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(px(12), 0))

        tallies = ttk.Frame(outer)
        tallies.pack(fill="x", pady=(px(4), 0))
        self.tally = {}
        for kind in ("ok", "warn", "bad"):
            self.tally[kind] = ttk.Label(tallies, foreground=c[kind], font=("Segoe UI Semibold", 10))
            self.tally[kind].pack(side="left", padx=(0, px(18)))

        foot = ttk.Frame(outer)  # packed before the list so it always keeps its space at the bottom
        foot.pack(side="bottom", fill="x", pady=(px(12), 0))
        self.saved = ttk.Label(foot, text="", foreground=c["muted"])
        self.saved.pack(side="left")
        self.folder_btn = ttk.Button(foot, text="Show in folder", command=self.show_folder, state="disabled")
        self.folder_btn.pack(side="right")
        self.open_btn = ttk.Button(foot, text="Open results", command=self.open_results, state="disabled")
        self.open_btn.pack(side="right", padx=(0, px(8)))
        self.save_btn = ttk.Button(foot, text="Save results…", command=self.save_results, state="disabled",
                                   style="Accent.TButton" if sv_ttk else "TButton")
        self.save_btn.pack(side="right", padx=(0, px(8)))

        pane = ttk.PanedWindow(outer, orient="horizontal")
        pane.pack(fill="both", expand=True, pady=(px(8), 0))
        left, right = ttk.Frame(pane), ttk.Frame(pane)
        pane.add(left, weight=5)
        pane.add(right, weight=3)
        self.root.after(200, lambda: pane.sashpos(0, int(pane.winfo_width() * 0.62)))

        self.tree = ttk.Treeview(left, columns=("n", "url", "result"), show="headings", selectmode="browse",
                                 height=1)
        for col, text, width, stretch in (("n", "#", 40, False), ("url", "URL", 210, True),
                                          ("result", "Result", 470, True)):
            self.tree.heading(col, text=text, anchor="w")
            self.tree.column(col, width=px(width), stretch=stretch, anchor="w")
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        for tag in ("ok", "warn", "bad"):
            self.tree.tag_configure(tag, foreground=c[tag])
        self.tree.tag_configure("running", foreground=c["accent"])
        self.tree.tag_configure("queued", foreground=c["muted"])
        self.tree.tag_configure("stopped", foreground=c["muted"])
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._show_steps())

        ttk.Label(right, text="Checks", font=("Segoe UI Semibold", 11)).pack(anchor="w", padx=(px(14), 0))
        self.steps = tk.Text(right, wrap="word", relief="flat", bd=0, padx=px(14), pady=px(10), cursor="arrow",
                             bg=c["panel"], fg=c["text"], font=("Segoe UI", 10), height=1, width=1,
                             highlightthickness=1, highlightbackground=c["line"], highlightcolor=c["line"])
        self.steps.pack(fill="both", expand=True, padx=(px(12), 0), pady=(px(4), 0))
        t, symbol = self.steps, ("Segoe UI Symbol", 11)
        t.tag_configure("host", font=("Segoe UI Semibold", 12), spacing3=px(10))
        for kind in ("ok", "warn", "bad"):
            t.tag_configure(kind, foreground=c[kind], font=symbol)
            t.tag_configure("result-" + kind, foreground=c[kind], font=("Segoe UI Semibold", 11))
        t.tag_configure("fail", foreground=c["bad"], font=symbol)
        t.tag_configure("running", foreground=c["accent"], font=symbol)
        t.tag_configure("label", lmargin2=px(26))
        t.tag_configure("detail", foreground=c["muted"], font=("Segoe UI", 9),
                        lmargin1=px(26), lmargin2=px(26), spacing3=px(6))
        t.tag_configure("heading", font=("Segoe UI Semibold", 10), spacing1=px(12), spacing3=px(4))
        t.tag_configure("hint", foreground=c["muted"])
        self._show_steps()

        if dnd:  # dropping a file anywhere on the window works
            for widget in (self.root, self.drop, self.tree, self.steps):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._drag_enter)
                widget.dnd_bind("<<DropLeave>>", self._drag_leave)
                widget.dnd_bind("<<Drop>>", self._drop)

    def _draw_drop(self, hover=False):
        d, c = self.drop, self.c
        d.delete("all")
        w, h = d.winfo_width(), d.winfo_height()
        d.create_rectangle(1, 1, w - 2, h - 2, outline=c["accent"] if hover else c["border"], dash=(5, 3))
        busy = self.batch_future and not self.batch_future.done()
        title = (f"Checking  {self.file_name}" if busy else "Drop it to start checking" if hover
                 else "Drop an Excel file of URLs here")
        sub = ("Drop another file when this one is finished" if busy else
               "or click to choose one  ·  URLs in column A  ·  .xlsx, .csv or .txt")
        d.create_text(w / 2, h / 2 - 11 * self.scale, text=title, fill=c["accent"] if hover else c["text"],
                      font=("Segoe UI Semibold", 12))
        d.create_text(w / 2, h / 2 + 13 * self.scale, text=sub, fill=c["muted"], font=("Segoe UI", 9))

    def _show_placeholder(self, _=None):
        if not self.url.get():
            self.url.insert(0, "Check one URL, e.g. example.com - press Enter")
            self.url.configure(foreground=self.c["muted"])
            self.placeholder = True

    def _hide_placeholder(self, _=None):
        if self.placeholder:
            self.url.delete(0, "end")
            self.url.configure(foreground=self.c["text"])
            self.placeholder = False

    # ---------- actions ----------

    def check_one(self):
        host = "" if self.placeholder else self.url.get().strip().strip("'\"")
        if "." not in host:
            messagebox.showinfo(APP, "Type a website address first, e.g. example.com", parent=self.root)
            return
        self.singles += 1
        key = f"s{self.singles}"
        self.rows[key] = Row(host, "")
        self.tree.insert("", 0, iid=key, values=("", host, ""), tags=("queued",))
        self._paint(key)
        self.tree.selection_set(key)
        self.tree.see(key)
        self.url.select_range(0, "end")  # typing replaces it, ready for the next URL
        self.engine.submit(self.engine.check(key, host))

    def browse(self):
        if self.batch_future and not self.batch_future.done():
            return
        path = filedialog.askopenfilename(parent=self.root, title="Choose a list of URLs", filetypes=[
            ("Excel, CSV or text files", "*.xlsx *.xlsm *.csv *.txt"), ("All files", "*.*")])
        if path:
            self.load_file(path)

    def load_file(self, path):
        path = Path(str(path).strip().strip('"'))
        if self.batch_future and not self.batch_future.done():
            messagebox.showinfo(APP, "Already checking a file. Wait for it to finish, or press Stop.",
                                parent=self.root)
            return
        if not self._results_safe():
            return
        try:
            book = checker.read_list(path)
        except PermissionError:
            messagebox.showerror(APP, f"Couldn't read {path.name} - another program has it locked.\n\n"
                                      "Close it (e.g. in Excel) and drop it again.", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror(APP, f"Couldn't read {path.name}\n\n{e}", parent=self.root)
            return
        hosts = list(dict.fromkeys(book.hosts))
        if not hosts:
            messagebox.showinfo(APP, f"No URLs found in column A of {path.name}.", parent=self.root)
            return

        self.batch += 1
        self.book, self.results, self.saved_path, self.file_name = book, {}, None, path.name
        self.unsaved = False
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.running.clear()
        jobs = []
        for n, host in enumerate(hosts, 1):
            key = f"b{self.batch}:{n}"
            self.rows[key] = Row(host, str(n))
            self.tree.insert("", "end", iid=key, values=(n, host, ""), tags=("queued",))
            self._paint(key)
            jobs.append((key, host))
        self.total, self.done = len(hosts), 0
        self.started, self.finished = time.monotonic(), None
        self.bar.configure(maximum=self.total, value=0)
        self.stop_btn.configure(state="normal")
        for button in (self.save_btn, self.open_btn, self.folder_btn):
            button.configure(state="disabled")
        self.saved.configure(text="When it finishes you'll be asked where to save the results",
                             foreground=self.c["muted"])
        self._update_tallies()
        self._show_steps()

        batch = self.batch
        self.batch_future = self.engine.submit(self.engine.check_list(jobs))
        self.batch_future.add_done_callback(
            lambda f: self.events.put(("batch_end", batch, f.cancelled() or bool(f.exception()))))
        self._draw_drop()

    def stop(self):
        if self.batch_future and not self.batch_future.done():
            self.batch_future.cancel()

    def save_results(self):
        """Ask where to save the results (suggesting "<name> - results.xlsx" next to the list) and save them.
        Returns True if they were saved."""
        if not self.book:
            return False
        while True:
            path = filedialog.asksaveasfilename(
                parent=self.root, title="Save the results as", initialdir=str(self.book.path.parent),
                initialfile=checker.results_name(self.book), defaultextension=".xlsx",
                filetypes=[("Excel workbook", "*.xlsx")])
            if not path:
                return False
            path = Path(path)
            if path.resolve() == self.book.path.resolve():
                messagebox.showinfo(APP, "That's your original list. Choose a different name so it isn't "
                                         "overwritten.", parent=self.root)
                continue
            try:
                self.saved_path = checker.save_results(self.book, self.results, out=path,
                                                       missing="Not checked (stopped)")
                break
            except PermissionError:
                messagebox.showerror(APP, f"Couldn't save {path.name} - it's probably open in Excel.\n\n"
                                          "Close it there, or choose a different name.", parent=self.root)
            except Exception as e:
                messagebox.showerror(APP, f"Couldn't save the results\n\n{e}", parent=self.root)
        self.unsaved = False
        self.saved.configure(text=f"Saved  {self.saved_path.name}", foreground=self.c["text"])
        self.open_btn.configure(state="normal")
        self.folder_btn.configure(state="normal")
        return True

    def _results_safe(self):
        """Before the results would be lost: True if they're saved, or the user chose not to save them."""
        if not self.unsaved:
            return True
        answer = messagebox.askyesnocancel(APP, "The results from the last file haven't been saved.\n\n"
                                                "Save them now?", parent=self.root)
        if answer is None:
            return False
        return self.save_results() if answer else True

    def open_results(self):
        if self.saved_path:
            os.startfile(self.saved_path)

    def show_folder(self):
        if self.saved_path:
            subprocess.Popen(["explorer", "/select,", str(self.saved_path)])

    def close(self):
        if self.batch_future and not self.batch_future.done():
            answer = messagebox.askyesnocancel(APP, "A file is still being checked.\n\n"
                                                    "Save what's been checked so far before quitting?",
                                               parent=self.root)
            if answer is None:
                return
            self.batch_future.cancel()
            self._finish_batch(stopped=True)
            self.batch += 1  # ignore the stopped run's last events
            if answer and not self.save_results():
                return
        elif not self._results_safe():
            return
        self.root.withdraw()
        self.engine.close()
        self.root.destroy()

    # ---------- drag and drop ----------

    def _drag_enter(self, event):
        self._draw_drop(hover=True)
        return event.action

    def _drag_leave(self, event):
        self._draw_drop()
        return event.action

    def _drop(self, event):
        self._draw_drop()
        files = self.root.tk.splitlist(event.data)
        if files:
            self.load_file(files[0])
        return event.action

    # ---------- live updates ----------

    def _poll(self):
        try:
            for _ in range(500):
                self._handle(*self.events.get_nowait())
        except queue.Empty:
            pass
        self.root.after(50, self._poll)

    def _handle(self, kind, *args):
        if kind == "browser":
            self.browser_label.configure(text=f"Checking with {args[0]}")
        elif kind == "browser_error":
            self.browser_label.configure(text="No browser available", foreground=self.c["bad"])
            messagebox.showerror(APP, "Couldn't start Google Chrome or Microsoft Edge, so sites can't be checked.\n\n"
                                      f"{args[0]}\n\nIf your computer blocks unknown programs, it may have blocked "
                                      "the part of this app that controls the browser (node.exe).", parent=self.root)
        elif kind == "batch_end":
            if args[0] == self.batch:
                self._finish_batch(stopped=args[1])
                self.root.after(200, self.save_results)  # ask where to save them
        else:
            key = args[0]
            row = self.rows.get(key)
            if not row or row.state == "stopped":
                return
            if kind == "start":
                row.state, row.started = "running", time.monotonic()
                self.running.add(key)
                self._paint(key)
                if not self.tree.selection():  # follow along until the user picks a row
                    self.tree.see(key)
            elif kind == "steps":
                row.steps = args[1]
            elif kind == "done":
                row.result, row.state = args[1], checker.verdict(args[1])
                self.running.discard(key)
                self._paint(key)
                if key.startswith("b"):
                    self.results[row.host] = row.result
                    self.done += 1
                    self.bar.configure(value=self.done)
                    self._update_tallies()
            if self._selected() == key:
                self._show_steps()

    def _tick(self):
        """Animates the spinners and the timers."""
        self.spin = (self.spin + 1) % len(SPINNER)
        for key in self.running:
            self._paint(key)
        if self._selected() in self.running:
            self._show_steps()
        if self.started:
            elapsed = clock((self.finished or time.monotonic()) - self.started)
            self.count.configure(text=f"{self.done} of {self.total} checked  ·  {elapsed}")
        self.root.after(150, self._tick)

    def _finish_batch(self, stopped):
        self.finished = time.monotonic()
        self.stop_btn.configure(state="disabled")
        for key in [k for k in self.rows if k.startswith(f"b{self.batch}:")]:
            if self.rows[key].state in ("queued", "running"):
                self.rows[key].state = "stopped"
                self.running.discard(key)
                self._paint(key)
        self.unsaved = True
        self.save_btn.configure(state="normal")
        self.saved.configure(text=f"{'Stopped' if stopped else 'Finished'}  ·  results not saved yet",
                             foreground=self.c["text"])
        self._draw_drop()
        self._show_steps()
        try:  # flash the taskbar button in case the window is behind others
            ctypes.windll.user32.FlashWindow(int(self.root.wm_frame(), 16), True)
        except Exception:
            pass

    def _update_tallies(self):
        counts = {k: sum(r.state == k for r in self.rows.values() if r.number) for k in ("ok", "warn", "bad")}
        for kind, word in (("ok", "working"), ("warn", "warnings"), ("bad", "errors")):
            self.tally[kind].configure(text=f"{ICONS[kind]} {counts[kind]} {word}" if self.total else "")

    def _selected(self):
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _paint(self, key):
        row = self.rows[key]
        if row.state == "running":
            current = next((s[0] for s in reversed(row.steps) if s[1] == "running"), "Starting")
            text = f"{SPINNER[self.spin]}  {current}…  {time.monotonic() - row.started:.0f}s"
        elif row.state == "queued":
            text = "•  waiting"
        elif row.state == "stopped":
            text = "–  not checked (stopped)"
        else:
            text = f"{ICONS[row.state]}  {row.result}"
        self.tree.item(key, values=(row.number, row.host, text), tags=(row.state,))

    def _show_steps(self):
        t = self.steps
        top = t.yview()[0]
        t.configure(state="normal")
        t.delete("1.0", "end")
        key = self._selected()
        row = self.rows.get(key) if key else None
        if not row:
            t.insert("end", "Type a URL above, or drop an Excel file.\n\nClick any URL in the list to watch "
                            "its checks tick off as they happen.", "hint")
        else:
            t.insert("end", row.host + "\n", "host")
            now = time.monotonic()
            for label, state, detail, started, ended in row.steps:
                icon = SPINNER[self.spin] if state == "running" else ICONS[state]
                secs = (ended or now) - started
                timing = f"{secs:.0f}s" if state == "running" else f"{secs:.1f}s" if secs >= 1 else ""
                t.insert("end", icon, state)
                t.insert("end", f"   {label}\n", "label")
                info = "  ·  ".join(x for x in (detail, timing) if x)
                if info:
                    t.insert("end", info + "\n", "detail")
            if row.state == "queued":
                t.insert("end", f"Waiting for a free slot - up to {checker.PARALLEL} sites are checked at once.",
                         "hint")
            elif row.state == "stopped":
                t.insert("end", "\nNot checked - the run was stopped.", "hint")
            elif row.result:
                t.insert("end", "Result\n", "heading")
                t.insert("end", f"{ICONS[row.state]}  {row.result}", "result-" + row.state)
        t.configure(state="disabled")
        t.yview_moveto(top)


def main():
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # sharp text on high-resolution screens
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("URLChecker.App")
        except Exception:
            pass
    dnd = False
    if TkinterDnD:
        try:
            root, dnd = TkinterDnD.Tk(), True
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()
    root.title(APP)
    if (HERE / "icon.ico").exists():
        root.iconbitmap(default=str(HERE / "icon.ico"))
    scale = root.winfo_fpixels("1i") / 96
    root.geometry(f"{int(1180 * scale)}x{int(780 * scale)}")
    root.minsize(int(860 * scale), int(600 * scale))
    app = App(root, dnd, sys.argv[1] if len(sys.argv) > 1 else None)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()


if __name__ == "__main__":
    main()
