import json
import re
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.font as tkfont
import sys
import os
import ctypes
import winsound
import webbrowser
from io import BytesIO
from typing import Optional, Dict, Tuple, List
from PIL import Image, ImageDraw
import tempfile
import shutil
import time
import urllib.request
import urllib.error

# ====================== App Constants ======================
APP_NAME_VERSION = "Windows App Updater v1.2.1"
APP_VERSION_ONLY = "1.2.1"

GITHUB_RELEASES_PAGE = "https://github.com/ilukezippo/Windows-App-Updater/releases"
GITHUB_API_LATEST = "https://api.github.com/repos/ilukezippo/Windows-App-Updater/releases/latest"

# Window sizes for log toggle
WIN_W = 950
WIN_H_FULL = 650
WIN_H_COMPACT = 500

# ====================== PyInstaller resource helper ======================
def resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)  # type: ignore[attr-defined]
    return os.path.join(os.path.abspath("."), relative_path)

# ====================== Hide child console windows ======================
CREATE_NO_WINDOW = 0x08000000

def _hidden_startupinfo() -> subprocess.STARTUPINFO:  # type: ignore[name-defined]
    si = subprocess.STARTUPINFO()  # type: ignore[attr-defined]
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW      # type: ignore[attr-defined]
    si.wShowWindow = 0  # SW_HIDE
    return si

# ====================== Elevation helpers ======================
def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def relaunch_as_admin():
    if is_admin():
        return
    if getattr(sys, "frozen", False):
        app = sys.executable
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    else:
        app = sys.executable
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{script}"'] + [f'"{a}"' for a in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    sys.exit(0)

# ====================== Tooltip helper ======================
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, _=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=6, ipady=2)

    def hide_tip(self, _=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

# ====================== Icon & sound helpers ======================
def set_app_icon(root: tk.Tk) -> Optional[str]:
    ico = resource_path("windows-updater.ico")
    if os.path.exists(ico):
        try:
            root.iconbitmap(ico)
            return ico
        except Exception:
            pass
    return None

def apply_icon_to_toplevel(tlv: tk.Toplevel, icon_path: Optional[str]):
    if not icon_path:
        return
    try:
        tlv.iconbitmap(icon_path)
    except Exception:
        pass

def load_flag_image() -> Optional[tk.PhotoImage]:
    png = resource_path("kuwait.png")
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except Exception:
            pass
    ico = resource_path("kuwait.ico")
    if os.path.exists(ico):
        try:
            from PIL import Image
            im = Image.open(ico)
            if hasattr(im, "n_frames"):
                im.seek(im.n_frames - 1)
            im = im.convert("RGBA")
            max_h = 18
            if im.height > max_h:
                ratio = max_h / float(im.height)
                im = im.resize((max(16, int(im.width * ratio)), max_h), Image.LANCZOS)
            bio = BytesIO()
            im.save(bio, format="PNG")
            bio.seek(0)
            return tk.PhotoImage(data=bio.read())
        except Exception:
            return None
    return None

def play_success_sound():
    wav = resource_path("success.wav")
    if os.path.exists(wav):
        try:
            winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
            return
        except Exception:
            pass
    try:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass

# ====================== Donate image (drawn at runtime) ======================
def make_donate_image(width=160, height=44):
    radius = height // 2
    top = (255, 187, 71); mid = (247, 162, 28); bot = (225, 140, 22)
    im = Image.new("RGBA", (width, height), (0, 0, 0, 0)); dr = ImageDraw.Draw(im)
    for y in range(height):
        if y < height * 0.6:
            t = y / (height * 0.6); col = tuple(int(top[i] * (1 - t) + mid[i] * t) for i in range(3)) + (255,)
        else:
            t = (y - height * 0.6) / (height * 0.4); col = tuple(int(mid[i] * (1 - t) + bot[i] * t) for i in range(3)) + (255,)
        dr.line([(0, y), (width, y)], fill=col)
    mask = Image.new("L", (width, height), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, width - 1, height - 1], radius=radius, fill=255)
    im.putalpha(mask)
    highlight = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    ImageDraw.Draw(highlight).rounded_rectangle([2, 2, width - 3, height // 2], radius=radius - 2, fill=(255, 255, 255, 70))
    im = Image.alpha_composite(im, highlight)
    ImageDraw.Draw(im).rounded_rectangle([0.5, 0.5, width - 1.5, height - 1.5], radius=radius, outline=(200, 120, 20, 255), width=2)
    bio = BytesIO(); im.save(bio, format="PNG"); bio.seek(0)
    return tk.PhotoImage(data=bio.read())

# ====================== winget helpers ======================
def run(cmd):
    env = os.environ.copy(); env["DOTNET_CLI_UI_LANGUAGE"] = "en"
    p = subprocess.run(cmd, capture_output=True, text=True, shell=False, encoding="utf-8", errors="replace",
                       env=env, startupinfo=_hidden_startupinfo(), creationflags=CREATE_NO_WINDOW)
    return p.returncode, p.stdout.strip(), p.stderr.strip()

def try_json_parsers(include_unknown: bool):
    base = ["--accept-source-agreements", "--disable-interactivity", "--output", "json"]
    flag = ["--include-unknown"] if include_unknown else []
    attempts = [
        ["winget", "upgrade", *flag, *base],
        ["winget", "list", "--upgrade-available", *base],
        ["winget", "list", "--upgrades", *base],
    ]
    last_err = ""
    for cmd in attempts:
        code, out, err = run(cmd)
        if code == 0 and out:
            try:
                data = json.loads(out)
                return normalize_winget_json(data)
            except Exception as e:
                last_err = f"{err or ''}\nJSON parse error: {e}"
        else:
            last_err = err or "winget returned a non-zero exit code."
    raise RuntimeError(last_err.strip() or "Failed to get JSON from winget.")

def normalize_winget_json(data):
    items = []
    if isinstance(data, list):
        iterable = data
    elif isinstance(data, dict):
        if "Sources" in data:
            iterable = []
            for src in data.get("Sources", []):
                iterable.extend(src.get("Packages", []))
        else:
            iterable = data.get("Packages", [])
    else:
        iterable = []
    for it in iterable:
        name      = it.get("PackageName") or it.get("Name") or ""
        pkg_id    = it.get("PackageIdentifier") or it.get("Id") or ""
        available = it.get("AvailableVersion") or it.get("Available") or ""
        current   = it.get("Version") or it.get("InstalledVersion") or ""
        if name and pkg_id and available:
            items.append({"name": name, "id": pkg_id, "available": available, "current": current})
    return items

def parse_table_upgrade_output(text):
    lines = [ln for ln in text.splitlines() if ln.strip()]
    header_idx = -1
    for i, ln in enumerate(lines):
        if re.search(r"\bName\b", ln) and re.search(r"\bId\b", ln) and re.search(r"\bAvailable\b", ln):
            header_idx = i; break
    if header_idx < 0 or header_idx + 1 >= len(lines): return []
    start = header_idx + 1
    if start < len(lines) and re.match(r"^[\s\-]+$", lines[start].replace(" ", "")): start += 1
    items = []
    for ln in lines[start:]:
        if "No applicable updates" in ln: return []
        parts = re.split(r"\s{2,}", ln.rstrip())
        if len(parts) < 4: continue
        if len(parts) >= 5:
            name, pkg_id = parts[0], parts[1]; current = parts[2] if len(parts) > 2 else ""; available = parts[3] if len(parts) > 3 else ""
        else:
            name, pkg_id = parts[0], parts[1]; current, available = "", parts[2]
        if name and pkg_id and available and not name.startswith("-"):
            items.append({"name": name, "id": pkg_id, "current": current, "available": available})
    return items

def get_winget_upgrades(include_unknown: bool):
    code, _, _ = run(["winget", "--version"])
    if code != 0: raise RuntimeError("winget not found. Install the App Installer from Microsoft Store.")
    try:
        return try_json_parsers(include_unknown)
    except Exception as e_json:
        cmd = ["winget", "upgrade", "--accept-source-agreements", "--disable-interactivity"]
        if include_unknown: cmd.insert(2, "--include-unknown")
        code, out, err = run(cmd)
        if code != 0: raise RuntimeError((err or str(e_json)).strip())
        parsed = parse_table_upgrade_output(out)
        if parsed: return parsed
        raise RuntimeError(str(e_json))

# ====================== Checkbox images ======================
def make_checkbox_images(size: int = 16):
    unchecked = tk.PhotoImage(width=size, height=size)
    unchecked.put("white", to=(0, 0, size, size))
    border = "gray20"
    unchecked.put(border, to=(0, 0, size, 1))
    unchecked.put(border, to=(0, size - 1, size, size))
    unchecked.put(border, to=(0, 0, 1, size))
    unchecked.put(border, to=(size - 1, 0, size, size))

    checked = tk.PhotoImage(width=size, height=size)
    checked.tk.call(checked, "copy", unchecked)
    mark = "#2e7d32"
    pts = [(3, size // 2), (4, size // 2 + 1), (5, size // 2 + 2),
           (6, size // 2 + 3), (7, size // 2 + 2),
           (8, size // 2 + 1), (9, size // 2), (10, size // 2 - 1)]
    for (x, y) in pts:
        checked.put(mark, to=(x, y, x + 1, y + 1))
        checked.put(mark, to=(x, y - 1, x + 1, y))
    return unchecked, checked

# ====================== UI Class ======================
class WingetUpdaterUI:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME_VERSION)

        # Window size (adjusted when log is hidden)
        self.root.geometry(f"{WIN_W}x{WIN_H_FULL}")
        self.root.minsize(WIN_W, WIN_H_COMPACT)
        self.root.maxsize(WIN_W, WIN_H_FULL)

        # Styles
        self.style = ttk.Style(self.root)
        self.style.configure("Big.TButton", padding=(14, 8), font=("Segoe UI", 10, "bold"))
        self.style.configure("Big.TCheckbutton", padding=(8, 6), font=("Segoe UI", 10))

        self.updating = False
        self.cancel_requested = False
        self.current_proc = None
        self.loading_win = None
        self.window_icon_path = set_app_icon(self.root)

        self.img_unchecked, self.img_checked = make_checkbox_images(16)
        self.checked_items = set()
        self.id_to_item: Dict[str, str] = {}
        self.pkg_downloads: Dict[str, List[str]] = {}

        # ===== Header =====
        header = ttk.Frame(self.root); header.pack(fill="x", pady=(10, 0))
        ttk.Label(header, text=APP_NAME_VERSION, font=("Segoe UI", 18, "bold")).pack(side="left", padx=12)

        right = ttk.Frame(header); right.pack(side="right", padx=12)
        self.btn_admin = ttk.Button(right, text="Run as Admin", command=self.run_as_admin, style="Big.TButton")
        self.btn_admin.pack(side="right")
        if is_admin():
            self.btn_admin.config(text="Running as Admin", state="disabled")
        else:
            ToolTip(self.btn_admin, "Run the app as admin if you like to install all apps silently")

        # ===== Top controls (same row) =====
        top = ttk.Frame(self.root); top.pack(fill="x", padx=12, pady=6)
        btn_text_w = max(len("Check for Updates"), len("Update Selected")) + 2

        self.btn_check = ttk.Button(top, text="Check for Updates",
                                    command=self.check_for_updates_async, style="Big.TButton", width=btn_text_w)
        self.btn_check.pack(side="left")

        self.include_unknown_var = tk.BooleanVar(value=False)
        self.chk_unknown = ttk.Checkbutton(top, text="Include unknown apps",
                                           variable=self.include_unknown_var, style="Big.TCheckbutton")
        self.chk_unknown.pack(side="left", padx=(12, 0))

        self.btn_update = ttk.Button(top, text="Update Selected",
                                     command=self.update_selected_async, style="Big.TButton", width=btn_text_w)
        self.btn_update.pack(side="left", padx=(12, 0))

        # Right-side temp buttons
        self.btn_open_temp  = ttk.Button(top, text="Open Temp",  command=self.open_temp,       style="Big.TButton")
        self.btn_clear_temp = ttk.Button(top, text="Clear Temp", command=self.clear_temp_async, style="Big.TButton")
        self.btn_open_temp.pack(side="right", padx=(10, 0))
        self.btn_clear_temp.pack(side="right", padx=(10, 0))

        # ===== Second row: Select All/None + counter =====
        sel_row = ttk.Frame(self.root); sel_row.pack(fill="x", padx=12, pady=(0, 4))
        self.btn_sel_all  = ttk.Button(sel_row, text="Select All",  command=self.select_all,  style="Big.TButton")
        self.btn_sel_none = ttk.Button(sel_row, text="Select None", command=self.select_none, style="Big.TButton")
        self.btn_sel_all.pack(side="left")
        self.btn_sel_none.pack(side="left", padx=(6, 0))
        self.counter_var = tk.StringVar(value="0 apps found • 0 selected")
        ttk.Label(sel_row, textvariable=self.counter_var).pack(side="left", padx=(12, 0))

        # Controls to disable during update (except Update which becomes Cancel)
        self._controls_to_disable = [
            self.btn_check, self.chk_unknown, self.btn_sel_all, self.btn_sel_none,
            self.btn_open_temp, self.btn_clear_temp,
        ]

        # ===== Apps list (shorter area) =====
        tree_wrap = ttk.Frame(self.root, height=240)
        tree_wrap.pack(fill="x", expand=False, padx=12, pady=(8, 8))
        tree_wrap.pack_propagate(False)

        cols = ("Name", "Id", "Current", "Available", "Result")
        self.fixed_cols = cols
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="tree headings", height=8, selectmode="none")

        self.tree.heading("#0",       text="Select",   anchor="center")
        self.tree.heading("Name",     text="Name",     anchor="w")
        self.tree.heading("Id",       text="Id",       anchor="w")
        self.tree.heading("Current",  text="Current",  anchor="center")
        self.tree.heading("Available",text="Available",anchor="center")
        self.tree.heading("Result",   text="Result",   anchor="center")

        font = tkfont.nametofont("TkDefaultFont")
        sel_w = max(50, font.measure("Select") + 18)
        self.tree.column("#0", width=sel_w, minwidth=sel_w, anchor="center", stretch=False)
        self.tree.column("Name",      width=260, minwidth=160, anchor="w",      stretch=False)
        self.tree.column("Id",        width=340, minwidth=220, anchor="w",      stretch=False)
        self.tree.column("Current",   width=90,  minwidth=70,  anchor="center", stretch=False)
        self.tree.column("Available", width=100, minwidth=80,  anchor="center", stretch=False)
        self.tree.column("Result",    width=120, minwidth=100, anchor="center", stretch=False)
        self._col_caps = {"Name": 320, "Id": 420, "Current": 120, "Available": 140, "Result": 160}

        self.tree.tag_configure("ok",      background="#e8f5e9")
        self.tree.tag_configure("fail",    background="#ffebee")
        self.tree.tag_configure("skip",    background="#fff8e1")
        self.tree.tag_configure("cancel",  background="#fff3e0")
        self.tree.tag_configure("checked", background="#e3f2fd")

        ysb = ttk.Scrollbar(tree_wrap, orient="vertical",   command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # Mouse behavior
        self._block_header_drag = False
        self._block_resize_select = False
        self.tree.bind("<Button-1>", self._on_mouse_down, add="+")
        self.tree.bind("<B1-Motion>", self._on_mouse_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_mouse_up, add="+")
        self.tree.bind("<Double-Button-1>", self._on_double_click_header, add="+")
        self.tree.bind("<Configure>", self._on_tree_configure, add="+")
        self.tree["displaycolumns"] = self.fixed_cols

        # Context menu for per-app downloads
        self.row_menu = tk.Menu(self.root, tearoff=0)
        self.row_menu.add_command(label="Open downloaded file(s)", command=self._menu_open_downloads)
        self.row_menu.add_command(label="Delete downloaded file(s)", command=self._menu_delete_downloads)
        self._menu_item = None
        def _show_row_menu(event, tree=self.tree):
            if self.updating: return
            item = tree.identify_row(event.y)
            if not item: return
            self._menu_item = item
            tree.selection_set(item)
            try: self.row_menu.tk_popup(event.x_root, event.y_root)
            finally: self.row_menu.grab_release()
        self.tree.bind("<Button-3>", _show_row_menu)

        # ===== Progress bars =====
        pb_wrap = ttk.Frame(self.root); pb_wrap.pack(fill="x", padx=12, pady=(0, 4))
        self.pb_label = ttk.Label(pb_wrap, text="Update"); self.pb_label.pack(side="left")
        self.pb = ttk.Progressbar(pb_wrap, orient="horizontal", mode="determinate")
        self.pb.pack(fill="x", expand=True, padx=10)

        pb2_wrap = ttk.Frame(self.root); pb2_wrap.pack(fill="x", padx=12, pady=(0, 8))
        self.pb2_label = ttk.Label(pb2_wrap, text="Download"); self.pb2_label.pack(side="left")
        self.pb2 = ttk.Progressbar(pb2_wrap, orient="horizontal", mode="determinate", maximum=100, value=0)
        self.pb2.pack(fill="x", expand=True, padx=10)

        # Spinner state
        self._spin_job = None
        self._spin_frames = ["|", "/", "-", "\\"]
        self._spin_index = 0
        self._spin_base = "Downloading"
        self._spin_name = None
        self._spin_pct = 0

        # ===== About (centered + clickable link) =====
        about = ttk.Frame(self.root); about.pack(fill="x", padx=12, pady=(2, 0))
        center = ttk.Frame(about); center.pack(expand=True)
        ttk.Label(center, text="Windows App Updater is a freeware Python App based on Windows Winget to update applications",
                  justify="center", anchor="center").pack()
        ttk.Label(center, text="Version 1.2.1 - Sep 2025", justify="center", anchor="center").pack()

        row = ttk.Frame(center); row.pack()
        ttk.Label(row, text="Author: ilukezippo (BoYaqoub) – ilukezippo@gmail.com").pack(side="left")
        self.flag_img = load_flag_image()
        if self.flag_img:
            tk.Label(row, image=self.flag_img).pack(side="left", padx=(6, 0))

        link_row = ttk.Frame(center); link_row.pack(pady=(2, 0))
        ttk.Label(link_row, text="Info and Latest Updates at ").pack(side="left")
        link = tk.Label(link_row, text="https://github.com/ilukezippo/Windows-App-Updater",
                        fg="#1a73e8", cursor="hand2", font=("Segoe UI", 9, "underline"))
        link.pack(side="left")
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ilukezippo/Windows-App-Updater"))

        # Optional donate button to the right
        self.donate_img = make_donate_image(width=160, height=44)
        tk.Button(about, image=self.donate_img, text="Donate", compound="center",
                  font=("Segoe UI", 11, "bold"), fg="#0f3462", activeforeground="#0f3462",
                  bd=0, highlightthickness=0, cursor="hand2", relief="flat",
                  command=lambda: webbrowser.open(GITHUB_RELEASES_PAGE)).pack(side="right")

        # ===== Log header with Show/Hide on the right =====
        log_header = ttk.Frame(self.root); log_header.pack(fill="x", padx=12, pady=(6, 2))
        ttk.Label(log_header, text="Update Log:", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.btn_toggle_log = ttk.Button(log_header, text="Hide Log", command=self.toggle_log, style="Big.TButton")
        self.btn_toggle_log.pack(side="right")

        # ===== Log =====
        self.log_wrap = ttk.Frame(self.root)
        self.log_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        self.log_box = tk.Text(self.log_wrap, height=10, wrap="none", font=("Consolas", 10))
        log_ysb = ttk.Scrollbar(self.log_wrap, orient="vertical",   command=self.log_box.yview)
        log_xsb = ttk.Scrollbar(self.log_wrap, orient="horizontal", command=self.log_box.xview)
        self.log_box.configure(yscrollcommand=log_ysb.set, xscrollcommand=log_xsb.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_ysb.grid(row=0, column=1, sticky="ns")
        log_xsb.grid(row=1, column=0, sticky="ew")
        self.log_wrap.rowconfigure(0, weight=1)
        self.log_wrap.columnconfigure(0, weight=1)
        self.log_visible = True

        # include toggle in disable list
        self._controls_to_disable.append(self.btn_toggle_log)

        self.root.after(0, self.center_on_screen)
        self.root.after(1200, self.check_latest_app_version_async)

    # ----- GitHub latest release check -----
    def _parse_ver_tuple(self, v: str):
        nums = re.findall(r"\d+", v)
        return tuple(int(n) for n in nums[:4]) or (0,)

    def check_latest_app_version_async(self):
        def worker():
            try:
                req = urllib.request.Request(GITHUB_API_LATEST, headers={"User-Agent": "Windows-App-Updater"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.loads(r.read().decode("utf-8", "replace"))
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                if not tag:
                    return
                if self._parse_ver_tuple(tag) > self._parse_ver_tuple(APP_VERSION_ONLY):
                    def prompt():
                        if messagebox.askyesno(
                            "New Version Available",
                            f"A newer version {tag} is available.\nOpen releases page?"
                        ):
                            webbrowser.open(GITHUB_RELEASES_PAGE)
                    self.root.after(0, prompt)
            except Exception:
                pass
        threading.Thread(target=worker, daemon=True).start()

    # ----- log show/hide (also resize window) -----
    def toggle_log(self):
        if self.log_visible:
            self.log_wrap.forget()
            self.btn_toggle_log.config(text="Show Log")
            self.root.geometry(f"{WIN_W}x{WIN_H_COMPACT}")
            self.log_visible = False
        else:
            self.log_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 10))
            self.btn_toggle_log.config(text="Hide Log")
            self.root.geometry(f"{WIN_W}x{WIN_H_FULL}")
            self.log_visible = True
        self.root.update_idletasks()

    # ----- enable/disable controls during update -----
    def _disable_controls_for_update(self):
        self.updating = True
        for w in self._controls_to_disable:
            try: w.config(state="disabled")
            except Exception: pass
        try: self.btn_admin.config(state="disabled")
        except Exception: pass
        self.btn_update.config(text="Cancel", state="normal")
        try:
            self._tree_prev_state = self.tree.cget("state")
        except Exception:
            self._tree_prev_state = "normal"
        try:
            self.tree.config(state="disabled")
        except Exception:
            pass

    def _enable_controls_after_update(self):
        for w in self._controls_to_disable:
            try: w.config(state="normal")
            except Exception: pass
        # Admin button: keep disabled if already admin
        try:
            if is_admin():
                self.btn_admin.config(text="Running as Admin", state="disabled")
            else:
                self.btn_admin.config(text="Run as Admin", state="normal")
        except Exception:
            pass
        try:
            self.tree.config(state=self._tree_prev_state if hasattr(self, "_tree_prev_state") else "normal")
        except Exception:
            pass
        self.btn_update.config(text="Update Selected", state="normal")
        self.updating = False

    # ----- TREE column clamping -----
    def _on_tree_configure(self, _event=None):
        self.root.after_idle(self._fit_columns_to_tree)

    def _fit_columns_to_tree(self):
        try:
            avail = max(0, self.tree.winfo_width() - 2)
        except Exception:
            return
        cols = ["#0", "Name", "Id", "Current", "Available", "Result"]

        widths = {c: int(self.tree.column(c, "width")) for c in cols if c in ("#0",) or c in self.fixed_cols}
        mins   = {c: int(self.tree.column(c, "minwidth") or 20) for c in cols if c in ("#0",) or c in self.fixed_cols}
        caps   = {"#0": widths.get("#0", 60), **self._col_caps}

        for c in cols:
            if c in widths:
                cap = caps.get(c, 10_000)
                if widths[c] > cap:
                    self.tree.column(c, width=cap)
                    widths[c] = cap

        total = sum(widths.get(c, 0) for c in cols if c in widths)
        if total <= avail or avail <= 0:
            return

        overflow = total - avail

        reducible = {}
        for c in cols:
            if c in widths:
                reducible[c] = max(0, widths[c] - mins.get(c, 20))

        order = ["Id", "Name", "Result", "Available", "Current", "#0"]
        total_reducible = sum(reducible.get(c, 0) for c in order)
        if total_reducible == 0:
            return

        for c in order:
            if overflow <= 0:
                break
            r = reducible.get(c, 0)
            if r <= 0:
                continue
            share = int(round(overflow * (r / total_reducible)))
            share = max(1, min(r, share))
            new_w = max(mins.get(c, 20), widths[c] - share)
            self.tree.column(c, width=new_w)
            overflow -= (widths[c] - new_w)

    # ----- Mouse & selection -----
    def _on_mouse_down(self, event):
        if self.updating:
            return "break"
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            self._block_header_drag = True
            return "break"
        if region == "separator":
            col_left = self.tree.identify_column(event.x - 1)
            if col_left == "#0":
                self._block_resize_select = True
                return "break"
            self._block_resize_select = False
            return
        if region in ("tree", "cell"):
            item = self.tree.identify_row(event.y)
            if item:
                self._toggle_row(item)
                return "break"
        self._block_header_drag = False
        self._block_resize_select = False

    def _on_mouse_drag(self, event):
        if self.updating or self._block_header_drag or self._block_resize_select:
            return "break"

    def _on_mouse_up(self, event):
        if self.updating:
            return "break"
        self._block_header_drag = False
        self._block_resize_select = False
        try:
            if tuple(self.tree["displaycolumns"]) != tuple(self.fixed_cols):
                self.tree["displaycolumns"] = self.fixed_cols
        except Exception:
            pass

    def _on_double_click_header(self, event):
        if self.updating:
            return "break"
        if self.tree.identify("region", event.x, event.y) != "separator":
            return
        col_left = self.tree.identify_column(event.x - 1)
        if not col_left or col_left == "#0":
            return
        self.autofit_column(col_left)
        self._fit_columns_to_tree()

    def _toggle_row(self, item: str):
        if not item or self.updating:
            return
        if item in self.checked_items:
            self.checked_items.remove(item)
            self.tree.item(item, image=self.img_unchecked)
            tags = set(self.tree.item(item, "tags") or ())
            tags.discard("checked")
            self.tree.item(item, tags=tuple(tags))
        else:
            self.checked_items.add(item)
            self.tree.item(item, image=self.img_checked)
            tags = set(self.tree.item(item, "tags") or ())
            tags.add("checked")
            self.tree.item(item, tags=tuple(tags))
        self.update_counter()

    def autofit_column(self, col_id: str):
        heading = self.tree.heading(col_id, "text") or ""
        font = tkfont.nametofont("TkDefaultFont")
        pad = 24
        max_px = font.measure(heading)
        for item in self.tree.get_children(""):
            val = self.tree.set(item, col_id) or ""
            px = font.measure(val)
            if px > max_px:
                max_px = px
        minw = int(self.tree.column(col_id, "minwidth") or 20)
        cap = self._col_caps.get(col_id, 360)
        new_w = max(minw, min(max_px + pad, cap))
        self.tree.column(col_id, width=new_w, stretch=False)

    def autofit_all(self):
        for col_id in self.fixed_cols:
            if col_id != "#0":
                self.autofit_column(col_id)
        self._fit_columns_to_tree()

    # ----- window centering -----
    def center_on_screen(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ----- admin & donate -----
    def run_as_admin(self):
        relaunch_as_admin()

    # ====================== Loading screen ======================
    def show_loading(self, text="Loading..."):
        if self.loading_win:
            return
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("")
        self.loading_win.transient(self.root)
        self.loading_win.grab_set()
        self.loading_win.resizable(False, False)
        apply_icon_to_toplevel(self.loading_win, self.window_icon_path)

        ttk.Label(self.loading_win, text=text, font=("Segoe UI", 12, "bold")).pack(padx=20, pady=(16, 8))
        pb = ttk.Progressbar(self.loading_win, mode="indeterminate", length=280)
        pb.pack(padx=20, pady=(0, 16))
        pb.start(10)
        self.loading_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - self.loading_win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - self.loading_win.winfo_height()) // 2
        self.loading_win.geometry(f"+{x}+{y}")
        self.loading_win.protocol("WM_DELETE_WINDOW", lambda: None)

    def hide_loading(self):
        if self.loading_win:
            try:
                self.loading_win.grab_release()
            except Exception:
                pass
            self.loading_win.destroy()
            self.loading_win = None

    # ====================== Progress helpers ======================
    def progress_start(self, phase: str, total: int):
        self.pb_phase = phase
        self.pb_total = max(0, int(total))
        self.pb_value = 0
        self.pb.configure(maximum=max(self.pb_total, 1), value=0, mode="determinate")
        self.pb_label.configure(text=f"Updating: 0/{self.pb_total}")
        self.pb2.configure(value=0, maximum=100, mode="determinate")
        self.pb2_label.configure(text="Downloading: 0%")
        self.root.update_idletasks()

    def progress_step(self, inc: int = 1):
        if self.pb_total <= 0:
            return
        self.pb_value = min(self.pb_total, self.pb_value + inc)
        self.pb.configure(value=self.pb_value)
        self.pb_label.configure(text=f"Updating: {self.pb_value}/{self.pb_total}")
        self.root.update_idletasks()

    def progress_finish(self, canceled=False):
        if getattr(self, "pb_total", 0) > 0:
            self.pb.configure(value=self.pb_total)
            suffix = " (canceled)" if canceled else " (done)"
            self.pb_label.configure(text=f"Update: {self.pb_total}/{self.pb_total}{suffix}")
        else:
            self.pb_label.configure(text="Update")
        self._spinner_stop("Download")
        self.pb2.configure(value=0)
        self.root.update_idletasks()

    def per_app_reset(self, name_or_id: str):
        self.pb2.configure(value=0, maximum=100, mode="determinate")
        self._spinner_start("Downloading", name_or_id)
        self.root.update_idletasks()

    def per_app_update_percent(self, pct: int, name_or_id: Optional[str] = None):
        pct = max(0, min(100, int(pct)))
        self.pb2.configure(value=pct)
        if name_or_id:
            self._spin_name = name_or_id
        self._spinner_set_pct(pct)
        if pct >= 100:
            self._spinner_stop(f"Download ({self._spin_name}): 100%")
        self.root.update_idletasks()

    # ----- spinner helpers -----
    def _spinner_start(self, base_text: str, name: str):
        self._spin_base = base_text
        self._spin_name = name
        self._spin_pct = 0
        self._spin_index = 0
        if self._spin_job is None:
            self._spinner_tick()

    def _spinner_tick(self):
        frame = self._spin_frames[self._spin_index % len(self._spin_frames)]
        self._spin_index += 1
        name = f" ({self._spin_name})" if self._spin_name else ""
        pct  = f" {self._spin_pct}%" if isinstance(self._spin_pct, int) else ""
        self.pb2_label.configure(text=f"{self._spin_base}{name} {frame}{pct}")
        self._spin_job = self.root.after(150, self._spinner_tick)

    def _spinner_set_pct(self, pct: int):
        self._spin_pct = max(0, min(100, int(pct)))

    def _spinner_stop(self, final_text: str = "Download"):
        if self._spin_job is not None:
            try:
                self.root.after_cancel(self._spin_job)
            except Exception:
                pass
            self._spin_job = None
        self.pb2_label.configure(text=final_text)

    # ====================== Selection helpers ======================
    def _iter_items(self):
        return self.tree.get_children("")

    def select_all(self):
        if self.updating:
            return
        for item in self._iter_items():
            self.checked_items.add(item)
            self.tree.item(item, image=self.img_checked)
            tags = set(self.tree.item(item, "tags") or ())
            tags.add("checked")
            self.tree.item(item, tags=tuple(tags))
        self.update_counter()

    def select_none(self):
        if self.updating:
            return
        for item in self._iter_items():
            self.checked_items.discard(item)
            self.tree.item(item, image=self.img_unchecked)
            tags = set(self.tree.item(item, "tags") or ())
            tags.discard("checked")
            self.tree.item(item, tags=tuple(tags))
        self.update_counter()

    def update_counter(self):
        total = len(self.tree.get_children(""))
        selected = len(self.checked_items)
        self.counter_var.set(f"{total} apps found • {selected} selected")

    def clear_tree(self):
        self.checked_items.clear()
        self.id_to_item.clear()
        for i in self._iter_items():
            self.tree.delete(i)

    # ====================== TEMP helpers (per-app + global) ======================
    def _temp_dir(self) -> str:
        return tempfile.gettempdir()

    def _fmt_bytes(self, n: int) -> str:
        for unit in ("B","KB","MB","GB","TB"):
            if n < 1024 or unit == "TB":
                return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
            n /= 1024.0

    def _snapshot_temp(self) -> Dict[str, float]:
        root = self._temp_dir()
        snap: Dict[str, float] = {}
        def add(path):
            try:
                st = os.stat(path)
                if st.st_size >= 1_000_000:
                    snap[path] = st.st_mtime
            except Exception:
                pass
        try:
            for e in os.scandir(root):
                if e.is_file(follow_symlinks=False):
                    add(e.path)
        except Exception:
            pass
        try:
            for e in os.scandir(root):
                if e.is_dir(follow_symlinks=False):
                    for ee in os.scandir(e.path):
                        if ee.is_file(follow_symlinks=False):
                            add(ee.path)
        except Exception:
            pass
        return snap

    def _find_new_installer_files(self, before: Dict[str, float], after: Dict[str, float]) -> List[str]:
        exts = (".exe",".msi",".msix",".msixbundle",".appx",".appxbundle",".zip",".7z",".rar",".cab")
        news: List[str] = []
        for path, mtime in after.items():
            if path not in before or after[path] > before.get(path, 0):
                if os.path.splitext(path)[1].lower() in exts:
                    news.append(path)
        news = sorted(set(news), key=lambda p: after.get(p, 0), reverse=True)
        return news

    def open_temp(self):
        try:
            os.startfile(self._temp_dir())
        except Exception as e:
            messagebox.showerror("Open Temp", str(e))

    def clear_temp_async(self):
        temp_dir = self._temp_dir()
        if not messagebox.askyesno("Clear Temp",
            f"This will delete files and folders inside:\n\n{temp_dir}\n\n"
            "Items currently in use will be skipped.\n\nProceed?"):
            return
        self.show_loading("Clearing Temp...")
        def worker():
            files_del = 0
            dirs_del = 0
            freed = 0
            def onerror(func, path, exc_info):
                try:
                    os.chmod(path, 0o666)
                    func(path)
                except Exception:
                    pass
            try:
                for e in os.scandir(temp_dir):
                    p = e.path
                    try:
                        if e.is_file(follow_symlinks=False):
                            try: freed += os.path.getsize(p)
                            except: pass
                            try:
                                os.remove(p); files_del += 1
                                self.root.after(0, lambda s=f"[Clear Temp] Deleted file: {p}": self.log(s))
                            except Exception: pass
                        elif e.is_dir(follow_symlinks=False):
                            size = 0
                            for dp,_,fns in os.walk(p):
                                for fn in fns:
                                    fp = os.path.join(dp, fn)
                                    try: size += os.path.getsize(fp)
                                    except: pass
                            try:
                                shutil.rmtree(p, onerror=onerror); dirs_del += 1; freed += size
                                self.root.after(0, lambda s=f"[Clear Temp] Deleted folder: {p}": self.log(s))
                            except Exception: pass
                    except Exception: pass
            except Exception as e:
                self.root.after(0, lambda: self.log(f"[Clear Temp] Error: {e}"))
            def done():
                self.hide_loading()
                self.log(f"[Clear Temp] Folders: {dirs_del} | Files: {files_del} | Freed: {self._fmt_bytes(freed)}")
                messagebox.showinfo("Clear Temp",
                    f"Folders deleted: {dirs_del}\nFiles deleted: {files_del}\nSpace freed: {self._fmt_bytes(freed)}")
            self.root.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    # ===== context menu handlers =====
    def _menu_open_downloads(self):
        item = self._menu_item
        if not item or self.updating:
            return
        pkg_id = self.tree.set(item, "Id")
        files = self.pkg_downloads.get(pkg_id) or []
        if not files:
            messagebox.showinfo("Downloads", "No downloaded files recorded for this app.")
            return
        for p in files:
            if os.path.exists(p):
                try: os.startfile(p)
                except Exception: pass
            else:
                folder = os.path.dirname(p) or self._temp_dir()
                try: os.startfile(folder)
                except Exception: pass

    def _menu_delete_downloads(self):
        item = self._menu_item
        if not item or self.updating:
            return
        pkg_id = self.tree.set(item, "Id")
        files = [p for p in (self.pkg_downloads.get(pkg_id) or []) if p and os.path.exists(p)]
        if not files:
            messagebox.showinfo("Delete Downloads", "No downloaded files to delete for this app.")
            return
        if not messagebox.askyesno("Delete Downloads",
            "Delete the downloaded file(s) for this app?\n\n" + "\n".join(files)):
            return
        deleted = 0
        freed = 0
        for p in files:
            try:
                sz = 0
                try: sz = os.path.getsize(p)
                except: pass
                os.remove(p)
                deleted += 1
                freed += sz
                self.log(f"[Delete] {p}")
            except Exception as e:
                self.log(f"[Delete] Could not delete {p}: {e}")
        self.pkg_downloads[pkg_id] = [p for p in self.pkg_downloads.get(pkg_id, []) if os.path.exists(p)]
        self.log(f"[Delete] Deleted {deleted} file(s), freed {self._fmt_bytes(freed)}")

    # ====================== Check for updates (async) ======================
    def check_for_updates_async(self):
        include_unknown = bool(self.include_unknown_var.get())
        self.btn_check.config(state="disabled")
        self.show_loading("Checking for updates...")

        def worker():
            try:
                pkgs = get_winget_upgrades(include_unknown=include_unknown)
            except Exception as e:
                self.root.after(0, lambda: (
                    self.hide_loading(),
                    self.btn_check.config(state="normal"),
                    self.counter_var.set("0 apps found • 0 selected"),
                    messagebox.showerror("winget error", f"Failed to query updates:\n{e}"),
                    self.log(f"[winget] {e}")
                ))
                return
            self.root.after(0, lambda: self.populate_tree(pkgs))

        threading.Thread(target=worker, daemon=True).start()

    def populate_tree(self, pkgs):
        self.hide_loading()
        self.clear_tree()
        self.btn_check.config(state="normal")
        if not pkgs:
            self.counter_var.set("0 apps found • 0 selected")
            self.log("No apps need updating.")
            return

        for p in pkgs:
            item = self.tree.insert(
                "", "end",
                text="",
                image=self.img_unchecked,
                values=(p["name"], p["id"], p.get("current", ""), p.get("available", ""), ""),
            )
            self.id_to_item[p["id"]] = item
            self.checked_items.discard(item)

        self.update_counter()
        self.autofit_all()

    # ====================== Update selected (async + Cancel) ======================
    def update_selected_async(self):
        if getattr(self, "updating", False):
            self.cancel_requested = True
            self.btn_update.config(text="Cancelling...", state="disabled")
            proc = self.current_proc
            if proc and proc.poll() is None:
                try:
                    proc.terminate()
                except Exception:
                    pass
            return

        targets: List[Tuple[str, str, str]] = []
        for item in list(self.checked_items):
            pkg_id  = self.tree.set(item, "Id")
            current = (self.tree.set(item, "Current") or "").strip()
            if pkg_id:
                targets.append((pkg_id, current, item))

        if not targets:
            messagebox.showinfo("No Selection", "No apps selected for update.")
            return

        self._disable_controls_for_update()

        self.cancel_requested = False
        self.current_proc = None
        self.log(f"Starting updates for {len(targets)} package(s)...")

        results: Dict[str, str] = {}
        self.progress_start("Updating", len(targets))

        def worker():
            percent_re = re.compile(r"(\d{1,3})\s*%")
            size_re = re.compile(r"([\d\.]+)\s*(KB|MB|GB)\s*/\s*([\d\.]+)\s*(KB|MB|GB)", re.I)
            unit_mult = {"KB": 1_000, "MB": 1_000_000, "GB": 1_000_000_000}
            spinner_re = re.compile(r"^[\s\\/\|\-\r\u2580-\u259F\u2500-\u257F]+$")

            for pkg_id, current, item in targets:
                if self.cancel_requested:
                    break

                name_for_label = self.tree.set(item, "Name") or pkg_id
                self.root.after(0, lambda n=name_for_label: self.per_app_reset(n))
                self.root.after(0, lambda pid=pkg_id: self.log(f"Updating {pid} ..."))

                try:
                    cmd = [
                        "winget", "upgrade", "--id", pkg_id,
                        "--accept-package-agreements", "--accept-source-agreements",
                        "--disable-interactivity", "-h"
                    ]
                    if self.include_unknown_var.get() or (not current) or (current.lower() == "unknown"):
                        cmd.insert(2, "--include-unknown")

                    env = os.environ.copy()
                    env["DOTNET_CLI_UI_LANGUAGE"] = "en"

                    snap_before = self._snapshot_temp()

                    self.current_proc = subprocess.Popen(
                        cmd,
                        shell=False, text=True,
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        encoding="utf-8", errors="replace", env=env,
                        startupinfo=_hidden_startupinfo(), creationflags=CREATE_NO_WINDOW
                    )

                    captured_lines: List[str] = []
                    last_pct = -1

                    while True:
                        if self.cancel_requested and self.current_proc and self.current_proc.poll() is None:
                            try:
                                self.current_proc.terminate()
                            except Exception:
                                pass
                            def mark_canceled(it=item):
                                self.tree.set(it, "Result", "🟧 Canceled")
                                self.tree.item(it, tags=("cancel",))
                                self._spinner_stop("Download")
                            self.root.after(0, mark_canceled)
                            break

                        line = self.current_proc.stdout.readline()
                        if not line:
                            break
                        ln = line.rstrip()
                        if not ln or spinner_re.match(ln):
                            continue

                        captured_lines.append(ln)
                        self.root.after(0, lambda s=ln: self.log(s))

                        m = None
                        for m in percent_re.finditer(ln):
                            pass
                        if m:
                            try:
                                pct = int(m.group(1))
                                if pct != last_pct:
                                    last_pct = pct
                                    self.root.after(0, lambda p=pct, n=name_for_label: self.per_app_update_percent(p, n))
                                continue
                            except Exception:
                                pass
                        m2 = size_re.search(ln)
                        if m2:
                            try:
                                have_val, have_u, tot_val, tot_u = m2.groups()
                                have = float(have_val) * unit_mult[have_u.upper()]
                                tot  = float(tot_val)  * unit_mult[tot_u.upper()]
                                if tot > 0:
                                    pct = max(0, min(100, int(round((have / tot) * 100))))
                                    if pct != last_pct:
                                        last_pct = pct
                                        self.root.after(0, lambda p=pct, n=name_for_label: self.per_app_update_percent(p, n))
                            except Exception:
                                pass

                    if self.current_proc and self.current_proc.poll() is None:
                        out, err = self.current_proc.communicate()
                    else:
                        out = err = ""
                        try:
                            out = (self.current_proc.stdout.read() or "") if self.current_proc and self.current_proc.stdout else ""
                            err = (self.current_proc.stderr.read() or "") if self.current_proc and self.current_proc.stderr else ""
                        except Exception:
                            pass

                    if err:
                        self.root.after(0, lambda e=err: self.log(e.strip()))

                    snap_after = self._snapshot_temp()
                    new_files = self._find_new_installer_files(snap_before, snap_after)
                    if new_files:
                        self.pkg_downloads.setdefault(pkg_id, [])
                        existing = set(self.pkg_downloads[pkg_id])
                        for p in new_files:
                            if p not in existing:
                                self.pkg_downloads[pkg_id].append(p)
                        self.root.after(0, lambda nf=new_files, pid=pkg_id:
                            self.log(f"[Downloads] {pid}: " + " | ".join(nf)))

                    if self.cancel_requested and (self.current_proc is None or self.current_proc.returncode is not None):
                        results[pkg_id] = "canceled"
                    else:
                        rc = self.current_proc.returncode if self.current_proc else 1
                        joined = "\n".join(captured_lines + ([out] if out else []))
                        text_low = (joined + "\n" + (err or "")).lower()

                        if rc != 0 or "failed" in text_low or "error" in text_low or "0x" in text_low:
                            status, tag = "❌ Failed", "fail"
                            results[pkg_id] = "failed"
                        elif "no applicable update" in text_low or "no packages found" in text_low:
                            status, tag = "⏭ No update", "skip"
                            results[pkg_id] = "skipped"
                        else:
                            status, tag = "✅ Success", "ok"
                            results[pkg_id] = "success"
                            if self.pkg_downloads.get(pkg_id):
                                status = "✅ Success (downloads)"

                        def apply_result(it=item, st=status, tg=tag):
                            if self.tree.set(it, "Result") in ("", None):
                                self.tree.set(it, "Result", st)
                                self.tree.item(it, tags=(tg,))
                        self.root.after(0, apply_result)

                except Exception as ex:
                    if self.cancel_requested:
                        results[pkg_id] = "canceled"
                        def apply_canceled(it=item):
                            self.tree.set(it, "Result", "🟧 Canceled")
                            self.tree.item(it, tags=("cancel",))
                            self._spinner_stop("Download")
                        self.root.after(0, apply_canceled)
                    else:
                        results[pkg_id] = "failed"
                        self.root.after(0, lambda ex=ex: self.log(f"Error: {ex}"))
                        def apply_result_err(it=item):
                            self.tree.set(it, "Result", "❌ Failed")
                            self.tree.item(it, tags=("fail",))
                        self.root.after(0, apply_result_err)
                finally:
                    self.root.after(0, lambda pid=pkg_id: self.log(f"✔ Finished {pid}"))
                    if not self.cancel_requested:
                        self.root.after(0, lambda n=name_for_label: self.per_app_update_percent(100, n))
                    self.root.after(0, lambda: self.progress_step(1))

            def done():
                canceled_overall = self.cancel_requested
                if canceled_overall:
                    for _, _, item in targets:
                        if not self.tree.set(item, "Result"):
                            self.tree.set(item, "Result", "🟧 Canceled")
                            self.tree.item(item, tags=("cancel",))
                    self.log("Cancelled by user.")

                ok = sum(1 for s in results.values() if s == "success")
                fail = sum(1 for s in results.values() if s == "failed")
                skip = sum(1 for s in results.values() if s == "skipped")
                canc = sum(1 for s in results.values() if s == "canceled")
                self.log(f"Summary → ✅ {ok} success • ❌ {fail} failed • ⏭ {skip} skipped • 🟧 {canc} canceled")

                if fail == 0 and not canceled_overall:
                    play_success_sound()

                self.cancel_requested = False
                self.current_proc = None

                self._enable_controls_after_update()
                self.progress_finish(canceled=canceled_overall)
            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ====================== Logging ======================
    def log(self, text: str):
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

# ====================== main ======================
if __name__ == "__main__":
    root = tk.Tk()
    app = WingetUpdaterUI(root)
    root.mainloop()
