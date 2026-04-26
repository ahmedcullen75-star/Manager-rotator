import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess, threading, psutil, time, os, json, sys, ctypes
import hashlib, uuid, csv, io, urllib.request, urllib.parse

# ── Version ───────────────────────────────────────────────────────────────────
CURRENT_VERSION = "1.0.1"

# ── GitHub Auto Update ────────────────────────────────────────────────────────
GITHUB_USER  = "ahmedcullen75-star"
GITHUB_REPO  = "Manager-rotator"
GITHUB_BRANCH = "main"
VERSION_URL = f"https://raw.githubusercontent.com/ahmedcullen75-star/Manager-rotator/main/version.txt"
EXE_URL     = f"https://github.com/ahmedcullen75-star/Manager-rotator/releases/latest/download/phBot.Rotator.exe"

# ── Colors ────────────────────────────────────────────────────────────────────
DARK   = "#1a1a2e"
PANEL  = "#16213e"
CARD   = "#0f3460"
ACCENT = "#e94560"
GREEN  = "#0f9b58"
AMBER  = "#f5a623"
BLUE   = "#5bc0de"
TEXT   = "#eaeaea"
MUTED  = "#8892a4"
WHITE  = "#ffffff"

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rotator_config.json")

# Google Sheets IDs
SHEET_ID  = "1eFS6YXDNBpY4U-Lsxqn0V_OMtBzJtvRnHExm5gffeyc"
SHEET_GID = "0"

# Column indices in the sheet (0-based)
COL_USER    = 0
COL_PASS    = 1
COL_MAXDEV  = 2
COL_EXPIRE  = 3
COL_HWID1   = 4

ADMIN_EMAIL   = "ahmedcullen75@gmail.com"
ADMIN_DISCORD = "nightfall0234"

APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbydSpZhLbhT4gJX7_6mMLS6gJUOMvu7WPcuCvJP9a35Vo7h-JOsS46Dpub09JhJAgC0yg/exec"

# ── Auto Update ───────────────────────────────────────────────────────────────
def check_for_update():
    try:
        req = urllib.request.Request(VERSION_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as r:
            latest = r.read().decode().strip()
        if latest != CURRENT_VERSION:
            return True, latest
    except Exception:
        pass
    return False, CURRENT_VERSION


def do_update():
    try:
        req = urllib.request.Request(EXE_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            new_code = r.read()

        # Path of the running exe
        current_exe = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
        exe_dir     = os.path.dirname(current_exe)
        exe_name    = os.path.basename(current_exe)
        new_exe     = os.path.join(exe_dir, "_update_" + exe_name)

        # Write new file next to current exe with a temp name
        with open(new_exe, "wb") as f:
            f.write(new_code)

        # Create a batch script that:
        # 1. Waits for current process to exit
        # 2. Replaces old exe with new one
        # 3. Launches new exe
        # 4. Deletes itself
        bat_path = os.path.join(exe_dir, "_updater.bat")
        pid      = os.getpid()
        bat_content = f"""@echo off
:WAIT
tasklist /FI "PID eq {pid}" 2>NUL | find "{pid}" >NUL
if not errorlevel 1 (
    timeout /t 1 /nobreak >NUL
    goto WAIT
)
move /Y "{new_exe}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
"""
        with open(bat_path, "w") as f:
            f.write(bat_content)

        messagebox.showinfo(
            "Update Downloaded",
            "Update downloaded!\n\nThe program will restart automatically.")

        subprocess.Popen(["cmd.exe", "/c", bat_path],
                         creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit()

    except Exception as e:
        messagebox.showerror("Update Failed", f"Could not download update:\n{e}")


def show_update_prompt(latest_version):
    win = tk.Tk()
    win.withdraw()
    while True:
        answer = messagebox.askyesno(
            "Update Required",
            f"A new version is required to continue!\n\n"
            f"  Current version : {CURRENT_VERSION}\n"
            f"  New version     : {latest_version}\n\n"
            f"You must update to use this program.\n"
            f"Update now?")
        if answer:
            win.destroy()
            do_update()
            sys.exit()
        else:
            messagebox.showwarning(
                "Update Required",
                "You cannot use this program without updating.\n\n"
                "Please press Yes to update.")
    win.destroy()
    sys.exit()

# ── Helpers ───────────────────────────────────────────────────────────────────
from datetime import datetime, date

def now_sec():
    n = datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second

def now_str():
    return datetime.now().strftime("%H:%M:%S")

def sec_to_hms(s):
    s = int(s)
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    return h, m, sec

def get_hwid():
    raw = str(uuid.getnode()).encode()
    return hashlib.sha256(raw).hexdigest()[:16].upper()

def force_kill(pid):
    try:
        parent   = psutil.Process(pid)
        children = parent.children(recursive=True)
    except Exception:
        children = []
    for c in children:
        try:
            subprocess.run(["taskkill", "/PID", str(c.pid), "/F"], capture_output=True)
        except Exception:
            pass
    try:
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    except Exception:
        pass

# ── Google Sheets API ─────────────────────────────────────────────────────────
def _fetch_sheet_csv():
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/export?format=csv&gid={SHEET_GID}")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as r:
        return r.read().decode("utf-8")

def _parse_csv(raw):
    reader = csv.reader(io.StringIO(raw))
    rows   = list(reader)
    return rows[1:] if rows else []

def _write_hwid(username, hwid):
    params = urllib.parse.urlencode({"action": "addHwid",
                                     "username": username,
                                     "hwid": hwid})
    url = f"{APPS_SCRIPT_URL}?{params}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as r:
        resp = r.read().decode()
    return "ok" in resp.lower()

def authenticate(username, password):
    hwid = get_hwid()
    try:
        raw  = _fetch_sheet_csv()
        rows = _parse_csv(raw)
    except Exception as e:
        return False, f"Cannot reach Google Sheets.\nCheck your internet connection.\n\n({e})"

    for row in rows:
        if len(row) < 2:
            continue
        if row[COL_USER].strip() == username.strip() and \
           row[COL_PASS].strip() == password.strip():

            if len(row) > COL_EXPIRE and row[COL_EXPIRE].strip():
                try:
                    exp = datetime.strptime(row[COL_EXPIRE].strip(), "%Y-%m-%d").date()
                    if date.today() > exp:
                        return False, (
                            "Your license has expired.\n\n"
                            "Please contact admin to renew:\n"
                            f"  Email:   {ADMIN_EMAIL}\n"
                            f"  Discord: {ADMIN_DISCORD}")
                except ValueError:
                    pass

            max_dev  = int(row[COL_MAXDEV].strip()) if (len(row) > COL_MAXDEV and
                           row[COL_MAXDEV].strip().isdigit()) else 1
            reg_hwids = [row[i].strip() for i in range(COL_HWID1, len(row))
                         if i < len(row) and row[i].strip()]

            if hwid in reg_hwids:
                return True, ""

            if len(reg_hwids) >= max_dev:
                return False, (
                    "This account is already in use on another device.\n\n"
                    "Contact admin to reset or upgrade your license:\n"
                    f"  Email:   {ADMIN_EMAIL}\n"
                    f"  Discord: {ADMIN_DISCORD}")

            try:
                _write_hwid(username, hwid)
            except Exception:
                pass
            return True, ""

    return False, "Incorrect username or password."

# ── Login Window ──────────────────────────────────────────────────────────────
class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("phBot Rotator — Login")
        self.geometry("400x320")
        self.resizable(False, False)
        self.configure(bg=DARK)
        self.result   = False
        self.username = ""
        self._build()

    def _build(self):
        tk.Label(self, text="phBot Manager Rotator",
                 bg=DARK, fg=WHITE,
                 font=("Segoe UI", 15, "bold")).pack(pady=(30, 4))
        tk.Label(self, text="Please log in to continue",
                 bg=DARK, fg=MUTED,
                 font=("Segoe UI", 9)).pack(pady=(0, 24))

        def field(lbl_text, show=""):
            tk.Label(self, text=lbl_text, bg=DARK, fg=MUTED,
                     font=("Segoe UI", 9), anchor="w").pack(
                fill="x", padx=50, pady=(0, 2))
            e = tk.Entry(self, show=show, width=28,
                         bg=PANEL, fg=WHITE, insertbackground=WHITE,
                         font=("Segoe UI", 11), bd=0,
                         highlightthickness=1, highlightbackground=CARD,
                         highlightcolor=ACCENT)
            e.pack(padx=50, pady=(0, 12), ipady=6)
            return e

        self.ent_user = field("Username")
        self.ent_pass = field("Password", show="•")

        self.lbl_err = tk.Label(self, text="", bg=DARK, fg=ACCENT,
                                font=("Segoe UI", 8), wraplength=300)
        self.lbl_err.pack(pady=(8, 0))

        self.btn_login = tk.Button(self, text="Login",
                                   command=self._login,
                                   bg=GREEN, fg=WHITE,
                                   activebackground=GREEN, activeforeground=WHITE,
                                   font=("Segoe UI", 11, "bold"),
                                   bd=0, padx=40, pady=8, cursor="hand2")
        self.btn_login.pack(pady=(12, 0))

        self.ent_user.focus()
        self.bind("<Return>", lambda e: self._login())

    def _login(self):
        user = self.ent_user.get().strip()
        pw   = self.ent_pass.get().strip()

        if not user or not pw:
            self.lbl_err.configure(text="Enter username and password.")
            return

        self.btn_login.configure(state="disabled", text="Checking...")
        self.lbl_err.configure(text="")
        self.update()

        def _check():
            ok, msg = authenticate(user, pw)
            if ok:
                self.result   = True
                self.username = user
                self.after(0, self.destroy)
            else:
                self.after(0, lambda: self.lbl_err.configure(text=msg))
                self.after(0, lambda: self.btn_login.configure(
                    state="normal", text="Login"))

        threading.Thread(target=_check, daemon=True).start()

# ── Slot Dialog ───────────────────────────────────────────────────────────────
class SlotDialog(tk.Toplevel):
    def __init__(self, parent, slot=None):
        super().__init__(parent)
        self.title("Scheduled Slot")
        self.configure(bg=DARK)
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        def lbl(text, row):
            tk.Label(self, text=text, bg=DARK, fg=MUTED,
                     font=("Segoe UI", 9)).grid(
                row=row, column=0, sticky="w", padx=14, pady=(10, 2))

        lbl("Manager.exe path:", 0)
        self.var_path = tk.StringVar(value=slot["path"] if slot else "")
        pf = tk.Frame(self, bg=DARK)
        pf.grid(row=0, column=1, sticky="ew", padx=14, pady=(10, 2))
        tk.Entry(pf, textvariable=self.var_path, width=32,
                 bg=PANEL, fg=WHITE, insertbackground=WHITE,
                 font=("Segoe UI", 9), bd=0,
                 highlightthickness=1, highlightbackground=CARD).pack(side="left")
        tk.Button(pf, text="Browse", command=self._browse,
                  bg=CARD, fg=TEXT, font=("Segoe UI", 8),
                  bd=0, padx=6, pady=2, cursor="hand2").pack(side="left", padx=(6, 0))

        lbl("Start time (HH : MM):", 1)
        sv = slot["start"].split(":") if slot else ["20", "00"]
        self.var_sh = tk.StringVar(value=sv[0])
        self.var_sm = tk.StringVar(value=sv[1])
        tf = tk.Frame(self, bg=DARK)
        tf.grid(row=1, column=1, sticky="w", padx=14, pady=(10, 2))
        tk.Entry(tf, textvariable=self.var_sh, width=4,
                 bg=PANEL, fg=WHITE, insertbackground=WHITE,
                 font=("Segoe UI", 11), bd=0,
                 highlightthickness=1, highlightbackground=CARD).pack(side="left")
        tk.Label(tf, text=":", bg=DARK, fg=WHITE,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=4)
        tk.Entry(tf, textvariable=self.var_sm, width=4,
                 bg=PANEL, fg=WHITE, insertbackground=WHITE,
                 font=("Segoe UI", 11), bd=0,
                 highlightthickness=1, highlightbackground=CARD).pack(side="left")

        lbl("Duration (minutes):", 2)
        self.var_dur = tk.StringVar(value=str(slot["duration"]) if slot else "60")
        tk.Entry(self, textvariable=self.var_dur, width=8,
                 bg=PANEL, fg=WHITE, insertbackground=WHITE,
                 font=("Segoe UI", 11), bd=0,
                 highlightthickness=1, highlightbackground=CARD).grid(
            row=2, column=1, sticky="w", padx=14, pady=(10, 2))

        bf = tk.Frame(self, bg=DARK)
        bf.grid(row=3, column=0, columnspan=2, pady=14)
        tk.Button(bf, text="Save", command=self._save,
                  bg=GREEN, fg=WHITE, font=("Segoe UI", 9, "bold"),
                  bd=0, padx=18, pady=6, cursor="hand2").pack(side="left", padx=8)
        tk.Button(bf, text="Cancel", command=self.destroy,
                  bg=PANEL, fg=MUTED, font=("Segoe UI", 9),
                  bd=0, padx=12, pady=6, cursor="hand2").pack(side="left")

    def _browse(self):
        p = filedialog.askopenfilename(
            title="Select Manager.exe",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if p:
            self.var_path.set(os.path.normpath(p))

    def _save(self):
        path = self.var_path.get().strip()
        if not path:
            messagebox.showerror("Error", "Select a Manager.exe file.", parent=self)
            return
        try:
            sh  = int(self.var_sh.get())
            sm  = int(self.var_sm.get())
            dur = int(self.var_dur.get())
            assert 0 <= sh <= 23 and 0 <= sm <= 59 and dur > 0
        except Exception:
            messagebox.showerror("Error", "Invalid time or duration.", parent=self)
            return
        self.result = {"path": path, "start": f"{sh:02d}:{sm:02d}", "duration": dur}
        self.destroy()

# ── Main App ──────────────────────────────────────────────────────────────────
class RotatorApp(tk.Tk):
    def __init__(self, logged_in_user=""):
        super().__init__()
        self.title(f"phBot Manager Rotator  —  {logged_in_user}  [v{CURRENT_VERSION}]")
        self.geometry("920x720")
        self.minsize(900, 660)
        self.configure(bg=DARK)

        self.primary_path     = tk.StringVar()
        self.slots            = []
        self.running          = False
        self.stop_event       = threading.Event()
        self.primary_proc     = None
        self.slot_proc        = None
        self.active_slot_path = None

        self._load_config()
        self._build_ui()
        self._tick()

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                d = json.load(open(CONFIG_FILE))
                self.primary_path.set(d.get("primary", ""))
                self.slots = d.get("slots", [])
            except Exception:
                pass

    def _save_config(self):
        try:
            d = {}
            if os.path.exists(CONFIG_FILE):
                try:
                    d = json.load(open(CONFIG_FILE))
                except Exception:
                    pass
            d["primary"] = self.primary_path.get()
            d["slots"]   = self.slots
            json.dump(d, open(CONFIG_FILE, "w"), indent=2)
        except Exception:
            pass

    def _build_ui(self):
        hdr = tk.Frame(self, bg=PANEL, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"phBot Manager Rotator  v{CURRENT_VERSION}",
                 bg=PANEL, fg=WHITE,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=20)
        self.lbl_clock = tk.Label(hdr, text="", bg=PANEL, fg=MUTED,
                                  font=("Segoe UI", 11, "bold"))
        self.lbl_clock.pack(side="right", padx=16)
        self.lbl_status = tk.Label(hdr, text="● Idle", bg=PANEL, fg=MUTED,
                                   font=("Segoe UI", 10))
        self.lbl_status.pack(side="right", padx=4)

        body = tk.Frame(self, bg=DARK)
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(2, weight=1)

        pf = tk.LabelFrame(body, text=" Primary Manager (always on) ",
                           bg=DARK, fg=MUTED, font=("Segoe UI", 9),
                           bd=1, relief="solid")
        pf.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
        tk.Label(pf, text="Manager.exe:", bg=DARK, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(8, 2))
        pr = tk.Frame(pf, bg=DARK)
        pr.pack(fill="x", padx=10, pady=(0, 10))
        tk.Entry(pr, textvariable=self.primary_path, width=26,
                 bg=PANEL, fg=WHITE, insertbackground=WHITE,
                 font=("Segoe UI", 9), bd=0,
                 highlightthickness=1, highlightbackground=CARD
                 ).pack(side="left", fill="x", expand=True)
        self._btn(pr, "Browse", self._browse_primary, CARD, TEXT).pack(side="left", padx=(6, 0))

        stf = tk.LabelFrame(body, text=" Status ",
                            bg=DARK, fg=MUTED, font=("Segoe UI", 9),
                            bd=1, relief="solid")
        stf.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        tk.Label(stf, text="Now running:", bg=DARK, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10, 2))
        self.lbl_current = tk.Label(stf, text="—", bg=DARK, fg=WHITE,
                                    font=("Segoe UI", 10, "bold"),
                                    wraplength=200, justify="left")
        self.lbl_current.pack(anchor="w", padx=10)
        tk.Label(stf, text="Next slot:", bg=DARK, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10, 2))
        self.lbl_next = tk.Label(stf, text="—", bg=DARK, fg=AMBER,
                                 font=("Segoe UI", 9))
        self.lbl_next.pack(anchor="w", padx=10)
        tk.Label(stf, text="Slot time remaining:", bg=DARK, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10, 2))
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("g.Horizontal.TProgressbar",
                        troughcolor=PANEL, background=GREEN,
                        darkcolor=GREEN, lightcolor=GREEN, bordercolor=DARK)
        self.progress = ttk.Progressbar(stf, length=200, mode="determinate",
                                        style="g.Horizontal.TProgressbar")
        self.progress.pack(anchor="w", padx=10, pady=(0, 2))
        self.lbl_remaining = tk.Label(stf, text="", bg=DARK, fg=GREEN,
                                      font=("Segoe UI", 10, "bold"))
        self.lbl_remaining.pack(anchor="w", padx=10, pady=(0, 10))

        sf = tk.LabelFrame(body, text=" Scheduled Slots ",
                           bg=DARK, fg=MUTED, font=("Segoe UI", 9),
                           bd=1, relief="solid")
        sf.grid(row=0, column=1, rowspan=2, sticky="nsew", pady=(0, 8))
        sf.rowconfigure(1, weight=1)
        sf.columnconfigure(0, weight=1)
        hr = tk.Frame(sf, bg=CARD)
        hr.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        for txt, w in [("#", 3), ("Manager folder", 18), ("Start", 7), ("Dur(min)", 9), ("Ends", 7)]:
            tk.Label(hr, text=txt, bg=CARD, fg=MUTED,
                     font=("Segoe UI", 8, "bold"), width=w,
                     anchor="w").pack(side="left", padx=4)
        lf2 = tk.Frame(sf, bg=DARK)
        lf2.grid(row=1, column=0, sticky="nsew", padx=8)
        lf2.rowconfigure(0, weight=1)
        lf2.columnconfigure(0, weight=1)
        self.slot_lb = tk.Listbox(lf2, bg=PANEL, fg=TEXT,
                                  selectbackground=CARD,
                                  font=("Consolas", 9),
                                  bd=0, highlightthickness=0,
                                  activestyle="none")
        self.slot_lb.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(lf2, orient="vertical",
                          command=self.slot_lb.yview, bg=DARK)
        sb.grid(row=0, column=1, sticky="ns")
        self.slot_lb.configure(yscrollcommand=sb.set)
        sbr = tk.Frame(sf, bg=DARK)
        sbr.grid(row=2, column=0, sticky="ew", padx=8, pady=8)
        self._btn(sbr, "+ Add",  self._add_slot,    CARD,  TEXT).pack(side="left", padx=(0, 6))
        self._btn(sbr, "Edit",   self._edit_slot,   PANEL, AMBER).pack(side="left", padx=(0, 6))
        self._btn(sbr, "Remove", self._remove_slot, PANEL, ACCENT).pack(side="left")
        self._refresh_slots()

        logf = tk.LabelFrame(body, text=" Log ",
                             bg=DARK, fg=MUTED, font=("Segoe UI", 9),
                             bd=1, relief="solid")
        logf.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 4))
        logf.rowconfigure(0, weight=1)
        logf.columnconfigure(0, weight=1)
        self.log_text = tk.Text(logf, bg=PANEL, fg=TEXT,
                                font=("Consolas", 8),
                                bd=0, highlightthickness=0,
                                state="disabled", wrap="word", height=8)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        lsb = tk.Scrollbar(logf, orient="vertical",
                           command=self.log_text.yview, bg=DARK)
        lsb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=lsb.set)
        for tag, col in [("ok", GREEN), ("warn", AMBER), ("err", ACCENT), ("info", BLUE)]:
            self.log_text.tag_config(tag, foreground=col)

        bot = tk.Frame(self, bg=DARK)
        bot.pack(fill="x", padx=14, pady=(0, 12))
        self.btn_start = self._btn(bot, "▶  Start", self._start, GREEN, WHITE, (18, 8))
        self.btn_start.pack(side="left", padx=(0, 10))
        self.btn_stop = self._btn(bot, "■  Stop", self._stop, ACCENT, WHITE, (18, 8))
        self.btn_stop.pack(side="left")
        self.btn_stop.configure(state="disabled")
        self._btn(bot, "Clear Log", self._clear_log, PANEL, MUTED).pack(side="right")

    def _btn(self, parent, text, cmd, bg, fg, pad=(10, 6)):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg=fg, activebackground=bg,
                         activeforeground=WHITE,
                         font=("Segoe UI", 9, "bold"),
                         bd=0, relief="flat", cursor="hand2",
                         padx=pad[0], pady=pad[1])

    def _tick(self):
        self.lbl_clock.configure(text=datetime.now().strftime("%H:%M:%S"))
        self._refresh_next_label()
        self.after(1000, self._tick)

    def _refresh_next_label(self):
        if not self.slots:
            self.lbl_next.configure(text="No slots configured")
            return
        ns = now_sec()
        best_diff = None
        best_slot = None
        for s in self.slots:
            sh, sm   = map(int, s["start"].split(":"))
            start_s  = sh * 3600 + sm * 60
            diff     = start_s - ns
            if diff <= 0:
                diff += 86400
            if best_diff is None or diff < best_diff:
                best_diff = diff
                best_slot = s
        if best_slot:
            h, m, sec = sec_to_hms(best_diff)
            folder = os.path.basename(os.path.dirname(best_slot["path"]))
            self.lbl_next.configure(
                text=f"{best_slot['start']}  {folder}  (in {h}h {m:02d}m {sec:02d}s)")

    def _log(self, msg, tag=""):
        def _do():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{now_str()}] {msg}\n", tag)
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _browse_primary(self):
        p = filedialog.askopenfilename(
            title="Select Primary Manager.exe",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if p:
            self.primary_path.set(os.path.normpath(p))
            self._save_config()

    def _refresh_slots(self):
        self.slot_lb.delete(0, "end")
        for i, s in enumerate(self.slots):
            sh, sm = map(int, s["start"].split(":"))
            end_s  = sh * 3600 + sm * 60 + s["duration"] * 60
            eh     = (end_s % 86400) // 3600
            em     = ((end_s % 86400) % 3600) // 60
            folder = os.path.basename(os.path.dirname(s["path"]))
            self.slot_lb.insert(
                "end",
                f"  {i+1}   {folder:<20}  {s['start']}   {s['duration']:<8}  {eh:02d}:{em:02d}")

    def _add_slot(self):
        dlg = SlotDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self.slots.append(dlg.result)
            self._refresh_slots()
            self._save_config()

    def _edit_slot(self):
        sel = self.slot_lb.curselection()
        if not sel:
            return
        dlg = SlotDialog(self, self.slots[sel[0]])
        self.wait_window(dlg)
        if dlg.result:
            self.slots[sel[0]] = dlg.result
            self._refresh_slots()
            self._save_config()

    def _remove_slot(self):
        sel = self.slot_lb.curselection()
        if sel:
            self.slots.pop(sel[0])
            self._refresh_slots()
            self._save_config()

    def _kill_proc(self, proc, label):
        if proc is None or proc.poll() is not None:
            return
        self._log(f"Killing: {label}  PID {proc.pid}", "warn")
        force_kill(proc.pid)

    def _launch(self, path):
        name = os.path.basename(path)
        if not os.path.exists(path):
            self._log(f"File not found: {path}", "err")
            return None
        try:
            proc = subprocess.Popen([path], cwd=os.path.dirname(path))
            self._log(f"[+] Started: {name}  PID {proc.pid}", "ok")
            return proc
        except Exception as e:
            self._log(f"Failed to start {name}: {e}", "err")
            return None

    def _scheduler(self):
        primary = self.primary_path.get().strip()
        if not primary:
            self._log("No primary Manager set!", "err")
            self._reset_ui()
            return

        self._log("=== Scheduler started ===", "ok")
        self._log(f"Primary: {os.path.basename(primary)}", "info")
        self.primary_proc = self._launch(primary)
        pname = os.path.basename(primary)
        self.after(0, lambda n=pname: self.lbl_current.configure(text=f"{n}  [primary]"))

        while not self.stop_event.is_set():
            ns = now_sec()
            active       = None
            active_start = 0
            active_end   = 0
            for s in self.slots:
                sh, sm  = map(int, s["start"].split(":"))
                s_start = sh * 3600 + sm * 60
                s_end   = s_start + s["duration"] * 60
                if s_end > 86400:
                    if ns >= s_start or ns < s_end % 86400:
                        active       = s
                        active_start = s_start if ns >= s_start else s_start - 86400
                        active_end   = active_start + s["duration"] * 60
                        break
                else:
                    if s_start <= ns < s_end:
                        active       = s
                        active_start = s_start
                        active_end   = s_end
                        break

            if active:
                if self.primary_proc and self.primary_proc.poll() is None:
                    self._kill_proc(self.primary_proc, os.path.basename(primary))
                    self.primary_proc = None
                if (self.slot_proc and self.slot_proc.poll() is None and
                        self.active_slot_path != active["path"]):
                    self._kill_proc(self.slot_proc, "previous slot")
                    self.slot_proc        = None
                    self.active_slot_path = None
                if not self.slot_proc or self.slot_proc.poll() is not None:
                    self.slot_proc        = self._launch(active["path"])
                    self.active_slot_path = active["path"]

                slot_folder    = os.path.basename(os.path.dirname(active["path"]))
                slot_start_str = active["start"]
                sh2, sm2 = map(int, active["start"].split(":"))
                end_s2   = sh2 * 3600 + sm2 * 60 + active["duration"] * 60
                eh2, em2 = (end_s2 % 86400) // 3600, ((end_s2 % 86400) % 3600) // 60
                label_txt = f"[Slot] {slot_folder}  ({slot_start_str} - {eh2:02d}:{em2:02d})"
                self.after(0, lambda t=label_txt: self.lbl_current.configure(text=t))

                total  = active_end - active_start
                done   = ns - active_start
                remain = max(0, active_end - ns)
                pct    = min(100.0, done / total * 100) if total > 0 else 0
                rh, rm, rs = sec_to_hms(remain)
                self.after(0, lambda p=pct: self.progress.configure(value=p, maximum=100))
                self.after(0, lambda h=rh, m=rm, s=rs:
                           self.lbl_remaining.configure(text=f"{h:02d}:{m:02d}:{s:02d}"))
            else:
                if self.slot_proc and self.slot_proc.poll() is None:
                    self._log("Slot ended — resuming primary.", "info")
                    self._kill_proc(self.slot_proc, "slot manager")
                    self.slot_proc        = None
                    self.active_slot_path = None
                if not self.primary_proc or self.primary_proc.poll() is not None:
                    self.primary_proc = self._launch(primary)
                pname2 = os.path.basename(primary)
                self.after(0, lambda n=pname2: self.lbl_current.configure(text=f"{n}  [primary]"))
                self.after(0, lambda: self.progress.configure(value=0))
                self.after(0, lambda: self.lbl_remaining.configure(text=""))

            time.sleep(1)

        self._kill_proc(self.slot_proc,    "slot manager")
        self._kill_proc(self.primary_proc, "primary")
        self._log("=== Scheduler stopped ===", "warn")
        self._reset_ui()

    def _start(self):
        if not self.primary_path.get().strip():
            messagebox.showwarning("Missing", "Set the Primary Manager.exe first.")
            return
        self._save_config()
        self.stop_event.clear()
        self.running = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.lbl_status.configure(text="● Running", fg=GREEN)
        threading.Thread(target=self._scheduler, daemon=True).start()

    def _stop(self):
        self.stop_event.set()
        self.btn_stop.configure(state="disabled")
        self._log("Stop requested...", "warn")

    def _reset_ui(self):
        self.running      = False
        self.primary_proc = None
        self.slot_proc    = None
        self.after(0, lambda: self.btn_start.configure(state="normal"))
        self.after(0, lambda: self.btn_stop.configure(state="disabled"))
        self.after(0, lambda: self.lbl_status.configure(text="● Idle", fg=MUTED))
        self.after(0, lambda: self.lbl_current.configure(text="—"))
        self.after(0, lambda: self.progress.configure(value=0))
        self.after(0, lambda: self.lbl_remaining.configure(text=""))


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not ctypes.windll.shell32.IsUserAnAdmin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

    # Check for update before showing login
    update_available, latest_ver = check_for_update()
    if update_available:
        show_update_prompt(latest_ver)

    login = LoginWindow()
    login.mainloop()
    if not login.result:
        sys.exit()

    app = RotatorApp(logged_in_user=login.username)
    app.mainloop()
