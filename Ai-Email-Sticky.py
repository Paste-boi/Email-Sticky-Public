
# Ai-Email-Sticky.py — UI build with:
# - VS Code–matched chrome colors
# - Larger checkbox
# - ❌ delete button
# - Header buttons: Poll Now, Clear Completed
# - NEW: When a task is completed, ALL text in the note turns subtle green (#6A9955)
# Logic (polling, AI, dedup, cutoff, leave-unread, etc.) unchanged.

import imaplib, email, re, time, threading, sqlite3, json, os, configparser, logging, logging.handlers
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime, parseaddr
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import font as tkfont

LAST_AI_ERROR = ""

try:
    import openai as _openai_pkg
    from openai import OpenAI
except Exception:
    _openai_pkg = None
    OpenAI = None

APP_DIR = Path.home() / ".ai_email_sticky"
APP_DIR.mkdir(exist_ok=True)
DB_PATH = APP_DIR / "tasks.db"
LOG_PATH = APP_DIR / "ai_sticky.log"
CONFIG_PATH = Path("config.ini")

# ----------------- Logging -----------------
def setup_logging():
    logger = logging.getLogger("ai_sticky")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fh = logging.handlers.RotatingFileHandler(LOG_PATH, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(fh); logger.addHandler(ch)
    logger.info("==== AI Email Sticky starting up ====")
    logger.info("Log file: %s", LOG_PATH)
    return logger

log = setup_logging()

# ----------------- Persistence -----------------
def ensure_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email_uid TEXT, subject TEXT, snippet TEXT,
        task_text TEXT NOT NULL,
        created_at TEXT NOT NULL,
        completed_at TEXT, is_completed INTEGER DEFAULT 0,
        archived_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS processed_uids (email_uid TEXT PRIMARY KEY)""")
    cur.execute("""CREATE INDEX IF NOT EXISTS idx_tasks_email_uid ON tasks(email_uid)""")
    conn.commit(); conn.close()

def save_metadata(key, value):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("INSERT INTO metadata(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))
    conn.commit(); conn.close()

def get_metadata(key, default=None):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("SELECT value FROM metadata WHERE key=?", (key,))
    row = cur.fetchone(); conn.close()
    return row[0] if row else default

def is_uid_processed(uid):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("SELECT 1 FROM processed_uids WHERE email_uid=?", (uid,))
    ok = cur.fetchone() is not None
    conn.close(); return ok

def mark_uid_processed(uid):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO processed_uids(email_uid) VALUES(?)", (uid,))
    conn.commit(); conn.close()

def add_task(task_text, subject="", snippet="", email_uid=None):
    if email_uid and is_uid_processed(email_uid):
        log.info("Skip duplicate UID=%s (already processed)", email_uid)
        return False
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("INSERT INTO tasks (email_uid, subject, snippet, task_text, created_at) VALUES (?,?,?,?,?)",
                (email_uid, subject, snippet, task_text, datetime.utcnow().isoformat()))
    conn.commit(); conn.close()
    if email_uid:
        mark_uid_processed(email_uid)
    log.info("Added task (UID=%s): %s", email_uid, task_text[:160])
    return True

def list_active_tasks(retention_hours=12, return_counts=False):
    cutoff = datetime.utcnow() - timedelta(hours=retention_hours)
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("""SELECT id, task_text, is_completed, completed_at, subject
                   FROM tasks WHERE archived_at IS NULL
                   ORDER BY is_completed, id DESC""")
    rows = cur.fetchall()
    # Archive completed items older than retention
    to_archive = []
    for tid, _, done, comp_at, _ in rows:
        if done and comp_at:
            try:
                if datetime.fromisoformat(comp_at) < cutoff:
                    to_archive.append(tid)
            except Exception:
                pass
    if to_archive:
        now = datetime.utcnow().isoformat()
        q = ",".join("?" for _ in to_archive)
        cur.execute(f"UPDATE tasks SET archived_at=? WHERE id IN ({q})", [now, *to_archive])
        conn.commit()
        cur.execute("""SELECT id, task_text, is_completed, completed_at, subject
                       FROM tasks WHERE archived_at IS NULL
                       ORDER BY is_completed, id DESC""")
        rows = cur.fetchall()

    cur.execute("SELECT COUNT(*) FROM tasks WHERE archived_at IS NULL AND is_completed=0")
    active_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM tasks WHERE archived_at IS NULL AND is_completed=1")
    completed_count = cur.fetchone()[0]
    conn.close()
    if return_counts:
        return rows, active_count, completed_count
    return rows

def mark_task_completed(task_id, done=True):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    if done:
        cur.execute("UPDATE tasks SET is_completed=1, completed_at=? WHERE id=?", (datetime.utcnow().isoformat(), task_id))
    else:
        cur.execute("UPDATE tasks SET is_completed=0, completed_at=NULL WHERE id=?", (task_id,))
    conn.commit(); conn.close()

def delete_task(task_id):
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    cur.execute("DELETE FROM tasks WHERE id=?", (task_id,)); conn.commit(); conn.close()

def archive_all_completed_now():
    conn = sqlite3.connect(DB_PATH); cur = conn.cursor()
    now = datetime.utcnow().isoformat()
    cur.execute("UPDATE tasks SET archived_at=? WHERE archived_at IS NULL AND is_completed=1", (now,))
    count = cur.rowcount
    conn.commit(); conn.close()
    return count

# ----------------- Config -----------------
def load_config():
    cfg = configparser.ConfigParser()
    if not CONFIG_PATH.exists():
        cfg["imap"] = {
            "server":"imap.gmail.com",
            "username":"",
            "password":"",
            "folder":"INBOX",
            "ssl":"true",
            "cutoff_date":"",
            "mark_as_read":"false"
        }
        cfg["ai"] = {
            "enabled":"true",
            "model":"gpt-5-mini",
            "temperature":"1",
            "api_key":"",
            "base_url":"",
            "classify_before_add":"true",
            "drop_labels":"marketing, fyi"
        }
        cfg["app"] = {"retention_hours":"12","poll_seconds":"300","ui_refresh_seconds":"30"}
        cfg["ui"]  = {"font_size":"10","always_on_top":"true","theme":"light","colorful_text":"true"}
        with open(CONFIG_PATH,"w") as f: cfg.write(f)
    cfg.read(CONFIG_PATH)
    return cfg

def get_bool(cfg, section, option, fallback=False):
    try:
        raw = cfg.get(section, option, fallback=str(fallback))
    except Exception:
        return fallback
    raw = str(raw).split(";",1)[0].split("#",1)[0].strip().lower()
    return raw in ("1","true","t","yes","y","on")

# ----------------- Summaries + AI -----------------
def heuristic_summary(text, max_len=140):
    for line in text.splitlines():
        s = line.strip()
        if s and not s.startswith(">"):
            s = re.sub(r"\s+", " ", s)
            return (s[:max_len]).rstrip()
    return (text.strip().replace("\n"," ")[:max_len]).rstrip()

def llm_summary(text, subject, model="gpt-5-mini", api_key=None, base_url=None, temperature=1, max_len=140):
    global LAST_AI_ERROR
    if OpenAI is None:
        LAST_AI_ERROR = "OpenAI SDK not importable; install with: pip install openai"
        return heuristic_summary(text, max_len), "OFF"
    api_key = api_key or os.getenv("OPENAI_API_KEY","")
    if not api_key:
        LAST_AI_ERROR = "OPENAI_API_KEY not set (env var or [ai] api_key)."
        return heuristic_summary(text, max_len), "OFF"
    try:
        kwargs = {"api_key": api_key}
        if base_url:
            kwargs["base_url"] = base_url
        client = OpenAI(**kwargs)
        prompt = (
            "Summarize this email into a single concise action-oriented sentence "
            f"(<= {max_len} characters). No preamble, no quotes — just the sentence.\n\n"
            f"Subject: {subject}\n\n{text[:6000]}"
        )
        r = client.chat.completions.create(
            model=model, temperature=float(temperature),
            messages=[{"role":"user","content":prompt}]
        )
        out = r.choices[0].message.content.strip()
        out = re.sub(r"^\W+|\W+$", "", out)
        return (out[:max_len].strip() or heuristic_summary(text, max_len)), "ON"
    except Exception as e:
        LAST_AI_ERROR = f"{type(e).__name__}: {e}"
        return heuristic_summary(text, max_len), "FALLBACK"

def ai_classify_label(text, subject, model="gpt-5-mini", api_key=None, base_url=None):
    global LAST_AI_ERROR
    if OpenAI is None:
        return "actionable", "OFF"
    api_key = api_key or os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        return "actionable", "OFF"
    try:
        kwargs = {"api_key": api_key}
        if base_url:
            kwargs["base_url"] = base_url
        client = OpenAI(**kwargs)
        prompt = (
            "Classify this email for triage with ONE WORD only:\n"
            "actionable = asks me to do something or likely needs a response\n"
            "fyi        = informational only, no action needed\n"
            "marketing  = promo/sales/newsletter/offer\n\n"
            f"Subject: {subject}\n\n{text[:3000]}\n\n"
            "Answer with exactly one label: actionable or fyi or marketing."
        )
        resp = client.chat.completions.create(
            model=model, temperature=1,
            messages=[{"role":"user","content":prompt}]
        )
        label = (resp.choices[0].message.content or "").strip().lower()
        if "market" in label:
            label = "marketing"
        elif "fyi" in label:
            label = "fyi"
        else:
            label = "actionable"
        return label, "ON"
    except Exception as e:
        LAST_AI_ERROR = f"{type(e).__name__}: {e}"
        return "actionable", "FALLBACK"

# ----------------- Themes (with sampled border tones) -----------------
LIGHT = {
    "root_bg": "#e9e9e9",  # sampled light chrome
    "bg": "#ffffff",
    "fg": "#1f1f1f",
    "summary_fg": "#000000",
    "from_fg": "#007acc",
    "received_fg": "#f14c4c",
    "btn_bg": "#e9e9e9",
    "btn_fg": "#1f1f1f",
    "btn_active": "#dcdcdc"
}
DARK = {
    "root_bg": "#2d2d2d",  # sampled dark chrome
    "bg": "#1e1e1e",
    "fg": "#d4d4d4",
    "summary_fg": "#ffffff",
    "from_fg": "#569cd6",
    "received_fg": "#f44747",
    "btn_bg": "#2d2d2d",
    "btn_fg": "#d4d4d4",
    "btn_active": "#3a3a3a"
}

COMPLETE_GREEN = "#6A9955"  # subtle VS Code success green

def get_theme(cfg):
    t = cfg.get("ui", "theme", fallback="light").strip().lower()
    return DARK if t == "dark" else LIGHT

# ----------------- Poller (logic unchanged) -----------------
class GmailPoller(threading.Thread):
    def __init__(self, cfg, status_callback=None):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.poll_seconds = int(cfg.get("app","poll_seconds", fallback="300"))
        self.status_callback = status_callback

    def run(self):
        log.info("Poller thread started (every %ss)", self.poll_seconds)
        while True:
            try:
                self.check_mail()
            except Exception as e:
                log.exception("Poll error: %s", e)
            time.sleep(self.poll_seconds)

    def check_mail(self):
        host   = self.cfg.get("imap","server",  fallback="imap.gmail.com")
        user   = self.cfg.get("imap","username",fallback="")
        pw     = self.cfg.get("imap","password",fallback="")
        folder = self.cfg.get("imap","folder",  fallback="INBOX")
        use_ssl= self.cfg.getboolean("imap","ssl",fallback=True)
        cutoff = self.cfg.get("imap","cutoff_date", fallback="").strip()
        mark_as_read = get_bool(self.cfg, "imap", "mark_as_read", fallback=False)

        if not (user and pw):
            if self.status_callback: self.status_callback("IMAP OFF")
            log.warning("IMAP creds not set; skipping poll")
            return

        M = imaplib.IMAP4_SSL(host) if use_ssl else imaplib.IMAP4(host)
        M.login(user, pw)
        M.select(folder)

        last_uid = get_metadata("last_uid", "0")
        base_q = f"(UID {int(last_uid)+1}:*)" if last_uid else "ALL"
        query  = f'(SINCE "{cutoff}") {base_q}' if cutoff else base_q

        typ, data = M.uid("search", None, query)
        if typ != "OK":
            M.logout()
            if self.status_callback: self.status_callback("IMAP ERR")
            return

        uids = [u.decode() for u in data[0].split() if u]
        if not uids:
            M.logout()
            if self.status_callback: self.status_callback("IMAP OK (no new)")
            return

        uids = sorted(uids, key=lambda s: int(s))
        model = self.cfg.get("ai","model", fallback="gpt-5-mini")
        ai_key = self.cfg.get("ai","api_key", fallback="").strip() or None
        base_url = self.cfg.get("ai","base_url", fallback="").strip() or None
        temperature = self.cfg.get("ai","temperature", fallback="1")
        ai_enabled = self.cfg.getboolean("ai","enabled", fallback=True)

        max_seen_uid = int(last_uid or 0)
        last_ai_mode = None
        new_count = 0
        dropped = 0

        for uid_s in uids:
            if is_uid_processed(uid_s):
                max_seen_uid = max(max_seen_uid, int(uid_s))
                continue

            typ, msg_data = M.uid("fetch", uid_s.encode(), "(BODY.PEEK[])")
            if typ != "OK" or not msg_data or not msg_data[0]:
                continue

            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            subject = msg.get("Subject","")
            from_hdr = msg.get("From","")
            sender_email = parseaddr(from_hdr)[1] or "(unknown)"
            date_hdr = msg.get("Date","")
            try:
                dt_utc = parsedate_to_datetime(date_hdr)
                if dt_utc.tzinfo is None:
                    dt_utc = dt_utc.replace(tzinfo=timezone.utc)
                dt_local = dt_utc.astimezone()
                received_str = dt_local.strftime("%Y-%m-%d %H:%M")
            except Exception:
                received_str = "(unknown date)"

            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    disp = str(part.get("Content-Disposition",""))
                    if ctype == "text/plain" and "attachment" not in disp.lower():
                        payload = part.get_payload(decode=True) or b""
                        try:
                            body += payload.decode(part.get_content_charset() or "utf-8", errors="ignore") + "\n"
                        except Exception:
                            body += payload.decode("utf-8", errors="ignore") + "\n"
            else:
                payload = msg.get_payload(decode=True) or b""
                try:
                    body += payload.decode(msg.get_content_charset() or "utf-8", errors="ignore")
                except Exception:
                    body += payload.decode("utf-8", errors="ignore")

            classify_enabled = get_bool(self.cfg, "ai", "classify_before_add", True)
            drop_labels_csv  = self.cfg.get("ai", "drop_labels", fallback="marketing, fyi")
            drop_set = {x.strip().lower() for x in drop_labels_csv.split(",") if x.strip()}

            if classify_enabled:
                cls_label, cls_mode = ai_classify_label(body, subject, model=model, api_key=ai_key, base_url=base_url)
                if cls_label in drop_set:
                    dropped += 1
                    mark_uid_processed(uid_s)
                    max_seen_uid = max(max_seen_uid, int(uid_s))
                    continue

            snippet = (body.strip().splitlines() or [""])[0][:140]

            if ai_enabled:
                summary, ai_mode = llm_summary(body, subject, model=model, api_key=ai_key, base_url=base_url, temperature=temperature, max_len=140)
            else:
                summary, ai_mode = heuristic_summary(body, 140), "OFF"
            last_ai_mode = ai_mode

            line = f"From: {sender_email} | Received: {received_str} | Summary: {summary}"
            created = add_task(line, subject=subject, snippet=snippet, email_uid=uid_s)
            if created and mark_as_read:
                try:
                    M.uid("store", uid_s.encode(), "+FLAGS", r"(\Seen)")
                except Exception:
                    pass

            max_seen_uid = max(max_seen_uid, int(uid_s))

        save_metadata("last_uid", str(max_seen_uid))
        M.logout()
        if self.status_callback and last_ai_mode:
            self.status_callback(f"AI {last_ai_mode} | +{new_count} / dropped {dropped}")

# ----------------- UI -----------------
class StickyUI(tk.Tk):
    def __init__(self, cfg):
        super().__init__()
        self.cfg = cfg
        self.title("AI Email Sticky Note")

        # UI prefs
        self.font_size = int(cfg.get("ui","font_size", fallback="10"))
        self.always_on_top = get_bool(cfg, "ui", "always_on_top", True)
        self.theme = get_theme(cfg)
        self.colorful = get_bool(cfg, "ui", "colorful_text", True)

        # Root chrome color
        self.configure(bg=self.theme["root_bg"])
        self.attributes("-topmost", self.always_on_top)

        # Fonts
        self.base_font = tkfont.nametofont("TkDefaultFont")
        try:
            self.base_font.configure(size=self.font_size)
        except Exception:
            pass
        self.checkbox_font = ("Segoe UI", max(self.font_size+2, 12))  # bigger checkbox

        self.retention_hours = int(cfg.get("app","retention_hours", fallback="12"))
        self.ui_refresh_seconds = int(cfg.get("app","ui_refresh_seconds", fallback="30"))

        # --- Shell frame with border look ---
        self.shell = tk.Frame(self, bg=self.theme["root_bg"])
        self.shell.pack(fill="both", expand=True)

        # Content frame (inner editor-like bg)
        self.content = tk.Frame(self.shell, bg=self.theme["bg"], highlightthickness=1, highlightbackground=self.theme["root_bg"])
        self.content.pack(fill="both", expand=True, padx=6, pady=6)

        # Header
        header = tk.Frame(self.content, bg=self.theme["bg"])
        header.pack(fill="x", padx=6, pady=(6,2))

        tk.Label(header, text="To-Dos (from Gmail)", bg=self.theme["bg"], fg=self.theme["fg"],
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        # Right side controls
        right = tk.Frame(header, bg=self.theme["bg"])
        right.pack(side="right")

        self.btn_poll = tk.Button(right, text="Poll Now", command=self.poll_now,
                                  bg=self.theme["btn_bg"], fg=self.theme["btn_fg"],
                                  activebackground=self.theme["btn_active"], relief="flat", padx=8, pady=2)
        self.btn_poll.pack(side="left", padx=(0,6))

        self.btn_clear = tk.Button(right, text="Clear Completed", command=self.clear_completed_now,
                                   bg=self.theme["btn_bg"], fg=self.theme["btn_fg"],
                                   activebackground=self.theme["btn_active"], relief="flat", padx=8, pady=2)
        self.btn_clear.pack(side="left", padx=(0,10))

        self.status_var = tk.StringVar(value="AI ?")
        tk.Label(right, textvariable=self.status_var, bg=self.theme["bg"], fg=self.theme["fg"]).pack(side="left")

        # Menus
        menubar = tk.Menu(self)
        viewmenu = tk.Menu(menubar, tearoff=0)
        self.dark_mode_var = tk.BooleanVar(value=(self.cfg.get("ui","theme",fallback="light").lower()=="dark"))
        self.colorful_var = tk.BooleanVar(value=self.colorful)
        viewmenu.add_checkbutton(label="Dark mode", variable=self.dark_mode_var, command=self.toggle_dark_mode)
        viewmenu.add_checkbutton(label="Colorful text", variable=self.colorful_var, command=self.toggle_colorful)
        menubar.add_cascade(label="View", menu=viewmenu)

        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Test OpenAI summarizer…", command=self.self_test_ai)
        helpmenu.add_command(label="Show last AI error…", command=self.show_last_ai_error)
        helpmenu.add_separator()
        helpmenu.add_command(label="View latest log…", command=self.view_latest_log)
        helpmenu.add_command(label="Open log folder", command=self.open_log_folder)
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.config(menu=menubar)

        # List area
        geom = get_metadata("window_geometry", "500x640+80+80")
        try:
            self.geometry(geom)
        except Exception:
            self.geometry("500x640+80+80")

        self.canvas = tk.Canvas(self.content, bg=self.theme["bg"], highlightthickness=0)
        self.scroll = ttk.Scrollbar(self.content, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.scroll.pack(side="right", fill="y", padx=(0,6), pady=(0,6))
        self.canvas.pack(side="left", fill="both", expand=True, padx=(6,0), pady=(0,6))

        self.list_frame = tk.Frame(self.canvas, bg=self.theme["bg"])
        self.canvas.create_window((0,0), window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Status bar
        self.statusbar = tk.Label(self.content, text="", bg=self.theme["bg"], fg=self.theme["fg"], anchor="w")
        self.statusbar.pack(fill="x", padx=6, pady=(0,8))

        self.refresh_ui()
        self.after(self.ui_refresh_seconds * 1000, self.periodic_refresh)

        # Poller
        self.poller = GmailPoller(cfg, status_callback=self.set_status)
        self.poller.start()

        # Hotkeys
        self.bind_all("<Control-r>", lambda e: self.poll_now())
        self.bind_all("<Control-R>", lambda e: self.poll_now())

        # Persist on close
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def clear_completed_now(self):
        n = archive_all_completed_now()
        messagebox.showinfo("Archive Completed", f"Archived {n} completed tasks.")
        self.refresh_ui()

    def toggle_dark_mode(self):
        self.cfg.set("ui", "theme", "dark" if self.dark_mode_var.get() else "light")
        with open(CONFIG_PATH, "w") as f: self.cfg.write(f)
        self.theme = get_theme(self.cfg)
        self.configure(bg=self.theme["root_bg"])
        self.content.configure(bg=self.theme["bg"], highlightbackground=self.theme["root_bg"])
        for btn in (self.btn_poll, self.btn_clear):
            btn.configure(bg=self.theme["btn_bg"], fg=self.theme["btn_fg"], activebackground=self.theme["btn_active"])
        self.refresh_ui()

    def toggle_colorful(self):
        self.colorful = self.colorful_var.get()
        self.cfg.set("ui", "colorful_text", "true" if self.colorful else "false")
        with open(CONFIG_PATH, "w") as f: self.cfg.write(f)
        self.refresh_ui()

    def poll_now(self):
        threading.Thread(target=self._poll_once, daemon=True).start()

    def _poll_once(self):
        try:
            self.set_status("Manual poll…")
            GmailPoller(self.cfg, status_callback=self.set_status).check_mail()
        except Exception as e:
            log.exception("Manual poll error: %s", e)

    def set_status(self, text):
        self.status_var.set(text)

    def periodic_refresh(self):
        self.refresh_ui()
        self.after(self.ui_refresh_seconds * 1000, self.periodic_refresh)

    def refresh_ui(self):
        for c in self.list_frame.winfo_children():
            c.destroy()

        self.canvas.configure(bg=self.theme["bg"])
        self.list_frame.configure(bg=self.theme["bg"])

        rows, active_count, completed_count = list_active_tasks(self.retention_hours, return_counts=True)

        if not rows:
            tk.Label(self.list_frame, text="No tasks (yet!)",
                     bg=self.theme["bg"], fg=self.theme["fg"]).pack(anchor="w", padx=6, pady=6)
        else:
            for tid, text, done, _, subject in rows:
                # Parse standardized line
                from_match = re.search(r"From:\s*(.*?)\s*\|", text)
                recv_match = re.search(r"Received:\s*([^|]+)", text)
                summ_match = re.search(r"Summary:\s*(.*)$", text)
                from_val = from_match.group(1).strip() if from_match else ""
                recv_val = recv_match.group(1).strip() if recv_match else ""
                summ_val = summ_match.group(1).strip() if summ_match else text

                row = tk.Frame(self.list_frame, bg=self.theme["bg"])
                row.pack(fill="x", pady=4, padx=6)

                var = tk.BooleanVar(value=bool(done))
                cb = tk.Checkbutton(row, variable=var, bg=self.theme["bg"],
                                    font=("Segoe UI", max(self.font_size+2, 12)),
                                    command=lambda t=tid, v=var: (mark_task_completed(t, v.get()), self.refresh_ui()))
                cb.pack(side="left", padx=(0,6))

                block = tk.Frame(row, bg=self.theme["bg"])
                block.pack(side="left", fill="x", expand=True)

                # Line 1
                line1 = tk.Frame(block, bg=self.theme["bg"])
                line1.pack(anchor="w", fill="x")
                from_fg = self.theme["from_fg"] if self.colorful else self.theme["fg"]
                recv_fg = self.theme["received_fg"] if self.colorful else self.theme["fg"]
                from_lbl = tk.Label(line1, text=f"From: {from_val}", bg=self.theme["bg"],
                                    fg=from_fg, font=("Segoe UI", 9, "bold"))
                from_lbl.pack(side="left")
                tk.Label(line1, text="   ", bg=self.theme["bg"], fg=self.theme["fg"]).pack(side="left")
                recv_lbl = tk.Label(line1, text=f"Received: {recv_val}", bg=self.theme["bg"],
                                    fg=recv_fg, font=("Segoe UI", 9))
                recv_lbl.pack(side="left")

                # Line 2
                line2 = tk.Frame(block, bg=self.theme["bg"])
                line2.pack(anchor="w", fill="x")
                sum_fg = self.theme["summary_fg"]
                sum_label = tk.Label(line2, text=f"Summary: {summ_val}",
                                     bg=self.theme["bg"], fg=sum_fg, wraplength=440, justify="left")
                sum_label.pack(side="left", fill="x", expand=True)

                # NEW: If completed, paint entire note green for quick scan
                if done:
                    green = COMPLETE_GREEN
                    from_lbl.configure(fg=green)
                    recv_lbl.configure(fg=green)
                    sum_label.configure(fg=green)

                tk.Button(row, text="❌", bg=self.theme["bg"], relief="flat",
                          command=lambda t=tid: (delete_task(t), self.refresh_ui())).pack(side="right")

        self.statusbar.config(text=f"Active: {active_count}  |  Completed: {completed_count}")

    # Help items (unchanged)
    def self_test_ai(self):
        sample_subject = "Order status and scheduling"
        sample_body = (
            "Hi Peyton,\n"
            "Please confirm the PO 4421 delivery ETA and update the CNC grind schedule for job 77105. "
            "Also send the revised drawing to the customer.\n\nThanks!"
        )
        model = self.cfg.get("ai","model", fallback="gpt-5-mini")
        ai_key = self.cfg.get("ai","api_key", fallback="").strip() or None
        base_url = self.cfg.get("ai","base_url", fallback="").strip() or None
        temperature = self.cfg.get("ai","temperature", fallback="1")
        enabled = self.cfg.getboolean("ai","enabled", fallback=True)

        if not enabled:
            messagebox.showinfo("AI Self-Test", "AI is disabled in config.ini ([ai] enabled=false)."); return

        summary, mode = llm_summary(sample_body, sample_subject, model=model, api_key=ai_key, base_url=base_url, temperature=temperature, max_len=140)
        sdk_ver = getattr(_openai_pkg, "__version__", "(unknown)")
        msg = f"Mode: {mode}\nModel: {model}\nOpenAI SDK: {sdk_ver}\nBase URL: {base_url or '(default)'}\nTemperature: {temperature}\n\nSummary:\n{summary}"
        messagebox.showinfo("AI Self-Test Result", msg)
        self.set_status(f"AI {mode}")

    def show_last_ai_error(self):
        global LAST_AI_ERROR
        if not LAST_AI_ERROR:
            messagebox.showinfo("Last AI Error", "No AI errors recorded in this session.")
        else:
            messagebox.showerror("Last AI Error", LAST_AI_ERROR)

    def view_latest_log(self):
        try:
            with open(LOG_PATH, "r", encoding="utf-8") as f:
                tail = f.readlines()[-80:]
            messagebox.showinfo("Latest Log", "".join(tail) or "(log is empty)")
        except Exception as e:
            messagebox.showerror("Latest Log", f"Unable to read log: {e}")

    def open_log_folder(self):
        try:
            os.startfile(APP_DIR)  # Windows File Explorer
        except Exception:
            messagebox.showinfo("Log Folder", f"Open this folder manually:\n{APP_DIR}")

    def on_close(self):
        save_metadata("window_geometry", self.geometry())
        with open(CONFIG_PATH, "w") as f:
            self.cfg.write(f)
        self.destroy()

# ----------------- Main -----------------
if __name__ == "__main__":
    ensure_db()
    cfg = load_config()
    StickyUI(cfg).mainloop()
