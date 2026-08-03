"""
crm_counter.py

CRM Counter — Windows terminal call tracker.
Single-file version. Run with: python crm_counter.py
"""
from __future__ import annotations

# ============================================================
# STANDARD LIBRARY
# ============================================================
import csv
import logging
import logging.handlers
import os
import sqlite3
import sys
import threading
import time
from datetime import date, datetime
from pathlib import Path
from typing import Callable, cast

# ============================================================
# THIRD-PARTY
# colorama  — Windows ANSI color support
# openpyxl  — Excel export
# reportlab — PDF export
# pycaw     — Windows mic-session detection (optional)
# comtypes  — pycaw dependency
# keyboard  — global hotkey support
# tkinter   — built into Python, used for outcome popup dialog
# ============================================================
try:
    from colorama import Fore, Style, init as colorama_init
    colorama_init(autoreset=True)
    _COLORS = True
except ImportError:
    _COLORS = False

try:
    import keyboard
    _KEYBOARD = True
except ImportError:
    _KEYBOARD = False

try:
    import tkinter as tk
    from tkinter import font as tkfont
    _TKINTER = True
except ImportError:
    _TKINTER = False

if not _COLORS:
    class Fore:  # type: ignore[no-redef]
        WHITE = CYAN = GREEN = YELLOW = RED = MAGENTA = BLUE = ""
    class Style:  # type: ignore[no-redef]
        BRIGHT = RESET_ALL = ""


# ============================================================
# CONFIG
# ============================================================

BASE_DIR: Path = Path(__file__).resolve().parent
DATABASE_PATH: Path = BASE_DIR / "crm_counter.db"
EXPORT_DIR: Path = BASE_DIR / "exports"
LOG_DIR: Path = BASE_DIR / "logs"
LOG_FILE: Path = LOG_DIR / "crm_counter.log"

EXPORT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

OUTCOMES: dict[str, str] = {
    "1": "Live Conversation",
    "2": "Gatekeeper",
    "3": "Decision Maker",
    "4": "Voicemail",
    "5": "Wrong Number",
    "6": "Callback",
    "7": "Other",
}

OUTCOME_COLORS: dict[str, str] = {
    "Live Conversation": Fore.GREEN,
    "Gatekeeper":        Fore.YELLOW,
    "Decision Maker":    Fore.CYAN,
    "Voicemail":         Fore.MAGENTA,
    "Wrong Number":      Fore.RED,
    "Callback":          Fore.BLUE,
    "Other":             Fore.WHITE,
}

OUTCOME_HEX_COLORS: dict[str, str] = {
    "Live Conversation": "#27ae60",
    "Gatekeeper":        "#f39c12",
    "Decision Maker":    "#2980b9",
    "Voicemail":         "#8e44ad",
    "Wrong Number":      "#c0392b",
    "Callback":          "#16a085",
    "Other":             "#7f8c8d",
}

DATE_FORMAT: str       = "%Y-%m-%d"
TIME_FORMAT: str       = "%H:%M:%S"

AUTO_DETECT_ENABLED: bool                = False
MONITORED_PROCESS_NAMES: tuple[str, ...] = ("chrome.exe", "msedge.exe", "firefox.exe")
AUTO_DETECT_POLL_INTERVAL_SECONDS: float = 1.0

MIN_VALID_CALL_DURATION_SECONDS: int = 1
APP_NAME: str    = "CRM Counter"
APP_VERSION: str = "1.1.0"

LOG_LEVEL: str  = os.environ.get("CRM_COUNTER_LOG_LEVEL", "INFO")
LOG_FORMAT: str = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"


# ============================================================
# UTILS
# ============================================================

def setup_logger(name: str) -> logging.Logger:
    """Sets up and returns a configured logger."""
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
    if not logger.handlers:
        fh = logging.handlers.RotatingFileHandler(
            LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        fh.setFormatter(logging.Formatter(LOG_FORMAT))
        logger.addHandler(fh)
        sh = logging.StreamHandler(sys.stderr)
        sh.setLevel(logging.WARNING)
        sh.setFormatter(logging.Formatter(LOG_FORMAT))
        logger.addHandler(sh)
    logger.propagate = False
    return logger


def format_duration(seconds: int) -> str:
    """Formats seconds into MM:SS or HH:MM:SS; returns 00:00 for invalid input."""
    if not isinstance(seconds, int) or seconds < 0:
        return "00:00"
    if seconds >= 3600:
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"
    m = seconds // 60
    s = seconds % 60
    return f"{m:02d}:{s:02d}"


def get_current_date_str() -> str:
    """Returns today's date as a DATE_FORMAT string."""
    return date.today().strftime(DATE_FORMAT)


def get_current_time_str() -> str:
    """Returns the current time as a TIME_FORMAT string."""
    return datetime.now().strftime(TIME_FORMAT)


def calculate_duration_seconds(start_time_str: str, end_time_str: str) -> int:
    """Returns seconds between two TIME_FORMAT strings; handles midnight crossover."""
    try:
        start = datetime.strptime(start_time_str, TIME_FORMAT)
        end   = datetime.strptime(end_time_str,   TIME_FORMAT)
        diff  = int((end - start).total_seconds())
        if diff < 0:
            diff += 86400
        return diff
    except Exception:
        return 0


def validate_outcome_choice(choice: str) -> bool:
    """Returns True if choice maps to a valid OUTCOMES key."""
    return choice.strip() in OUTCOMES


def get_outcome_label(choice: str) -> str | None:
    """Returns the outcome label string for a valid choice key, or None."""
    return OUTCOMES.get(choice.strip())


def clear_screen() -> None:
    """Clears the Windows terminal screen."""
    os.system("cls")


# ============================================================
# COLOR HELPERS
# ============================================================

def _c(text: str, color: str, bright: bool = False) -> str:
    """Wraps text in colorama color; returns plain if colorama unavailable."""
    if not _COLORS:
        return text
    b = Style.BRIGHT if bright else ""
    return f"{b}{color}{text}{Style.RESET_ALL}"

def _bright(text: str) -> str:
    """Returns text in bright white."""
    return _c(text, Fore.WHITE, bright=True)

def _outcome_color(label: str, value: str) -> str:
    """Returns value wrapped in the color for this outcome label."""
    return _c(value, OUTCOME_COLORS.get(label, Fore.WHITE))


# ============================================================
# OUTCOME POPUP DIALOG
# ============================================================

def show_outcome_popup(duration_str: str) -> str | None:
    """Shows an always-on-top tkinter popup for outcome selection; returns label or None."""
    if not _TKINTER:
        return None

    result: list[str | None] = [None]

    def _run() -> None:
        root = tk.Tk()
        root.title("CRM Counter — Select Outcome")
        root.configure(bg="#1e1e1e")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        root.update_idletasks()
        w, h = 340, 420
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        tk.Label(root, text="Call Finished", bg="#1e1e1e", fg="#ffffff",
                 font=("Segoe UI", 14, "bold")).pack(pady=(18, 2))
        tk.Label(root, text=f"Duration:  {duration_str}", bg="#1e1e1e", fg="#00bcd4",
                 font=("Segoe UI", 11)).pack(pady=(0, 14))
        tk.Label(root, text="Select Outcome", bg="#1e1e1e", fg="#aaaaaa",
                 font=("Segoe UI", 9)).pack(pady=(0, 8))

        def select(label: str) -> None:
            result[0] = label
            root.destroy()

        for key, label in OUTCOMES.items():
            hex_color = OUTCOME_HEX_COLORS.get(label, "#555555")
            tk.Button(
                root, text=f"  {key}   {label}", anchor="w",
                bg=hex_color, fg="#ffffff", activebackground=hex_color,
                activeforeground="#ffffff", font=("Segoe UI", 10, "bold"),
                relief="flat", cursor="hand2", padx=12, pady=6,
                command=lambda l=label: select(l),
            ).pack(fill="x", padx=24, pady=3)

        tk.Button(
            root, text="✕  Dismiss (don't save)", bg="#2c2c2c", fg="#888888",
            activebackground="#333333", activeforeground="#aaaaaa",
            font=("Segoe UI", 9), relief="flat", cursor="hand2",
            command=root.destroy,
        ).pack(fill="x", padx=24, pady=(10, 18))

        root.protocol("WM_DELETE_WINDOW", root.destroy)
        root.mainloop()

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout=120)
    return result[0]


def show_confirm_popup(title: str, message: str) -> bool:
    """Shows a yes/no confirmation popup; returns True if confirmed, False otherwise."""
    if not _TKINTER:
        return False

    result: list[bool] = [False]

    def _run() -> None:
        root = tk.Tk()
        root.title(title)
        root.configure(bg="#1e1e1e")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        root.update_idletasks()
        w, h = 340, 180
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        tk.Label(root, text=title, bg="#1e1e1e", fg="#ffffff",
                 font=("Segoe UI", 12, "bold")).pack(pady=(18, 6))
        tk.Label(root, text=message, bg="#1e1e1e", fg="#aaaaaa",
                 font=("Segoe UI", 9), wraplength=300).pack(pady=(0, 16))

        btn_frame = tk.Frame(root, bg="#1e1e1e")
        btn_frame.pack(fill="x", padx=24)

        def confirm() -> None:
            result[0] = True
            root.destroy()

        tk.Button(btn_frame, text="Yes, delete", bg="#c0392b", fg="#ffffff",
                  activebackground="#e74c3c", activeforeground="#ffffff",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=12, pady=6, command=confirm).pack(side="left", fill="x", expand=True, padx=(0, 4))

        tk.Button(btn_frame, text="Cancel", bg="#2c2c2c", fg="#aaaaaa",
                  activebackground="#333333", activeforeground="#ffffff",
                  font=("Segoe UI", 10), relief="flat", cursor="hand2",
                  padx=12, pady=6, command=root.destroy).pack(side="left", fill="x", expand=True, padx=(4, 0))

        root.protocol("WM_DELETE_WINDOW", root.destroy)
        root.mainloop()

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout=30)
    return result[0]


# ============================================================
# DATABASE
# ============================================================

logger = setup_logger(__name__)


class DatabaseManager:
    """Manages SQLite connection and all CRUD operations for call records."""

    def __init__(self, db_path: Path = DATABASE_PATH) -> None:
        """Initializes DB, connects, and creates schema."""
        self.db_path = db_path
        self.connection: sqlite3.Connection | None = None
        try:
            self._connect()
            self._create_schema()
        except sqlite3.Error as e:
            logger.critical(f"Database initialization failed: {e}")
            raise RuntimeError("Failed to initialize database. See logs for details.") from e

    def _connect(self) -> None:
        """Opens the SQLite connection."""
        self.connection = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.connection.row_factory = sqlite3.Row
        logger.info(f"Database connected at {self.db_path}")

    def _create_schema(self) -> None:
        """Creates the calls table and date index if they don't already exist."""
        assert self.connection is not None
        cursor = self.connection.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS calls (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                date             TEXT NOT NULL,
                start_time       TEXT NOT NULL,
                end_time         TEXT NOT NULL,
                duration_seconds INTEGER NOT NULL,
                outcome          TEXT NOT NULL
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_calls_date ON calls(date)")
        self.connection.commit()
        logger.info("Database schema verified/created.")

    def insert_call(self, date: str, start_time: str, end_time: str,
                    duration_seconds: int, outcome: str) -> int | None:
        """Inserts a call record; returns new row id or None on failure."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return None
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO calls (date, start_time, end_time, duration_seconds, outcome) "
                "VALUES (?, ?, ?, ?, ?)",
                (date, start_time, end_time, duration_seconds, outcome)
            )
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            logger.error(f"Failed to insert call: {e}")
            return None

    def get_calls_by_date(self, date: str) -> list[sqlite3.Row]:
        """Returns all call rows for date ordered by id asc."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return []
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM calls WHERE date = ? ORDER BY id ASC", (date,))
            return cursor.fetchall()
        except sqlite3.Error as e:
            logger.error(f"Failed to get calls by date {date}: {e}")
            return []

    def get_all_dates_with_calls(self) -> list[str]:
        """Returns distinct dates with calls, newest first."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return []
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT DISTINCT date FROM calls ORDER BY date DESC")
            return [row["date"] for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"Failed to get all dates with calls: {e}")
            return []

    def get_dial_count_for_date(self, date: str) -> int:
        """Returns count of call rows for date."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return 0
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT COUNT(*) FROM calls WHERE date = ?", (date,))
            result = cursor.fetchone()
            return result[0] if result else 0
        except sqlite3.Error as e:
            logger.error(f"Failed to get dial count for date {date}: {e}")
            return 0

    def get_all_calls(self) -> list[sqlite3.Row]:
        """Returns all call rows ordered by date desc, id asc."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return []
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM calls ORDER BY date DESC, id ASC")
            return cursor.fetchall()
        except sqlite3.Error as e:
            logger.error(f"Failed to get all calls: {e}")
            return []

    def get_last_call(self) -> sqlite3.Row | None:
        """Returns the most recently inserted call row, or None if no calls exist."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return None
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM calls ORDER BY id DESC LIMIT 1")
            return cursor.fetchone()
        except sqlite3.Error as e:
            logger.error(f"Failed to get last call: {e}")
            return None

    def delete_calls_by_date(self, date: str) -> int:
        """Deletes all calls for the given date; returns count of deleted rows."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return 0
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM calls WHERE date = ?", (date,))
            self.connection.commit()
            deleted = cursor.rowcount
            logger.info(f"Deleted {deleted} calls for date {date}")
            return deleted
        except sqlite3.Error as e:
            logger.error(f"Failed to delete calls for date {date}: {e}")
            return 0

    def delete_call_by_id(self, call_id: int) -> bool:
        """Deletes a single call by id; returns True on success."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return False
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM calls WHERE id = ?", (call_id,))
            self.connection.commit()
            logger.info(f"Deleted call id={call_id}")
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"Failed to delete call id={call_id}: {e}")
            return False

    def delete_all_calls(self) -> int:
        """Deletes every call record; returns count of deleted rows."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return 0
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM calls")
            self.connection.commit()
            deleted = cursor.rowcount
            logger.info(f"Deleted all {deleted} call records.")
            return deleted
        except sqlite3.Error as e:
            logger.error(f"Failed to delete all calls: {e}")
            return 0

    def get_all_time_stats(self) -> dict[str, object]:
        """Returns all-time stats: total calls, best day, top outcome, avg duration."""
        if self.connection is None:
            logger.error("Database connection is not available.")
            return {}
        try:
            cursor = self.connection.cursor()

            # Total calls
            cursor.execute("SELECT COUNT(*) FROM calls")
            total = cursor.fetchone()[0]

            if total == 0:
                return {"total": 0, "best_day": None, "best_day_count": 0,
                        "top_outcome": None, "avg_duration": 0}

            # Best day
            cursor.execute("""
                SELECT date, COUNT(*) as cnt FROM calls
                GROUP BY date ORDER BY cnt DESC LIMIT 1
            """)
            best = cursor.fetchone()

            # Top outcome
            cursor.execute("""
                SELECT outcome, COUNT(*) as cnt FROM calls
                GROUP BY outcome ORDER BY cnt DESC LIMIT 1
            """)
            top = cursor.fetchone()

            # Avg duration
            cursor.execute("SELECT AVG(duration_seconds) FROM calls")
            avg = cursor.fetchone()[0] or 0

            return {
                "total":          total,
                "best_day":       best["date"] if best else None,
                "best_day_count": best["cnt"]  if best else 0,
                "top_outcome":    top["outcome"] if top else None,
                "avg_duration":   round(avg),
            }
        except sqlite3.Error as e:
            logger.error(f"Failed to get all-time stats: {e}")
            return {}

    def close(self) -> None:
        """Closes the database connection safely."""
        if self.connection is not None:
            self.connection.close()
            self.connection = None
            logger.info("Database connection closed.")


# ============================================================
# TRACKER
# ============================================================

class CallTracker:
    """Manages manual call timing state."""

    def __init__(self, db: DatabaseManager) -> None:
        """Initializes CallTracker with a DatabaseManager instance."""
        self.db                             = db
        self.active: bool                   = False
        self.current_date: str | None       = None
        self.current_start_time: str | None = None
        logger.info("CallTracker initialized.")

    def is_active(self) -> bool:
        """Returns True if a call is currently being timed."""
        return self.active

    def start_call(self) -> bool:
        """Starts a call timer; returns False if already active."""
        if self.active:
            logger.warning("Call already in progress; ignoring start.")
            return False
        self.current_date       = get_current_date_str()
        self.current_start_time = get_current_time_str()
        self.active             = True
        logger.info(f"Call started at {self.current_start_time}")
        return True

    def end_call(self) -> dict[str, object] | None:
        """Ends the active call; returns call data dict or None if inactive/too short."""
        if not self.active:
            logger.warning("No active call to end; ignoring end.")
            return None
        end_time = get_current_time_str()
        assert self.current_start_time is not None
        duration_seconds = calculate_duration_seconds(self.current_start_time, end_time)
        call_data = {
            "date":             self.current_date,
            "start_time":       self.current_start_time,
            "end_time":         end_time,
            "duration_seconds": duration_seconds,
        }
        self.active             = False
        self.current_date       = None
        self.current_start_time = None
        if duration_seconds < MIN_VALID_CALL_DURATION_SECONDS:
            logger.warning(f"Call duration {duration_seconds}s below minimum; discarding.")
            return None
        logger.info(f"Call ended. Duration: {duration_seconds}s")
        return call_data

    def save_completed_call(self, call_data: dict[str, object], outcome: str) -> int | None:
        """Persists a completed call and returns today's dial number, or None on failure."""
        date             = cast(str, call_data["date"])
        start_time       = cast(str, call_data["start_time"])
        end_time         = cast(str, call_data["end_time"])
        duration_seconds = cast(int, call_data["duration_seconds"])
        row_id = self.db.insert_call(
            date=date, start_time=start_time,
            end_time=end_time, duration_seconds=duration_seconds, outcome=outcome
        )
        if row_id is None:
            logger.error("Failed to save completed call to database.")
            return None
        return self.db.get_dial_count_for_date(date)


class AudioSessionMonitor:
    """Optional Windows-only WASAPI mic-session monitor."""

    def __init__(self, on_call_start: Callable[[], None], on_call_end: Callable[[], None]) -> None:
        """Initializes the monitor with start/end callbacks."""
        self.on_call_start         = on_call_start
        self.on_call_end           = on_call_end
        self._thread: threading.Thread | None = None
        self._stop_event           = threading.Event()
        self._session_active: bool = False
        logger.info("AudioSessionMonitor initialized.")

    def is_available(self) -> bool:
        """Returns True if pycaw is installed."""
        try:
            import pycaw.pycaw  # noqa: F401
            return True
        except ImportError:
            return False

    def start(self) -> bool:
        """Starts background polling thread; returns False if disabled or pycaw missing."""
        if not AUTO_DETECT_ENABLED:
            logger.info("Auto-detect disabled in config; skipping monitor start.")
            return False
        if not self.is_available():
            logger.warning("pycaw not available. Run: pip install pycaw comtypes")
            return False
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._poll_loop, daemon=True)
        self._thread.start()
        logger.info("AudioSessionMonitor started.")
        return True

    def stop(self) -> None:
        """Stops the background thread."""
        if self._thread is not None:
            self._stop_event.set()
            self._thread.join(timeout=2.0)
            self._thread = None
        logger.info("AudioSessionMonitor stopped.")

    def _poll_loop(self) -> None:
        """Polls mic-session state; triggers callbacks on state changes."""
        while not self._stop_event.is_set():
            try:
                is_active_now = self._is_mic_capturing()
                if is_active_now and not self._session_active:
                    self._session_active = True
                    self.on_call_start()
                elif not is_active_now and self._session_active:
                    self._session_active = False
                    self.on_call_end()
            except Exception as e:
                logger.error(f"AudioSessionMonitor poll error: {e}")
            self._stop_event.wait(timeout=AUTO_DETECT_POLL_INTERVAL_SECONDS)

    def _is_mic_capturing(self) -> bool:
        """Returns True if a monitored process has an active WASAPI capture session."""
        try:
            from pycaw.pycaw import AudioUtilities
            for session in AudioUtilities.GetAllSessions():
                if session.Process is not None:
                    proc = session.Process.name().lower()
                    if any(m.lower() in proc for m in MONITORED_PROCESS_NAMES):
                        if session.State == 1:
                            return True
            return False
        except Exception as e:
            logger.error(f"Failed to check mic capture state: {e}")
            return False


# ============================================================
# REPORTS
# ============================================================

class ReportGenerator:
    """Generates call reports and exports to CSV, Excel, and PDF."""

    def __init__(self, db: DatabaseManager) -> None:
        """Initializes ReportGenerator with a DatabaseManager instance."""
        self.db = db
        logger.info("ReportGenerator initialized.")

    def generate_report_for_date(self, date: str) -> dict[str, object]:
        """Builds a report dict for the given date."""
        rows = self.db.get_calls_by_date(date)
        outcome_counts: dict[str, int] = {label: 0 for label in OUTCOMES.values()}
        for row in rows:
            outcome = row["outcome"]
            if outcome in outcome_counts:
                outcome_counts[outcome] += 1
        total_dials = len(rows)
        avg_secs = round(sum(r["duration_seconds"] for r in rows) / total_dials) if total_dials else 0
        return {
            "date":                       date,
            "total_dials":                total_dials,
            "outcome_counts":             outcome_counts,
            "average_duration_seconds":   avg_secs,
            "average_duration_formatted": format_duration(avg_secs),
        }

    def generate_today_report(self) -> dict[str, object]:
        """Builds a report dict for today."""
        return self.generate_report_for_date(get_current_date_str())

    def format_report_text(self, report: dict[str, object]) -> str:
        """Formats a report dict into a colored terminal string."""
        oc   = cast(dict[str, int], report["outcome_counts"])
        sep  = _c("==================================", Fore.CYAN)
        sep2 = _c("----------------------------------", Fore.CYAN)
        lines = [
            sep,
            _c(f" Report for {report['date']}", Fore.CYAN, bright=True),
            sep,
            _bright("Total Dials:") + f"            {_bright(str(report['total_dials']))}",
            _outcome_color("Live Conversation", f"Live Conversations:     {oc.get('Live Conversation', 0)}"),
            _outcome_color("Gatekeeper",        f"Gatekeepers:            {oc.get('Gatekeeper', 0)}"),
            _outcome_color("Decision Maker",    f"Decision Makers:        {oc.get('Decision Maker', 0)}"),
            _outcome_color("Voicemail",         f"Voicemails:             {oc.get('Voicemail', 0)}"),
            _outcome_color("Wrong Number",      f"Wrong Numbers:          {oc.get('Wrong Number', 0)}"),
            _outcome_color("Callback",          f"Callbacks:              {oc.get('Callback', 0)}"),
            _outcome_color("Other",             f"Others:                 {oc.get('Other', 0)}"),
            sep2,
            f"Average Call Duration:  {_c(str(report['average_duration_formatted']), Fore.CYAN)}",
            sep,
        ]
        return "\n".join(lines)

    def get_history_summary(self) -> list[dict[str, object]]:
        """Returns report dicts for all dates with calls, newest first."""
        return [self.generate_report_for_date(d) for d in self.db.get_all_dates_with_calls()]

    def format_history_text(self, history: list[dict[str, object]]) -> str:
        """Formats history list into a colored terminal string."""
        if not history:
            return _c("No call history found.", Fore.YELLOW)
        lines = [_c("Call History", Fore.CYAN, bright=True)]
        for r in history:
            lines.append(
                f"{_c(str(r['date']), Fore.CYAN)}  |  "
                f"Dials: {_bright(str(r['total_dials']))}  |  "
                f"Avg Duration: {_c(str(r['average_duration_formatted']), Fore.YELLOW)}"
            )
        return "\n".join(lines)

    def format_stats_text(self, stats: dict[str, object]) -> str:
        """Formats all-time stats dict into a colored terminal string."""
        sep  = _c("==================================", Fore.CYAN)
        sep2 = _c("----------------------------------", Fore.CYAN)
        if not stats or stats.get("total", 0) == 0:
            return _c("  No call data yet.", Fore.YELLOW)
        top_outcome = str(stats.get("top_outcome") or "N/A")
        lines = [
            sep,
            _c(" All-Time Stats", Fore.CYAN, bright=True),
            sep,
            _bright("Total Calls Ever:  ") + _bright(str(stats["total"])),
            _c(f"Best Day:          {stats['best_day']}  ({stats['best_day_count']} dials)", Fore.GREEN),
            _c(f"Top Outcome:       {top_outcome}", OUTCOME_COLORS.get(top_outcome, Fore.WHITE)),
            _c(f"Avg Call Duration: {format_duration(cast(int, stats['avg_duration']))}", Fore.CYAN),
            sep2,
        ]
        return "\n".join(lines)

    def export_csv(self) -> Path | None:
        """Exports all calls to a CSV file; returns path or None on failure."""
        rows = self.db.get_all_calls()
        if not rows:
            logger.warning("No call data to export.")
            return None
        filepath = EXPORT_DIR / f"crm_counter_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        try:
            with open(filepath, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["ID", "Date", "Start Time", "End Time", "Duration (seconds)", "Outcome"])
                for row in rows:
                    w.writerow([row["id"], row["date"], row["start_time"],
                                row["end_time"], row["duration_seconds"], row["outcome"]])
            logger.info(f"CSV export saved to {filepath}")
            return filepath
        except Exception as e:
            logger.error(f"Failed to export CSV: {e}")
            return None

    def export_excel(self) -> Path | None:
        """Exports all calls to an Excel file; returns path or None on failure."""
        rows = self.db.get_all_calls()
        if not rows:
            logger.warning("No call data to export.")
            return None
        filepath = EXPORT_DIR / f"crm_counter_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            from openpyxl import Workbook
        except ImportError:
            logger.error("openpyxl not installed. Run: pip install openpyxl")
            return None
        try:
            wb = Workbook()
            ws = wb.active
            assert ws is not None
            ws.title = "Call Log"
            ws.append(["ID", "Date", "Start Time", "End Time", "Duration (seconds)", "Outcome"])
            for row in rows:
                ws.append([row["id"], row["date"], row["start_time"],
                           row["end_time"], row["duration_seconds"], row["outcome"]])
            wb.save(str(filepath))
            logger.info(f"Excel export saved to {filepath}")
            return filepath
        except Exception as e:
            logger.error(f"Failed to export Excel: {e}")
            return None

    def export_pdf(self) -> Path | None:
        """Exports all calls to a PDF file; returns path or None on failure."""
        rows = self.db.get_all_calls()
        if not rows:
            logger.warning("No call data to export.")
            return None
        filepath = EXPORT_DIR / f"crm_counter_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        except ImportError:
            logger.error("reportlab not installed. Run: pip install reportlab")
            return None
        try:
            table_data = [["ID", "Date", "Start Time", "End Time", "Duration (s)", "Outcome"]]
            for row in rows:
                table_data.append([str(row["id"]), str(row["date"]), str(row["start_time"]),
                                   str(row["end_time"]), str(row["duration_seconds"]), str(row["outcome"])])
            doc   = SimpleDocTemplate(str(filepath), pagesize=letter)
            table = Table(table_data)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
                ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
                ("GRID",       (0, 0), (-1, -1), 0.5, colors.black),
                ("PADDING",    (0, 0), (-1, -1), 6),
            ]))
            doc.build([table])
            logger.info(f"PDF export saved to {filepath}")
            return filepath
        except Exception as e:
            logger.error(f"Failed to export PDF: {e}")
            return None


# ============================================================
# MAIN APP
# ============================================================

class CRMCounterApp:
    """Main CLI application — command loop and component coordination."""

    def __init__(self) -> None:
        """Initializes all component references to None before startup."""
        self.db:      DatabaseManager | None     = None
        self.tracker: CallTracker | None         = None
        self.reports: ReportGenerator | None     = None
        self.monitor: AudioSessionMonitor | None = None
        self.running: bool                       = True
        self._quit_confirm_pending: bool         = False
        self._interrupt_count: int               = 0

    def initialize(self) -> bool:
        """Starts all components; returns False if database fails to connect."""
        clear_screen()
        print(_bright(f"\n  {APP_NAME}  v{APP_VERSION}"))
        print(_c("  ─────────────────────────────", Fore.CYAN))
        print()
        try:
            self.db = DatabaseManager()
            print(_c("  ✔ Database Connected", Fore.GREEN))
        except Exception as e:
            logger.critical(f"Database initialization failed: {e}", exc_info=True)
            print(_c("  ✘ Failed to connect to database. Check logs. Exiting.", Fore.RED))
            return False

        self.tracker = CallTracker(self.db)
        self.reports = ReportGenerator(self.db)
        self.monitor = AudioSessionMonitor(
            on_call_start=self._handle_auto_start,
            on_call_end=self._handle_auto_end,
        )
        self.monitor.start()

        print()
        print(_c("  Waiting for Calls...", Fore.YELLOW))
        print()
        print(_c("  Type 'help' for commands.", Fore.WHITE))

        if _KEYBOARD:
            keyboard.add_hotkey("`", self._hotkey_toggle, suppress=True)
            print(_c("  Hotkey: press  `  to start / end a call.", Fore.CYAN))
        else:
            print(_c("  Tip: pip install keyboard  to enable hotkey.", Fore.YELLOW))

        if _TKINTER:
            print(_c("  Popup: outcome dialog will appear automatically.", Fore.CYAN))
        else:
            print(_c("  Note: tkinter not found — using terminal outcome menu.", Fore.YELLOW))

        print()
        return True

    # ----------------------------------------------------------
    # HOTKEY
    # ----------------------------------------------------------

    def _hotkey_toggle(self) -> None:
        """Toggles call start/end on backtick press; suppresses the character."""
        assert self.tracker is not None
        if self.tracker.is_active():
            threading.Thread(target=self._finish_call, daemon=True).start()
        else:
            self.tracker.start_call()
            print(_c("\n  ● Call started.  (press ` to end)", Fore.GREEN, bright=True))
            print(_c("> ", Fore.CYAN), end="", flush=True)

    # ----------------------------------------------------------
    # AUTO-DETECT CALLBACKS
    # ----------------------------------------------------------

    def _handle_auto_start(self) -> None:
        """Callback: auto-starts a call when mic session is detected."""
        assert self.tracker is not None
        if not self.tracker.is_active():
            self.tracker.start_call()
            print(_c("\n  [Auto] Call started.", Fore.GREEN))

    def _handle_auto_end(self) -> None:
        """Callback: auto-ends a call when mic session ends."""
        assert self.tracker is not None
        if self.tracker.is_active():
            threading.Thread(target=self._finish_call, daemon=True).start()

    # ----------------------------------------------------------
    # CORE CALL FLOW
    # ----------------------------------------------------------

    def _finish_call(self) -> None:
        """Ends timer → shows outcome popup (or terminal menu) → saves → prints result."""
        assert self.tracker is not None
        call_data = self.tracker.end_call()
        if call_data is None:
            print(_c("\n  Call was too short to record.", Fore.YELLOW))
            print(_c("> ", Fore.CYAN), end="", flush=True)
            return

        dur = format_duration(cast(int, call_data["duration_seconds"]))

        outcome = show_outcome_popup(dur)
        if outcome is None and not _TKINTER:
            outcome = self._terminal_outcome_menu()
        if outcome is None:
            print(_c(f"\n  Call dismissed (not saved).  Duration was {dur}", Fore.YELLOW))
            print(_c("> ", Fore.CYAN), end="", flush=True)
            return

        dial_number = self.tracker.save_completed_call(call_data, outcome)
        if dial_number is None:
            print(_c("\n  Warning: call could not be saved. Check logs.", Fore.RED))
            print(_c("> ", Fore.CYAN), end="", flush=True)
            return

        print()
        print(_c(f"  Dial #{dial_number} Finished", Fore.WHITE, bright=True))
        print(_c(f"  Duration : {dur}", Fore.CYAN))
        print(_c(f"  Outcome  : {outcome}", OUTCOME_COLORS.get(outcome, Fore.WHITE)))
        print()
        print(_c("> ", Fore.CYAN), end="", flush=True)

    def _terminal_outcome_menu(self) -> str | None:
        """Fallback terminal outcome menu when tkinter is unavailable."""
        print()
        print(_bright("  Select Outcome"))
        print()
        for key, label in OUTCOMES.items():
            print(f"  {_c(key, Fore.CYAN)}  {_c(label, OUTCOME_COLORS.get(label, Fore.WHITE))}")
        print()
        while True:
            try:
                choice = input(_c("  > ", Fore.CYAN)).strip()
                if validate_outcome_choice(choice):
                    return cast(str, get_outcome_label(choice))
                print(_c("  Invalid choice. Enter a number 1-7.", Fore.RED))
            except (EOFError, KeyboardInterrupt):
                return None

    def _confirm_delete(self, message: str) -> bool:
        """Asks for delete confirmation via popup or terminal fallback."""
        if _TKINTER:
            return show_confirm_popup("Confirm Delete", message)
        # Terminal fallback
        print(_c(f"\n  {message}", Fore.YELLOW))
        print(_c("  Type 'yes' to confirm, anything else to cancel: ", Fore.YELLOW), end="")
        try:
            return input().strip().lower() == "yes"
        except (EOFError, KeyboardInterrupt):
            return False

    # ----------------------------------------------------------
    # COMMAND HANDLERS
    # ----------------------------------------------------------

    def run(self) -> None:
        """Main command loop."""
        assert self.db is not None
        assert self.tracker is not None
        assert self.reports is not None
        while self.running:
            try:
                command = input(_c("> ", Fore.CYAN)).strip().lower()
                self._interrupt_count = 0

                if   command == "start":           self._handle_start()
                elif command == "end":             self._handle_end()
                elif command == "today":           self._handle_today()
                elif command == "history":         self._handle_history()
                elif command == "stats":           self._handle_stats()
                elif command == "clear today":     self._handle_clear_today()
                elif command == "clear last":      self._handle_clear_last()
                elif command == "clear history":   self._handle_clear_history()
                elif command == "export csv":      self._handle_export("csv")
                elif command == "export excel":    self._handle_export("excel")
                elif command == "export pdf":      self._handle_export("pdf")
                elif command == "help":            self._handle_help()
                elif command in ("quit", "exit"):  self._handle_quit()
                elif command == "":                pass
                else:
                    print(_c(f"  Unknown command: '{command}'. Type 'help'.", Fore.RED))

                if command not in ("quit", "exit"):
                    self._quit_confirm_pending = False

            except KeyboardInterrupt:
                self._interrupt_count += 1
                if self._interrupt_count >= 2:
                    self.running = False
                    break
                print(_c("\n  Interrupted. Type 'quit' to exit, or Ctrl+C again to force.", Fore.YELLOW))
            except Exception as e:
                logger.error(f"Unexpected error in main loop: {e}", exc_info=True)
                print(_c("  An unexpected error occurred. Check logs. Continuing.", Fore.RED))

        self.shutdown()

    def _handle_start(self) -> None:
        """Starts a call timer if none is active."""
        assert self.tracker is not None
        if self.tracker.is_active():
            print(_c("  A call is already in progress.", Fore.YELLOW))
            return
        self.tracker.start_call()
        print(_c("  ● Call started.  (press ` to end, or type 'end')", Fore.GREEN, bright=True))

    def _handle_end(self) -> None:
        """Ends the active call."""
        assert self.tracker is not None
        if not self.tracker.is_active():
            print(_c("  No call is currently active.", Fore.YELLOW))
            return
        self._finish_call()

    def _handle_today(self) -> None:
        """Prints today's call report."""
        assert self.reports is not None
        print(self.reports.format_report_text(self.reports.generate_today_report()))

    def _handle_history(self) -> None:
        """Prints call history across all days."""
        assert self.reports is not None
        print(self.reports.format_history_text(self.reports.get_history_summary()))

    def _handle_stats(self) -> None:
        """Prints all-time performance stats."""
        assert self.db is not None
        assert self.reports is not None
        stats = self.db.get_all_time_stats()
        print(self.reports.format_stats_text(stats))

    def _handle_clear_today(self) -> None:
        """Deletes all calls logged today after confirmation."""
        assert self.db is not None
        today = get_current_date_str()
        count = self.db.get_dial_count_for_date(today)
        if count == 0:
            print(_c("  No calls logged today to clear.", Fore.YELLOW))
            return
        confirmed = self._confirm_delete(
            f"Delete all {count} calls from today ({today})?"
        )
        if confirmed:
            deleted = self.db.delete_calls_by_date(today)
            print(_c(f"  ✔ Cleared {deleted} calls from today.", Fore.GREEN))
        else:
            print(_c("  Cancelled.", Fore.YELLOW))

    def _handle_clear_last(self) -> None:
        """Deletes the most recent call entry after confirmation."""
        assert self.db is not None
        last = self.db.get_last_call()
        if last is None:
            print(_c("  No calls on record to clear.", Fore.YELLOW))
            return
        confirmed = self._confirm_delete(
            f"Delete last call? ({last['date']}  {last['start_time']}  {last['outcome']})"
        )
        if confirmed:
            ok = self.db.delete_call_by_id(last["id"])
            if ok:
                print(_c("  ✔ Last call deleted.", Fore.GREEN))
            else:
                print(_c("  Failed to delete. Check logs.", Fore.RED))
        else:
            print(_c("  Cancelled.", Fore.YELLOW))

    def _handle_clear_history(self) -> None:
        """Deletes ALL call records after double confirmation."""
        assert self.db is not None
        all_calls = self.db.get_all_calls()
        if not all_calls:
            print(_c("  No call history to clear.", Fore.YELLOW))
            return
        total = len(all_calls)
        # First confirmation
        confirmed = self._confirm_delete(
            f"This will permanently delete ALL {total} call records. Are you sure?"
        )
        if not confirmed:
            print(_c("  Cancelled.", Fore.YELLOW))
            return
        # Second confirmation — extra safety for full wipe
        confirmed2 = self._confirm_delete(
            "Final warning: this cannot be undone. Delete everything?"
        )
        if not confirmed2:
            print(_c("  Cancelled.", Fore.YELLOW))
            return
        deleted = self.db.delete_all_calls()
        print(_c(f"  ✔ All {deleted} call records deleted.", Fore.GREEN))

    def _handle_export(self, fmt: str) -> None:
        """Exports calls to the specified format."""
        assert self.reports is not None
        if   fmt == "csv":   path = self.reports.export_csv()
        elif fmt == "excel": path = self.reports.export_excel()
        elif fmt == "pdf":   path = self.reports.export_pdf()
        else: return
        if path:
            print(_c(f"  Export saved to: {path}", Fore.GREEN))
        else:
            print(_c("  Export failed or no data yet. Check logs.", Fore.RED))

    def _handle_help(self) -> None:
        """Prints the command reference."""
        pad = "  "
        sep = _c("  ──────────────────────────────────────────", Fore.CYAN)
        print()
        print(_bright("  Commands"))
        print(sep)
        cmds = [
            ("start",          "Start a call timer"),
            ("end",            "End the current call"),
            ("today",          "Show today's report"),
            ("history",        "Show call history across all days"),
            ("stats",          "Show all-time performance stats"),
            ("clear today",    "Delete all calls logged today"),
            ("clear last",     "Delete the most recent call entry"),
            ("clear history",  "Delete ALL call records (double confirmed)"),
            ("export csv",     "Export all calls to a CSV file"),
            ("export excel",   "Export all calls to an Excel file"),
            ("export pdf",     "Export all calls to a PDF file"),
            ("help",           "Show this help message"),
            ("quit",           "Exit the application safely"),
        ]
        for cmd, desc in cmds:
            print(f"{pad}{_c(cmd.ljust(18), Fore.CYAN)}  {desc}")
        print(sep)
        print(_c("  Hotkey:  `  (backtick) = start / end call", Fore.CYAN))
        print(sep)
        print()

    def _handle_quit(self) -> None:
        """Quits safely; warns if a call is still active."""
        assert self.tracker is not None
        if self.tracker.is_active():
            if self._quit_confirm_pending:
                self.running = False
                print(_c("  Shutting down CRM Counter. Goodbye.", Fore.YELLOW))
            else:
                self._quit_confirm_pending = True
                print(_c("  Warning: a call is still active. Type 'end' to save it, "
                          "or 'quit' again to discard and exit.", Fore.YELLOW))
        else:
            self.running = False
            print(_c("  Shutting down CRM Counter. Goodbye.", Fore.YELLOW))

    def shutdown(self) -> None:
        """Stops the monitor and closes the database on exit."""
        try:
            if self.monitor is not None:
                self.monitor.stop()
            if self.db is not None:
                self.db.close()
        except Exception as e:
            logger.error(f"Error during shutdown: {e}", exc_info=True)


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    app = CRMCounterApp()
    if app.initialize():
        app.run()
    else:
        sys.exit(1)