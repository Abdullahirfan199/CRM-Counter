"""
crm_counter.py

CRM Counter — Windows terminal call tracker.
Single-file version. Run with: python crm_counter.py
"""
from __future__ import annotations

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
    "1": "Gatekeeper",
    "2": "Voicemail",
    "3": "Direct Decision Maker",
    "4": "Answering Machine",
    "5": "Callback",
    "6": "Not Interested",
    "7": "Unsuccessful",
}

OUTCOME_COLORS: dict[str, str] = {
    "Gatekeeper":             Fore.YELLOW,
    "Voicemail":              Fore.MAGENTA,
    "Direct Decision Maker":  Fore.GREEN,
    "Answering Machine":      Fore.WHITE,
    "Callback":               Fore.BLUE,
    "Not Interested":         Fore.RED,
    "Unsuccessful":           Fore.CYAN,
}

OUTCOME_HEX_COLORS: dict[str, str] = {
    "Gatekeeper":             "#f39c12",
    "Voicemail":              "#8e44ad",
    "Direct Decision Maker":  "#27ae60",
    "Answering Machine":      "#7f8c8d",
    "Callback":               "#16a085",
    "Not Interested":         "#c0392b",
    "Unsuccessful":           "#2980b9",
}

DATE_FORMAT: str       = "%Y-%m-%d"
TIME_FORMAT: str       = "%H:%M:%S"

AUTO_DETECT_ENABLED: bool                = False
MONITORED_PROCESS_NAMES: tuple[str, ...] = ("chrome.exe", "msedge.exe", "firefox.exe")
AUTO_DETECT_POLL_INTERVAL_SECONDS: float = 1.0

MIN_VALID_CALL_DURATION_SECONDS: int = 1
APP_NAME: str    = "CRM Counter"
APP_VERSION: str = "1.2.0"

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
            cursor.execute("SELECT COUNT(*) FROM calls")
            total = cursor.fetchone()[0]
            if total == 0:
                return {"total": 0, "best_day": None, "best_day_count": 0,
                        "top_outcome": None, "avg_duration": 0}
            cursor.execute("""
                SELECT date, COUNT(*) as cnt FROM calls
                GROUP BY date ORDER BY cnt DESC LIMIT 1
            """)
            best = cursor.fetchone()
            cursor.execute("""
                SELECT outcome, COUNT(*) as cnt FROM calls
                GROUP BY outcome ORDER BY cnt DESC LIMIT 1
            """)
            top = cursor.fetchone()
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
        """Returns True if a call is currently being timed.""