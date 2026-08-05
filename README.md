# CRM Counter

A lightweight, fast, Windows terminal call tracker for sales reps using web-based CRMs.

No subscriptions. No cloud. No bloat. Runs locally, stores everything in a single file, launches with one double-click.

---

## What it does

- Tracks every call with a start/end timer
- Pops up an outcome selection dialog the moment a call ends
- Logs date, time, duration, and outcome to a local SQLite database
- Shows today's report and full call history in the terminal
- Exports to CSV, Excel, or PDF
- Global hotkey (`` ` `` backtick) starts and ends calls without touching the terminal

---

## Requirements

- Windows 10 or 11
- Python 3.12 or newer

---

## Installation

### Step 1 — Download the files

Download these two files and put them in the same folder (e.g. `C:\CRM Counter\`):

```
crm_counter.py
setup.bat
CRM_Counter.bat
```

### Step 2 — Install Python

Go to **https://www.python.org/downloads/** and download Python 3.12+.

When the installer opens, **tick every checkbox on the first screen** — especially:
- ✅ Add Python to PATH
- ✅ Install pip

Then click **Install Now**.

### Step 3 — Run setup

Double-click `setup.bat`.

It will check your Python installation and install all required packages automatically. You only need to do this once per machine.

### Step 4 — Launch

Double-click `CRM_Counter.bat` to start the app.

> **Note:** Right-click `CRM_Counter.bat` → Properties → Shortcut → Advanced → tick **Run as administrator**. This is required for the global hotkey to work.

---

## Usage

### Hotkey
Press `` ` `` (backtick, top-left of keyboard) anywhere on your screen to start a call. Press it again when the call ends — an outcome dialog will pop up on top of your CRM automatically.

### Commands

| Command | Description |
|---|---|
| `start` | Start a call timer |
| `end` | End the current call |
| `today` | Show today's report |
| `history` | Show call history across all days |
| `stats` | Show all-time performance stats |
| `clear today` | Delete all calls logged today |
| `clear last` | Delete the most recent call entry |
| `clear history` | Delete ALL call records (double confirmed) |
| `export csv` | Export all calls to a CSV file |
| `export excel` | Export all calls to an Excel file |
| `export pdf` | Export all calls to a PDF file |
| `help` | Show command list |
| `quit` | Exit safely |

---

## Sharing with a colleague

Copy the folder to their machine (USB, email, Google Drive — anything). They run `setup.bat` once, then `CRM_Counter.bat` to launch. Each person gets their own local database — data never leaves the machine.

---

## File structure

```
CRM Counter/
    crm_counter.py      — the application
    CRM_Counter.bat     — launcher (double-click to run)
    setup.bat           — first-time setup (run once per machine)
    crm_counter.db      — your call data (auto-created, do not share)
    exports/            — CSV / Excel / PDF exports land here
    logs/               — application logs land here
```

---

## Troubleshooting

**"python is not recognized"**
Python isn't on PATH. Re-run the Python installer and tick all checkboxes on the first screen.

**Hotkey not working**
The app must be run as administrator. Right-click `CRM_Counter.bat` → Run as administrator.

**Outcome popup not appearing**
tkinter is built into Python — if it's missing, re-install Python and tick all checkboxes. The app falls back to the terminal menu automatically.

**Export failing**
Run `setup.bat` again to reinstall packages.

---

## Built with

- Python 3.12
- SQLite (built-in)
- tkinter (built-in) — outcome popup dialog
- colorama — terminal colors
- keyboard — global hotkey
- openpyxl — Excel export
- reportlab — PDF export

---

*Built for sales reps who want to track calls without the overhead of a full CRM.*
