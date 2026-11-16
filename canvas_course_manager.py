"""Canvas Course Manager: build Canvas modules from the lesson folders on your computer.

FOR TEACHERS (you never need to open or edit this file)
  Windows: double-click  "Start Canvas Course Manager.bat"
  Mac:     double-click  "Start Canvas Course Manager.command"

Everything is set up inside the window:
  Course setup tab   your Canvas address, access token, course ID, and course folder
  Names & look tab   folder words, emojis, and indentation, with a live sample module
  Run tab            what to upload, where, preview or apply, and the progress panel

Files kept next to this program (they belong to you and this computer):
  canvas_config.txt        your Canvas address and access token (PRIVATE, never share it)
  program_settings.json    which course folders you have used lately, light or dark mode

Files kept inside each course folder (they belong to that course):
  course_settings.json               course ID, spreadsheet, words, emojis, indents, and
                                     which gradebook group each kind of quiz goes into
  .canvas_lesson_uploader_state.json what has already been uploaded
  canvas_lesson_uploader.log         a record of every run

The program can live in any folder. Each course folder remembers its own setup,
so one copy of the program can manage several courses.

Safety rules that cannot be turned off: run a preview first, live changes need the
word APPLY, nothing in Canvas is ever deleted, answer keys are always hidden from
students, and unclear title matches are skipped instead of guessed. Gradebook groups are
only created when you ask and type APPLY, and a quiz someone already put in a group by
hand is never moved.

Updates: when a newer version is published, the window offers to install it. Your
settings, token, and course files are never touched; only this program file is
replaced, and the old copy is saved in the "backups" folder first.

Requires Python 3.10+. The settings spreadsheet also needs pandas and openpyxl, and the
modern look (with an optional dark mode) uses sv-ttk; the window offers to install them
the first time. Without sv-ttk the program still works, it just looks plainer.
"""
from __future__ import annotations

import sys


def _fatal_popup(message: str) -> None:
    """Show a startup problem in a window (the .pyw launcher has no console)."""
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        root = _tk.Tk()
        root.withdraw()
        _mb.showerror("Canvas Course Manager", message)
        root.destroy()
    except Exception:
        print(message)
        try:
            input("Press Enter to close this window...")
        except Exception:
            pass


if sys.version_info < (3, 10):
    _fatal_popup(
        "Canvas Course Manager needs Python 3.10 or newer.\n\n"
        f"This computer is using Python {sys.version_info.major}.{sys.version_info.minor}.\n\n"
        "Download the newest Python from python.org/downloads, install it, "
        "and open the program again."
    )
    raise SystemExit(1)

import copy
import difflib
import hashlib
import json
import math
import mimetypes
import os
import queue
import re
import subprocess
import threading
import time
import traceback
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable
from zoneinfo import ZoneInfo

# Emojis in progress messages must never crash the program when it runs in an old
# Windows console (UnicodeEncodeError). The .pyw launcher has no console at all.
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass

try:
    import pandas as pd
except ImportError:  # Uploading still works; the spreadsheet step asks to install it.
    pd = None  # type: ignore[assignment]

try:
    import openpyxl
except ImportError:  # Needed only to make a new settings spreadsheet.
    openpyxl = None  # type: ignore[assignment]

try:
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, simpledialog, ttk
except ImportError:
    tk = None  # type: ignore[assignment]

try:  # The modern "Sun Valley" look. Optional: without it the window uses the plain system look.
    import sv_ttk
except ImportError:
    sv_ttk = None  # type: ignore[assignment]


# ========================= WHERE THINGS LIVE =========================
# Everything is found relative to this file, so the folder can be copied to
# any Windows or Mac computer without editing a single path.
APP_TITLE = "Canvas Course Manager"
APP_DIR = Path(__file__).resolve().parent
CREDENTIALS_PATH = APP_DIR / "canvas_config.txt"            # Private Canvas login, next to the program.
PROGRAM_SETTINGS_PATH = APP_DIR / "program_settings.json"    # Recent course folders on this computer.
COURSE_SETTINGS_NAME = "course_settings.json"                # One per course folder.
OLD_SETTINGS_PATH = APP_DIR / COURSE_SETTINGS_NAME           # Where older versions kept everything.
MAX_RECENT_COURSES = 8

# ========================= VERSION AND UPDATES =========================
# Raise __version__ for every release (see RELEASING.md) and add a CHANGELOG.md entry.
__version__ = "1.3.1"
PROGRAM_FILE = Path(__file__).resolve()
PROGRAM_FILE_NAME = "canvas_course_manager.py"   # Name of this file in the GitHub repository.
# The public GitHub repository updates come from, as "owner/repository".
# Leave it empty to turn update checks off.
GITHUB_REPO = "odysseyCTE-hub/Canvas-Course-Manager"
UPDATE_STATUS_PATH = APP_DIR / ".update_status.json"
BACKUP_DIR = APP_DIR / "backups"
BACKUPS_TO_KEEP = 5

# ========================= DEFAULT COURSE SETUP =========================
# These defaults are the Fundamentals of Computing conventions. Colleagues change
# them in the "Names & look" tab; the choices are saved in course_settings.json
# inside each course folder.
DEFAULT_SETTINGS: dict[str, Any] = {
    "course_id": "",
    "course_folder": "",
    "spreadsheet": "",
    "plan_module_name": "Lesson Plans",
    "create_missing_modules": True,
    "publish_new_pages": False,
    "capitalize_titles": True,
    "use_emojis": True,
    # Off: items already in Canvas keep whatever emoji they have; only titles that change
    # in other ways are fixed. New course folders start with this off so a colleague's
    # hand-picked emojis are never replaced by the standard ones.
    "fix_existing_emojis": False,
    # Presentations: PowerPoints in lesson folders are shown in Canvas ("canvas"), converted to
    # Google Slides ("google"), or ignored ("off"). Links in a lesson folder always work.
    "slides_mode": "canvas",
    "drive_folder": "",          # Optional Google Drive course folder link.
    # Words are used two ways: to recognize folders/files AND to build Canvas titles.
    "words": {
        "unit": "Unit",
        "lesson": "Lesson",
        "checkpoint": "Checkpoint",
        "activity": "Activity",
        "exam": "Exam",
        "slides": "Slides",
    },
    "emojis": {
        "lesson_header": "",
        "page": "📚",
        "slides": "🎞️",
        "checkpoint": "📝",
        "activity_header": "🪛",
        "activity": "📝",
        "exam_header": "⭐",
        "exam": "⭐",
        "plan": "📄",
    },
    # Canvas allows indentation levels 0 (none) to 5.
    "indents": {
        "lesson_header": 0,
        "page": 1,
        "slides": 1,
        "checkpoint": 2,
        "checkpoint_key": 3,
        "assessment_header": 0,
        "assessment": 1,
        "assessment_key": 2,
        "plan": 0,
    },
    # Gradebook (assignment) group for NEW quizzes of each kind, chosen on the Course setup tab.
    # Each choice is None (not chosen yet), {"leave": True} (stay where Canvas puts them), or
    # {"id": "123", "name": "Checkpoints"}. Saved with the course ID, because group IDs only
    # mean something inside one Canvas course.
    "assignment_groups": {"course_id": "", "checkpoint": None, "activity": None, "exam": None},
    # Remembered between runs so the Run tab opens where you left off.
    "last_task": "lesson",
    "last_scope": "lesson",
    "last_unit": "",
    "last_lesson": "",
}
LOOK_KEYS = ("words", "emojis", "indents", "use_emojis", "capitalize_titles", "plan_module_name")

# Labels shown in the window: (key, label, examples).
WORD_FIELDS = [
    ("unit", "Unit folders and modules", "Unit, Module, Chapter"),
    ("lesson", "Lesson folders", "Lesson, Day, Week"),
    ("checkpoint", "Lesson quiz", "Checkpoint, Quiz, Exit Ticket"),
    ("activity", "Unit activity", "Activity, Project, Lab, Daily"),
    ("exam", "Unit exam", "Exam, Test, Assessment"),
    ("slides", "Presentations", "Slides, Presentation, Deck"),
]
EMOJI_FIELDS = [
    ("lesson_header", "Lesson heading"),
    ("page", "HTML lesson page"),
    ("slides", "Presentation"),
    ("checkpoint", "Lesson quiz"),
    ("activity_header", "Activity heading"),
    ("activity", "Activity quiz"),
    ("exam_header", "Exam heading"),
    ("exam", "Exam quiz"),
    ("plan", "Lesson plan file"),
]
INDENT_FIELDS = [
    ("lesson_header", "Lesson heading"),
    ("page", "HTML lesson page"),
    ("slides", "Presentation"),
    ("checkpoint", "Lesson quiz"),
    ("checkpoint_key", "Lesson quiz answer key"),
    ("assessment_header", "Activity / exam heading"),
    ("assessment", "Activity / exam quiz"),
    ("assessment_key", "Activity / exam answer key"),
    ("plan", "Lesson plan file"),
]
# The "More emojis" window. Grouped so they're easy to find; any emoji can also be pasted.
EMOJI_LIBRARY: list[tuple[str, str]] = [
    ("School", "📚 📖 📘 📗 📙 📕 📓 📒 📔 📝 ✏️ ✒️ 🖊️ 🖍️ 📏 📐 🎒 🏫 🎓 🗒️ 📋 🧾 🔤 🔡"),
    ("Tech", "💻 🖥️ ⌨️ 🖱️ 🖨️ 💾 💿 📀 🔌 🔋 📱 📡 🛰️ 🤖 🧮 🕹️ 🔐 🔒 🔑 🗝️ 🌐 📶 ⚙️ 🧩"),
    ("Tools & science", "🔧 🔨 🪛 🪚 ⚒️ 🛠️ 🔩 🧰 🧲 🧪 🧫 🧬 🔬 🔭 ⚗️ 🌡️ 💡 🔦 🧯 📈 📉 📊 🧠 🌍"),
    ("Files & time", "📄 📃 📑 🗂️ 📁 📂 🗃️ 🗄️ 📦 🏷️ 📌 📍 📎 🖇️ 🔖 📅 📆 🗓️ ⏰ ⏱️ ⌛ 🎞️ 📽️ 🎬"),
    ("Marks", "⭐ 🌟 ✨ ✅ ☑️ ✔️ ❌ ❗ ❓ ⚠️ 🚩 🏁 🎯 🏆 🥇 🎖️ 🔥 💯 🆕 🔔 📣 💬 💭 👀"),
    ("Shapes", "🔴 🟠 🟡 🟢 🔵 🟣 🟤 ⚫ ⚪ 🟥 🟧 🟨 🟩 🟦 🟪 🟫 ⬛ ⬜ 🔶 🔷 🔸 🔹 🔺 🔻"),
    ("Numbers", "0️⃣ 1️⃣ 2️⃣ 3️⃣ 4️⃣ 5️⃣ 6️⃣ 7️⃣ 8️⃣ 9️⃣ 🔟 #️⃣ *️⃣"),
]
EMOJI_CHOICES = ["(none)", "📚", "🎞️", "🖥️", "📝", "📄", "🪛", "⭐", "📘", "📗", "📖", "✏️", "🧪", "🎯",
                 "✅", "🔧", "💻", "🗂️", "🏁", "🧠", "📌", "🔒"]

# ========================= FIXED BEHAVIOUR =========================
# Safety settings are fixed on purpose and are not offered in the window.
HIDE_ANSWER_KEYS = True            # Answer keys are always locked (hidden from students).
IMPORT_AS_NEW_QUIZZES = True
RESYNC_EXISTING = False
REIMPORT_CHANGED_CHECKPOINTS = False
IFRAME_HEIGHT = "650px"
UPLOAD_FOLDER = "Automated_Uploads"
POLL_SECONDS = 5
MIGRATION_TIMEOUT_SECONDS = 900
NEW_QUIZ_APPEAR_SECONDS = 120

# ========================= CURRENT RUN (set by the window) =========================
COURSE_ID = ""
MAIN_DIR = APP_DIR
SPREADSHEET_PATH = APP_DIR / "settings.xlsx"
PLAN_MODULE_NAME = "Lesson Plans"
STATE_PATH = APP_DIR / ".canvas_lesson_uploader_state.json"
LOG_PATH = APP_DIR / "canvas_lesson_uploader.log"
WORDS: dict[str, str] = dict(DEFAULT_SETTINGS["words"])
EMOJIS: dict[str, str] = dict(DEFAULT_SETTINGS["emojis"])
INDENTS: dict[str, int] = dict(DEFAULT_SETTINGS["indents"])
USE_EMOJIS = True
FIX_EXISTING_EMOJIS = False
CAPITALIZE_TITLES = True
CREATE_MISSING_MODULES = True
PUBLISH_NEW_PAGES = False
SLIDES_MODE = "canvas"
DRIVE_FOLDER = ""
GROUP_CHOICES: dict[str, Any] = {}

SCOPE = "lesson"
UNIT_NAME = ""
LESSON_NAME = ""
RUN_LESSON_UPLOAD = False
RUN_SPREADSHEET_SETTINGS = False
UPLOAD_LESSON_PLANS = False
UPLOAD_LESSON_HEADERS = False
UPLOAD_HTML_LESSONS = False
UPLOAD_SLIDES = False
UPLOAD_CHECKPOINTS = False
UPLOAD_ACTIVITIES = False
UPLOAD_EXAMS = False
UPLOAD_ANSWER_KEYS = False
DRY_RUN = True                     # Preview is always the starting point.
SETTINGS_DRY_RUN = True

BASE_URL = ""
TOKEN = ""
API_ROOT = ""
STATE: dict[str, Any] = {}


class CanvasError(RuntimeError):
    pass


# ----------------------- logging -----------------------

READ_ONLY = False   # True while making a spreadsheet: any request that would change Canvas is refused.
LOG_HOOK: Callable[[str], None] | None = None   # The window sets this to show progress.
COUNTS = {"would": 0, "warn": 0, "skip": 0}


def log(message: str) -> None:
    if "DRY RUN: would" in message:
        COUNTS["would"] += 1
    if "WARNING" in message or message.startswith("ERROR"):
        COUNTS["warn"] += 1
    if "skipped —" in message or "SAFETY SKIP" in message:
        COUNTS["skip"] += 1
    if LOG_HOOK is not None:
        LOG_HOOK(message)
    else:
        print(f"[{time.strftime('%H:%M:%S')}] {message}")
    try:
        with LOG_PATH.open("a", encoding="utf-8") as log_file:
            log_file.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except OSError:
        pass


# ----------------------- saved settings -----------------------

def _merge(defaults: dict[str, Any], saved: dict[str, Any]) -> dict[str, Any]:
    """Defaults filled in with saved values; unknown keys are ignored."""
    merged = copy.deepcopy(defaults)
    for key, value in saved.items():
        if key not in merged:
            continue
        if isinstance(merged[key], dict) and isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if sub_key in merged[key]:
                    merged[key][sub_key] = sub_value
        else:
            merged[key] = value
    for key, value in merged["indents"].items():
        try:
            merged["indents"][key] = max(0, min(5, int(value)))
        except (TypeError, ValueError):
            merged["indents"][key] = DEFAULT_SETTINGS["indents"][key]
    for key, value in merged["words"].items():
        if not str(value or "").strip():
            merged["words"][key] = DEFAULT_SETTINGS["words"][key]
    return merged


def looks_like_course_folder(folder: Path) -> bool:
    try:
        return any(p.is_dir() and norm(p.name).startswith(wkey("unit")) for p in folder.iterdir())
    except OSError:
        return False


def _read_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def _write_json(path: Path, data: dict[str, Any]) -> None:
    """Write safely: a half-written file can never replace a good one."""
    temp = path.with_name(path.name + ".tmp")
    temp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)


def _same_folder(a: str | Path, b: str | Path) -> bool:
    try:
        return Path(a).resolve() == Path(b).resolve()
    except OSError:
        return str(a) == str(b)


# --- program settings (this computer): recent course folders ---

def load_program_settings() -> dict[str, Any]:
    data = _read_json(PROGRAM_SETTINGS_PATH)
    recent = [str(f) for f in data.get("recent_courses", []) if isinstance(f, str) and f.strip()]
    appearance = data.get("appearance") if data.get("appearance") in ("light", "dark") else "light"
    return {"last_course_folder": str(data.get("last_course_folder") or ""), "recent_courses": recent,
            "appearance": appearance}


def save_program_settings(program: dict[str, Any]) -> None:
    _write_json(PROGRAM_SETTINGS_PATH, program)


def remember_course_folder(program: dict[str, Any], folder: str) -> None:
    """Put a course folder at the top of the recent list (max MAX_RECENT_COURSES)."""
    if not folder:
        return
    recent = [f for f in program.get("recent_courses", []) if not _same_folder(f, folder)]
    program["recent_courses"] = [folder, *recent][:MAX_RECENT_COURSES]
    program["last_course_folder"] = folder


def recent_course_folders(program: dict[str, Any]) -> list[str]:
    """Recent course folders that still exist on this computer."""
    return [f for f in program.get("recent_courses", []) if Path(f).is_dir()]


# --- course settings (one file inside each course folder) ---

def find_spreadsheet(folder: Path) -> str:
    """Pick the settings workbook in a course folder: one with 'setting' in its name wins."""
    try:
        sheets = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    except OSError:
        return ""
    preferred = [p for p in sheets if "setting" in p.name.lower()] or sheets
    try:
        preferred.sort(key=lambda p: p.stat().st_mtime, reverse=True)   # Newest spreadsheet first.
    except OSError:
        pass
    return str(preferred[0]) if preferred else ""


def load_course_settings(folder: str) -> tuple[dict[str, Any], bool]:
    """Settings for one course folder, and whether that folder already had saved settings.

    Older versions kept one course_settings.json next to the program. If that file
    describes this course folder, it is used (and saved into the folder next time).
    """
    saved: dict[str, Any] = {}
    found = False
    if folder:
        path = Path(folder) / COURSE_SETTINGS_NAME
        if path.is_file():
            saved, found = _read_json(path), True
        elif OLD_SETTINGS_PATH.is_file():
            old = _read_json(OLD_SETTINGS_PATH)
            if old.get("course_folder") and _same_folder(old["course_folder"], folder):
                saved, found = old, True
    if found and "fix_existing_emojis" not in saved:
        saved["fix_existing_emojis"] = True   # Courses set up before this option keep renaming as before.
    settings = _merge(DEFAULT_SETTINGS, saved)
    settings["course_folder"] = folder
    sheet = str(settings["spreadsheet"] or "").strip()
    if sheet and folder and not Path(sheet).is_absolute():
        sheet = str(Path(folder) / sheet)          # Saved relative so the folder can move.
    if (not sheet or not Path(sheet).is_file()) and folder:
        sheet = find_spreadsheet(Path(folder)) or sheet
    settings["spreadsheet"] = sheet
    return settings, found


def save_course_settings(settings: dict[str, Any]) -> Path:
    """Save one course's choices inside its own course folder."""
    folder_text = str(settings.get("course_folder") or "").strip()
    if not folder_text or not Path(folder_text).is_dir():
        raise CanvasError("Choose your course folder on the Course setup tab first. Its settings are saved inside it.")
    folder = Path(folder_text)
    data = {k: v for k, v in settings.items() if k != "course_folder"}
    sheet = str(data.get("spreadsheet") or "").strip()
    if sheet:
        try:
            data["spreadsheet"] = Path(sheet).resolve().relative_to(folder.resolve()).as_posix()
        except (ValueError, OSError):
            data["spreadsheet"] = sheet            # Outside the course folder: keep the full path.
    path = folder / COURSE_SETTINGS_NAME
    _write_json(path, data)
    return path


def starting_course_folder(program: dict[str, Any]) -> str:
    """Which course folder to open with: the last one used, an older setup, or the program's own folder."""
    last = program.get("last_course_folder", "")
    if last and Path(last).is_dir():
        return last
    old = _read_json(OLD_SETTINGS_PATH).get("course_folder", "") if OLD_SETTINGS_PATH.is_file() else ""
    if old and Path(old).is_dir():
        return str(Path(old))
    for folder in recent_course_folders(program):
        return folder
    if looks_like_course_folder(APP_DIR):
        return str(APP_DIR)
    return ""


def apply_look(settings: dict[str, Any]) -> None:
    """Load naming, emoji, and indentation choices into the uploader."""
    global WORDS, EMOJIS, INDENTS, USE_EMOJIS, FIX_EXISTING_EMOJIS, CAPITALIZE_TITLES, PLAN_MODULE_NAME
    WORDS = {k: str(v).strip() for k, v in settings["words"].items()}
    EMOJIS = {k: clean_emoji(v) for k, v in settings["emojis"].items()}
    INDENTS = {k: int(v) for k, v in settings["indents"].items()}
    USE_EMOJIS = bool(settings["use_emojis"])
    FIX_EXISTING_EMOJIS = bool(settings.get("fix_existing_emojis", False))
    CAPITALIZE_TITLES = bool(settings["capitalize_titles"])
    PLAN_MODULE_NAME = str(settings["plan_module_name"]).strip() or "Lesson Plans"


def apply_settings(settings: dict[str, Any]) -> None:
    """Load every saved choice (course, folders, look) into the uploader."""
    global COURSE_ID, MAIN_DIR, SPREADSHEET_PATH, STATE_PATH, LOG_PATH, STATE
    global CREATE_MISSING_MODULES, PUBLISH_NEW_PAGES, SLIDES_MODE, DRIVE_FOLDER, GROUP_CHOICES
    apply_look(settings)
    COURSE_ID = str(settings["course_id"]).strip()
    MAIN_DIR = Path(settings["course_folder"]) if settings["course_folder"] else APP_DIR
    SPREADSHEET_PATH = Path(settings["spreadsheet"]) if settings["spreadsheet"] else MAIN_DIR / "settings.xlsx"
    # Upload history and the log stay in the course folder, so switching to this
    # program never re-uploads files that the old script already uploaded.
    STATE_PATH = MAIN_DIR / ".canvas_lesson_uploader_state.json"
    LOG_PATH = MAIN_DIR / "canvas_lesson_uploader.log"
    CREATE_MISSING_MODULES = bool(settings["create_missing_modules"])
    PUBLISH_NEW_PAGES = bool(settings["publish_new_pages"])
    SLIDES_MODE = settings["slides_mode"] if settings["slides_mode"] in ("canvas", "google", "off") else "canvas"
    DRIVE_FOLDER = str(settings["drive_folder"] or "").strip()
    GROUP_CHOICES = copy.deepcopy(settings.get("assignment_groups") or {})
    STATE = read_state()


# ----------------------- Canvas connection -----------------------

def credential_paths(settings: dict[str, Any]) -> list[Path]:
    paths = [CREDENTIALS_PATH]
    if settings.get("course_folder"):
        paths.append(Path(settings["course_folder"]) / "canvas_config.txt")
    return paths


def read_credentials(settings: dict[str, Any]) -> tuple[dict[str, str], Path]:
    """Find canvas_config.txt next to the program, or in the course folder (older setups)."""
    for path in credential_paths(settings):
        if path.exists():
            config: dict[str, str] = {}
            try:
                for raw in path.read_text(encoding="utf-8-sig").splitlines():
                    line = raw.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, value = line.split("=", 1)
                        config[key.strip().upper()] = value.strip()
            except OSError:
                continue
            return config, path
    return {}, CREDENTIALS_PATH


def write_credentials(url: str, token: str, path: Path) -> None:
    """Save URL and TOKEN, keeping any other lines already in the file."""
    lines: list[str] = []
    if path.exists():
        lines = [line for line in path.read_text(encoding="utf-8-sig").splitlines()
                 if not re.match(r"^\s*(URL|TOKEN)\s*=", line, re.IGNORECASE)]
    if not lines:
        lines = ["# Canvas connection for Canvas Course Manager.",
                 "# PRIVATE: this token lets the program act as you. Never share this file."]
    lines += [f"URL={url}", f"TOKEN={token}"]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def clean_canvas_url(text: str) -> str:
    """Accept any Canvas link someone pastes and keep just https://school.instructure.com."""
    text = text.strip()
    if not text:
        return ""
    if not re.match(r"^https?://", text, re.IGNORECASE):
        text = "https://" + text
    parsed = urllib.parse.urlparse(text)
    return f"https://{parsed.netloc}" if parsed.netloc else ""


def course_id_from_text(text: str) -> str:
    """Accept a plain number or a pasted course link like .../courses/1075/modules."""
    text = text.strip()
    match = re.search(r"/courses/(\d+)", text)
    if match:
        return match.group(1)
    return text if text.isdigit() else text


def set_connection(url: str, token: str) -> None:
    global BASE_URL, TOKEN, API_ROOT
    BASE_URL = clean_canvas_url(url)
    TOKEN = token.strip()
    API_ROOT = BASE_URL + "/api/v1"


# ----------------------- naming conventions -----------------------

def norm(value: str) -> str:
    """Normalize titles by ignoring emoji, case, punctuation, and separators."""
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def word(key: str) -> str:
    """The word chosen for a convention, e.g. word('checkpoint') -> 'Quiz'."""
    return WORDS.get(key) or DEFAULT_SETTINGS["words"][key]


def wkey(key: str) -> str:
    """The normalized form used to spot the word inside file and folder names."""
    return norm(word(key))


def _word_rx(key: str) -> str:
    return r"[\s_-]*".join(re.escape(part) for part in word(key).split())


def leading_number(name: str, key: str) -> int | None:
    """N from names like 'Unit 4', 'Unit 4 - Data', 'unit4', 'Lesson 32_Topic'."""
    match = re.match(r"^\s*" + _word_rx(key) + r"[\s_-]*(\d+)(?!\d)", name, re.IGNORECASE)
    return int(match.group(1)) if match else None


def is_unit_name(name: str) -> bool:
    return norm(name).startswith(wkey("unit"))


def is_lesson_name(name: str) -> bool:
    return norm(name).startswith(wkey("lesson"))


def natural_key(text: str) -> list[Any]:
    """Sort 'Lesson 2' before 'Lesson 10'."""
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


# Characters that are part of an emoji but are not "symbols" themselves: the
# invisible joiner, the two variation selectors, the keycap mark, and the tag
# characters used by some flags.
_EMOJI_GLUE = {"\u200d", "\ufe0e", "\ufe0f", "\u20e3"}


def _is_emoji_char(ch: str) -> bool:
    if ch in _EMOJI_GLUE or "\U000e0020" <= ch <= "\U000e007f":
        return True
    return unicodedata.category(ch) in ("So", "Sk")   # Symbols, including skin-tone modifiers.


def fix_surrogates(text: str) -> str:
    """Some Tk builds hand back an emoji as two halves; join them into one character."""
    try:
        return text.encode("utf-16", "surrogatepass").decode("utf-16")
    except UnicodeError:
        return text.encode("utf-8", "replace").decode("utf-8")


def clean_emoji(value: Any) -> str:
    """The emoji typed in a box, tidied: '(none)' or blank means no emoji."""
    text = fix_surrogates(str(value or "")).strip()
    return "" if text in ("", "(none)") else text


def emoji_problem(value: Any) -> str:
    """Why an emoji box value can't be used, or '' when it's fine.

    Letters and numbers are refused: the program recognizes items it made by their
    words, so letters at the start of a title would make it add duplicates.
    """
    text = clean_emoji(value)
    if not text:
        return ""
    if any(unicodedata.category(ch)[0] in ("L", "N") and ch not in "0123456789#*" for ch in text):
        return "letters or words"
    if any(ch.isdigit() or ch in "#*" for ch in text) and "\u20e3" not in text:
        return "numbers"
    if len(text) > 12:
        return "too many characters (use one or two emojis)"
    return ""


def strip_leading_emoji(title: str) -> str:
    """Remove every emoji (and the spaces after it) from the start of a title.

    'Lesson 4' -> 'Lesson 4', '📝Lesson 4' -> 'Lesson 4', '🎞️ 📝Lesson 4' -> 'Lesson 4'.
    Keycap emojis such as 1️⃣ are removed too.
    """
    text = fix_surrogates(title)
    while True:
        before = text
        text = text.lstrip()
        if len(text) >= 2 and text[0] in "0123456789#*" and "\u20e3" in text[1:3]:
            text = text[text.index("\u20e3") + 1:]
        while text and _is_emoji_char(text[0]):
            text = text[1:]
        if text == before:
            return text


def emoji_prefix_of(title: str) -> str:
    """The emoji part at the start of a title ('' when there is none)."""
    text = fix_surrogates(title)
    return text[: len(text) - len(strip_leading_emoji(text))].strip()


def emoji(key: str) -> str:
    """The emoji prefix for an item type, or '' when emojis are turned off."""
    return EMOJIS.get(key, "") if USE_EMOJIS else ""


def with_emoji(key: str, title: str) -> str:
    """Put this item type's emoji at the front, replacing any emoji already there.

    Older versions only checked for their own emoji, so a file named '🔧Lab 2' or a
    title holding a different emoji ended up with two ('📝🔧Lab 2').
    """
    prefix = emoji(key)
    if not prefix:
        return title
    return prefix + strip_leading_emoji(title)


def ind(key: str) -> int:
    return int(INDENTS.get(key, DEFAULT_SETTINGS["indents"][key]))


def capitalize_words(text: str) -> str:
    """Capitalize the first letter of every word without lowercasing the rest,
    so "intro to HTML" becomes "Intro To HTML". Can be turned off in Names & look."""
    if not CAPITALIZE_TITLES:
        return text
    return re.sub(r"(^|\s)([a-z])", lambda m: m.group(1) + m.group(2).upper(), text)


# Title builders. The uploader and the live sample in the window both use these,
# so the sample always matches what really appears in Canvas.

def fmt_lesson_title(number: str, name: str) -> str:
    base = f"{word('lesson')} {number}" if number else ""
    if name:
        base = (base + " " + capitalize_words(name)).strip()
    return base


def lesson_header_title(number: str, name: str) -> str:
    return with_emoji("lesson_header", fmt_lesson_title(number, name))


def page_title(number: str, name: str) -> str:
    return with_emoji("page", fmt_lesson_title(number, name))


def checkpoint_title(number: str, name: str) -> str:
    if number:
        title = f"{word('lesson')} {number} {word('checkpoint')}" + (f" - {capitalize_words(name)}" if name else "")
    else:
        title = f"{capitalize_words(name)} {word('checkpoint')}".strip()
    return with_emoji("checkpoint", title)


def checkpoint_key_title(number: str) -> str:
    return f"{word('lesson')} {number} {word('checkpoint')} - Answer Key"


def plan_title(number: str, folder_name: str = "") -> str:
    return with_emoji("plan", f"{word('lesson')} {number} Plan" if number else f"{folder_name} Plan")


def assessment_header_title(unit_label: str, kind: str) -> str:
    return with_emoji(f"{kind}_header", f"{unit_label} {word(kind)}")


def assessment_quiz_title(stem: str, kind: str) -> str:
    """Activities and exams use the package file name as the title."""
    name = re.sub(r"\s+", " ", re.sub(r"[_]+", " ", stem)).strip()
    return with_emoji(kind, name)


def assessment_key_title(unit_label: str, kind: str) -> str:
    return f"{unit_label} {word(kind)} - Answer Key"


# ----------------------- Canvas API (unchanged) -----------------------

def api_request(method: str, path_or_url: str, *, json_body: Any = None,
                form_body: dict[str, Any] | None = None, allow_404: bool = False) -> Any:
    if not API_ROOT:
        raise CanvasError("Not connected to Canvas yet. Fill in the Course setup tab first.")
    if READ_ONLY and method.upper() != "GET":
        raise CanvasError(f"Safety stop: tried to change Canvas ({method.upper()}) while only reading the course.")
    url = path_or_url if path_or_url.startswith("https://") else API_ROOT + path_or_url
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    body = None
    if json_body is not None:
        body = json.dumps(json_body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    elif form_body is not None:
        body = urllib.parse.urlencode(form_body).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    request = urllib.request.Request(url, data=body, headers=headers, method=method.upper())
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            raw = response.read()
            if not raw:
                return {}
            try:
                return json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                return {"raw_response": raw.decode("utf-8", errors="replace")}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if allow_404 and exc.code == 404:
            return None
        raise CanvasError(f"{method} {url} failed ({exc.code}): {detail[:1200]}") from exc
    except urllib.error.URLError as exc:
        raise CanvasError(f"Network error calling {url}: {exc}") from exc


def get_all_pages(path: str) -> list[dict[str, Any]]:
    url = API_ROOT + path
    results: list[dict[str, Any]] = []
    while url:
        request = urllib.request.Request(url, headers={"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=90) as response:
                data = json.loads(response.read().decode("utf-8"))
                if isinstance(data, list):
                    results.extend(item for item in data if isinstance(item, dict))
                link_header = response.headers.get("Link", "")
        except (urllib.error.HTTPError, urllib.error.URLError, json.JSONDecodeError) as exc:
            raise CanvasError(f"Could not read paginated endpoint {url}: {exc}") from exc
        url = ""
        for part in link_header.split(","):
            if 'rel="next"' in part:
                match = re.search(r"<([^>]+)>", part)
                if match:
                    url = match.group(1)
                    break
    return results


def read_state() -> dict[str, Any]:
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError):
        log("WARNING: could not read state file; starting with empty state.")
        return {}


def save_state() -> None:
    if DRY_RUN:
        return
    temp_path = STATE_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(STATE, indent=2), encoding="utf-8")
    temp_path.replace(STATE_PATH)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _match_score(wanted: Any, candidate: Any) -> float:
    """Conservative title similarity score for legacy Canvas naming variations."""
    wanted_norm = norm(_cell_text(wanted))
    candidate_norm = norm(_cell_text(candidate))
    if not wanted_norm or not candidate_norm:
        return 0.0
    if wanted_norm == candidate_norm:
        return 1.0
    wanted_list = wanted_norm.split()
    candidate_list = candidate_norm.split()
    wanted_tokens = set(wanted_list)
    candidate_tokens = set(candidate_list)
    # A shorter Canvas title may be an intentional prefix of the spreadsheet title,
    # e.g. "Lesson 17 Computing Threats" vs. "Lesson 17 Computing Threats Malware Phishing Ransomware".
    # Require a meaningful prefix (at least three tokens) to avoid matching generic titles.
    shorter, longer = (candidate_list, wanted_list) if len(candidate_list) < len(wanted_list) else (wanted_list, candidate_list)
    if len(shorter) >= 3 and longer[:len(shorter)] == shorter:
        return 0.91 + min(0.07, len(shorter) / max(len(longer), 1) * 0.07)
    # Also allow a shorter expected title to match when it is fully contained in the Canvas title.
    if len(wanted_tokens) >= 2 and wanted_tokens.issubset(candidate_tokens):
        return 0.90 + min(0.09, len(wanted_tokens) / max(len(candidate_tokens), 1) * 0.09)
    overlap = len(wanted_tokens & candidate_tokens) / max(len(wanted_tokens | candidate_tokens), 1)
    sequence = difflib.SequenceMatcher(None, wanted_norm, candidate_norm).ratio()
    return max(sequence, overlap)


def _unique_best_match(wanted: Any, candidates: list[dict[str, Any]], name_getter,
                       *, threshold: float = 0.82, margin: float = 0.08) -> tuple[dict[str, Any] | None, str]:
    """Return a unique strong fuzzy match, never guessing between near-ties."""
    scored = sorted(
        ((_match_score(wanted, name_getter(candidate)), candidate) for candidate in candidates),
        key=lambda pair: pair[0], reverse=True,
    )
    if not scored or scored[0][0] < threshold:
        return None, "no confident match"
    if len(scored) > 1 and scored[0][0] - scored[1][0] < margin:
        return None, "ambiguous match"
    return scored[0][1], f"fuzzy match score {scored[0][0]:.2f}"



# ----------------------- change tracking (unchanged) -----------------------

def file_state_key(path: Path) -> str:
    return f"file:{COURSE_ID}:{path.relative_to(MAIN_DIR).as_posix()}"


def sync_status(path: Path) -> tuple[str, str]:
    """Return (digest, status): "new" if never recorded, otherwise "same" or "changed"."""
    digest = sha256_file(path)
    saved = STATE.get(file_state_key(path), {}).get("sha256")
    if saved is None:
        return digest, "new"
    return digest, ("same" if saved == digest else "changed")


def remember_upload(path: Path, digest: str, file_id: Any) -> None:
    if DRY_RUN:
        return
    STATE[file_state_key(path)] = {
        "sha256": digest,
        "file_id": file_id,
        "recorded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_state()


def upload_multipart(upload_url: str, upload_params: dict[str, Any], file_path: Path) -> dict[str, Any]:
    """Post Canvas-provided upload parameters and file bytes to the signed upload URL."""
    boundary = "----CanvasUpload" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for key, value in upload_params.items():
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
        chunks.append(str(value).encode("utf-8") + b"\r\n")
    mime_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
    chunks.append(f"--{boundary}\r\n".encode())
    chunks.append(f'Content-Disposition: form-data; name="file"; filename="{file_path.name}"\r\n'.encode())
    chunks.append(f"Content-Type: {mime_type}\r\n\r\n".encode())
    chunks.append(file_path.read_bytes())
    chunks.append(b"\r\n" + f"--{boundary}--\r\n".encode())

    # No Canvas bearer token is sent to the signed storage URL.
    request = urllib.request.Request(
        upload_url,
        data=b"".join(chunks),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            raw = response.read()
            payload: Any = {}
            if raw:
                try:
                    payload = json.loads(raw.decode("utf-8"))
                except json.JSONDecodeError:
                    payload = {"raw_response": raw.decode("utf-8", errors="replace")}
            return {"status": response.status, "payload": payload, "location": response.headers.get("Location")}
    except urllib.error.HTTPError as exc:
        # Some upload endpoints respond with a redirect to Canvas's completion endpoint.
        if exc.code in (301, 302, 303, 307, 308):
            return {"status": exc.code, "payload": {}, "location": exc.headers.get("Location")}
        detail = exc.read().decode("utf-8", errors="replace")
        raise CanvasError(f"File transfer failed for {file_path.name} ({exc.code}): {detail[:1000]}") from exc
    except urllib.error.URLError as exc:
        raise CanvasError(f"File transfer failed for {file_path.name}: {exc}") from exc


def finish_upload(result: dict[str, Any]) -> dict[str, Any]:
    payload = result.get("payload")
    if isinstance(payload, dict) and payload.get("id"):
        return payload
    location = result.get("location")
    if location:
        if location.startswith("/"):
            location = BASE_URL + location
        if not location.startswith(BASE_URL + "/"):
            raise CanvasError("Upload completion URL is not on the configured Canvas host; refusing to send token.")
        completed = api_request("GET", location)
        if isinstance(completed, dict) and completed.get("id"):
            return completed
    raise CanvasError(f"Upload transfer finished but Canvas did not confirm a file ID. Response: {result}")


def upload_course_file(path: Path, folder_path: str) -> dict[str, Any]:
    if DRY_RUN:
        log(f"DRY RUN: would upload {path.name} to {folder_path}")
        return {"id": None, "display_name": path.name}
    init = api_request("POST", f"/courses/{COURSE_ID}/files", form_body={
        "name": path.name,
        "size": path.stat().st_size,
        "content_type": mimetypes.guess_type(str(path))[0] or "application/octet-stream",
        "parent_folder_path": folder_path,
        "on_duplicate": "overwrite",
    })
    if not isinstance(init, dict) or not init.get("upload_url") or not isinstance(init.get("upload_params"), dict):
        raise CanvasError(f"Canvas did not return upload instructions for {path.name}: {init}")
    transfer = upload_multipart(init["upload_url"], init["upload_params"], path)
    return finish_upload(transfer)


def hide_file(file_id: Any, name: str) -> None:
    """Unpublish a course file so only teachers can see it, then confirm Canvas applied it."""
    if not file_id:
        if DRY_RUN:
            log(f"DRY RUN: would hide {name} from students after uploading it")
        return
    current = api_request("GET", f"/files/{file_id}", allow_404=True)
    if isinstance(current, dict) and current.get("locked"):
        log(f"Already hidden from students: {name}")
        return
    if DRY_RUN:
        log(f"DRY RUN: would hide {name} from students (it is currently visible to them)")
        return
    api_request("PUT", f"/files/{file_id}", form_body={"locked": "true"})
    check = api_request("GET", f"/files/{file_id}")
    if not (isinstance(check, dict) and check.get("locked")):
        raise CanvasError(f"Could not confirm {name} (file {file_id}) is hidden from students. Check it in Canvas Files.")
    log(f"Hidden from students: {name}")


def find_module_item(module_id: str, item_type: str, title: str) -> dict[str, Any] | None:
    for item in module_items(module_id):
        if item.get("type", "").lower() == item_type.lower() and norm(item.get("title", "")) == norm(title):
            return item
    return None


def add_module_item(module_id: str, item_type: str, title: str, *, indent: int = 0,
                    content_id: int | str | None = None, page_url: str | None = None,
                    position: int | None = None, extra: dict[str, Any] | None = None) -> dict[str, Any] | None:
    existing = find_module_item(module_id, item_type, title)
    if existing:
        log(f"Already in module; skipping duplicate: {item_type} — {title}")
        return existing
    item: dict[str, Any] = {"type": item_type, "title": title, "indent": indent}
    if content_id is not None:
        item["content_id"] = content_id
    if page_url:
        item["page_url"] = page_url
    if position is not None:
        item["position"] = position
    item.update(extra or {})
    if DRY_RUN:
        log(f"DRY RUN: would add {item_type} '{title}' to module '{_MODULE_NAMES.get(str(module_id), module_id)}', indent {indent}")
        return None
    return api_request("POST", f"/courses/{COURSE_ID}/modules/{module_id}/items", json_body={"module_item": item})


def ensure_module_item_title(module_id: str, item: dict[str, Any], title: str) -> None:
    """Rename a module item whose title differs only in capitalization or emoji.

    When only the emoji differs, the item is renamed only if "Update emojis on items
    already in Canvas" is ticked, so emojis someone picked by hand are left alone.
    """
    current = str(item.get("title") or "")
    if current == title or not item.get("id"):
        return
    if strip_leading_emoji(current) == strip_leading_emoji(title) and not FIX_EXISTING_EMOJIS:
        log(f"Kept the emoji already in Canvas: '{current}' (tick \"Update emojis on items already in "
            f"Canvas\" on the Names & look tab to change it to '{emoji_prefix_of(title) or 'no emoji'}')")
        return
    if DRY_RUN:
        log(f"DRY RUN: would rename module item '{item.get('title')}' to '{title}'")
        return
    api_request("PUT", f"/courses/{COURSE_ID}/modules/{module_id}/items/{item['id']}",
                json_body={"module_item": {"title": title}})
    log(f"Renamed module item '{item.get('title')}' to '{title}'")


def replace_module_item(module_id: str, item: dict[str, Any], item_type: str,
                        content_id: Any, title: str, indent: int) -> None:
    """Point a module entry at new content: add the new item in the same spot, then remove the old one."""
    if str(item.get("content_id")) == str(content_id):
        return
    new_item = {"type": item_type, "title": title, "indent": indent, "content_id": content_id}
    if item.get("position") is not None:
        new_item["position"] = item["position"]
    api_request("POST", f"/courses/{COURSE_ID}/modules/{module_id}/items", json_body={"module_item": new_item})
    api_request("DELETE", f"/courses/{COURSE_ID}/modules/{module_id}/items/{item['id']}")
    log(f"Module item now points to the new version: {title}")


def sync_file_item(path: Path, module_id: str, title: str, indent: int,
                   existing: dict[str, Any] | None, label: str, hide: bool = False) -> None:
    """Upload a file and add it to a module, or re-upload it if it changed on this computer."""
    digest, status = sync_status(path)
    if existing and not RESYNC_EXISTING and status in {"new", "same"}:
        if status == "new":
            log(f"{label} already in module; recording current version so future changes upload: {title}")
            remember_upload(path, digest, existing.get("content_id"))
        else:
            log(f"{label} unchanged; skipping: {title}")
        if hide:
            hide_file(existing.get("content_id"), path.name)
        return

    log(f"{'Re-uploading changed' if existing else 'Uploading'} {label.lower()}: {path.name}")
    uploaded = upload_course_file(path, upload_folder_for(path))
    if DRY_RUN:
        if hide:
            hide_file(None, path.name)
        if existing:
            log(f"DRY RUN: would point module item '{title}' at the new upload")
        else:
            add_module_item(module_id, "File", title, indent=indent)
        return
    file_id = uploaded.get("id")
    if not file_id:
        raise CanvasError(f"Canvas did not return a file ID for {path.name}")
    if hide:
        hide_file(file_id, path.name)
    if existing:
        replace_module_item(module_id, existing, "File", file_id, title, indent)
    else:
        add_module_item(module_id, "File", title, content_id=file_id, indent=indent)
    remember_upload(path, digest, file_id)


def upload_folder_for(path: Path) -> str:
    unit_name = path.parent.parent.name
    lesson_name = path.parent.name
    return f"{UPLOAD_FOLDER}/{unit_name}/{lesson_name}"


def page_list() -> list[dict[str, Any]]:
    return get_all_pages(f"/courses/{COURSE_ID}/pages?per_page=100")


def import_qti(path: Path) -> dict[str, Any]:
    # Canvas's content-migration endpoint uses pre_attachment to upload the QTI ZIP.
    form = {
        "migration_type": "qti_converter",
        "pre_attachment[name]": path.name,
        "pre_attachment[size]": path.stat().st_size,
        "pre_attachment[content_type]": "application/zip",
    }
    if IMPORT_AS_NEW_QUIZZES:
        form["settings[import_quizzes_next]"] = "true"
    # We add the quiz explicitly after import, so we do not rely on undocumented placement settings.
    migration = api_request("POST", f"/courses/{COURSE_ID}/content_migrations", form_body=form)
    if not isinstance(migration, dict) or not migration.get("id"):
        raise CanvasError(f"Canvas did not return a migration ID for {path.name}: {migration}")
    pre = migration.get("pre_attachment") or {}
    if not pre.get("upload_url") or not isinstance(pre.get("upload_params"), dict):
        raise CanvasError(f"Canvas did not return QTI upload instructions for {path.name}: {migration}")
    upload_result = upload_multipart(pre["upload_url"], pre["upload_params"], path)
    # Canvas usually processes the migration after receiving the pre-attachment upload.
    if upload_result.get("location"):
        finish_upload(upload_result)

    migration_id = migration["id"]
    deadline = time.time() + MIGRATION_TIMEOUT_SECONDS
    previous_state = None
    while time.time() < deadline:
        current = api_request("GET", f"/courses/{COURSE_ID}/content_migrations/{migration_id}")
        state = current.get("workflow_state", "unknown") if isinstance(current, dict) else "unknown"
        if state != previous_state:
            log(f"QTI migration {migration_id}: {state}")
            previous_state = state
        if state == "completed":
            return current
        if state in {"failed", "waiting_for_select"}:
            issues = api_request("GET", f"/courses/{COURSE_ID}/content_migrations/{migration_id}/migration_issues", allow_404=True)
            raise CanvasError(f"QTI migration {migration_id} ended as '{state}'. Issues: {issues}")
        time.sleep(POLL_SECONDS)
    raise CanvasError(f"Timed out waiting for QTI migration {migration_id}. Check Canvas before rerunning.")


def assignment_name(assignment: dict[str, Any]) -> str:
    return assignment.get("name", assignment.get("title", "")) or "(untitled)"


def create_embedded_page(html_path: Path, lesson_folder: Path, uploaded: dict[str, Any],
                         existing_url: str | None = None) -> dict[str, Any]:
    """Create the Canvas page that embeds the HTML file, or point an existing page at the new upload."""
    title = page_title(*lesson_details(lesson_folder, html_path))
    slug = re.sub(r"-+", "-", re.sub(r"[^a-z0-9_-]+", "-", title.lower())).strip("-")[:200] or "lesson-page"
    if DRY_RUN:
        log(f"DRY RUN: would create/update Canvas page '{title}' embedding {html_path.name}")
        return {"title": title, "url": existing_url or slug}
    file_id = uploaded.get("id")
    if not file_id:
        raise CanvasError(f"No Canvas file ID was returned for {html_path.name}")
    file_url = f"{BASE_URL}/courses/{COURSE_ID}/files/{file_id}/download"
    body = (
        '<p><iframe src="' + file_url + '" width="100%" height="' + IFRAME_HEIGHT + '" '
        'style="border: none; border-radius: 8px;" loading="lazy"></iframe></p>'
    )
    existing = next((p for p in page_list() if p.get("url") in {existing_url, slug} or norm(p.get("title", "")) == norm(title)), None)
    if existing:
        page = api_request("PUT", f"/courses/{COURSE_ID}/pages/{urllib.parse.quote(existing['url'], safe='')}",
                           json_body={"wiki_page": {"title": title, "body": body}})
        log(f"Updated existing page: {title}")
        return page
    page = api_request("POST", f"/courses/{COURSE_ID}/pages", json_body={"wiki_page": {
        "title": title, "url": slug, "body": body, "editing_roles": "teachers", "published": PUBLISH_NEW_PAGES
    }})
    log(f"Created Canvas page: {title}")
    return page


def ensure_page(path: Path, lesson_folder: Path) -> None:
    material_module_id, material_module_name = get_material_module(lesson_folder)
    log(f"Using Canvas module: {material_module_name} (ID {material_module_id})")
    title = page_title(*lesson_details(lesson_folder, path))
    existing = find_module_item(material_module_id, "Page", title)
    digest, status = sync_status(path)
    if existing:
        ensure_module_item_title(material_module_id, existing, title)
        if not RESYNC_EXISTING and status == "new":
            log(f"Lesson page already in module; recording current version so future changes upload: {title}")
            remember_upload(path, digest, None)
            return
        if not RESYNC_EXISTING and status == "same":
            log(f"Lesson page unchanged; skipping: {title}")
            return
        log(f"Re-uploading changed HTML lesson: {path.name}")
    uploaded = upload_course_file(path, upload_folder_for(path))
    page = create_embedded_page(path, lesson_folder, uploaded, existing.get("page_url") if existing else None)
    if not existing:
        if not page.get("url"):
            raise CanvasError(f"Canvas did not return a page URL for {title}")
        add_module_item(material_module_id, "Page", title, page_url=page["url"], indent=ind("page"))
    remember_upload(path, digest, uploaded.get("id"))


def ensure_quiz(path: Path, lesson_folder: Path, html_path: Path | None = None) -> None:
    """Import a QTI package and add its New Quiz as an Assignment module item.

    Canvas New Quizzes are LTI assignments. Their IDs are found through the
    course assignments endpoint, not the legacy /quizzes endpoint.
    """
    material_module_id, material_module_name = get_material_module(lesson_folder)
    log(f"Using Canvas module: {material_module_name} (ID {material_module_id})")
    key = f"{COURSE_ID}:{lesson_folder.as_posix()}:{path.name}"
    digest = sha256_file(path)
    saved = STATE.get(key, {})
    number, name = lesson_details(lesson_folder, html_path)
    title_kind = assessment_kind(path, lesson_folder)
    if title_kind == "checkpoint":
        # Lesson quizzes are titled from the lesson, e.g. 📝Lesson 32 Checkpoint - Topic.
        quiz_title = checkpoint_title(number, name or ("" if number else lesson_title(lesson_folder, html_path)))
    else:
        # Activities and exams use the package file name as the title.
        quiz_title = assessment_quiz_title(path.stem, title_kind)
    desired_title = norm(quiz_title)

    existing_item = find_module_item(material_module_id, "Assignment", quiz_title)
    changed = False
    if existing_item:
        ensure_module_item_title(material_module_id, existing_item, quiz_title)
        # A module-item title and the underlying New Quiz assignment title are
        # separate Canvas fields. Keep both aligned with the agreed naming rule.
        existing_assignment_id = existing_item.get("content_id")
        if existing_assignment_id:
            existing_assignment = api_request("GET", f"/courses/{COURSE_ID}/assignments/{existing_assignment_id}", allow_404=True)
            if isinstance(existing_assignment, dict) and existing_assignment.get("id"):
                changes: dict[str, Any] = {}
                if assignment_name(existing_assignment) != quiz_title:
                    log(f"Renaming existing Canvas assignment {existing_assignment_id} from '{assignment_name(existing_assignment)}' to '{quiz_title}'")
                    changes["name"] = quiz_title
                # Already-uploaded quizzes only move while they are still in "Imported Assignments".
                changes.update(planned_group_change(existing_assignment, title_kind, quiz_title, new_import=False))
                if changes and not DRY_RUN:
                    updated = api_request("PUT", f"/courses/{COURSE_ID}/assignments/{existing_assignment_id}",
                                          json_body={"assignment": changes})
                    confirm_group_change(updated, changes, quiz_title)
        if not saved.get("sha256"):
            log(f"Checkpoint already in module; recording current version so future changes are detected: {quiz_title}")
            if not DRY_RUN:
                STATE[key] = {"sha256": digest, "workflow_state": "completed",
                              "assignment_id": existing_item.get("content_id")}
                save_state()
            return
        if saved["sha256"] == digest:
            log(f"Checkpoint unchanged; skipping: {quiz_title}")
            return
        if not REIMPORT_CHANGED_CHECKPOINTS:
            log(f"WARNING: {path.name} changed since it was imported, but it was NOT re-imported because "
                f"REIMPORT_CHANGED_CHECKPOINTS is False. Edit the quiz in Canvas, or set it to True to import a replacement.")
            return
        changed = True

    if DRY_RUN:
        target = GROUP_TARGETS.get(title_kind)
        group_note = f", in gradebook group '{target.get('name')}'" if target else ""
        if changed:
            log(f"DRY RUN: would import the changed {path.name} as a new quiz{group_note}, replace '{quiz_title}' "
                f"in the module with it, and leave the old quiz in Assignments")
        else:
            log(f"DRY RUN: would find/import New Quiz for {path.name} and add it as an Assignment module item "
                f"titled '{quiz_title}'{group_note}")
        return

    assignment = None
    imported_now = False
    if changed or saved.get("sha256") != digest or saved.get("workflow_state") != "completed":
        imported_now = True
        before_ids = {str(a.get("id")) for a in get_all_pages(f"/courses/{COURSE_ID}/assignments?per_page=100")}
        migration = import_qti(path)
        STATE[key] = {
            "sha256": digest,
            "migration_id": migration.get("id"),
            "workflow_state": "completed",
            "completed_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "previous_assignment_id": saved.get("assignment_id"),
        }
        save_state()
        # A New Quiz can appear in Assignments a little after the migration reports "completed".
        deadline = time.time() + NEW_QUIZ_APPEAR_SECONDS
        while True:
            assignments = get_all_pages(f"/courses/{COURSE_ID}/assignments?per_page=100")
            newly_created = [a for a in assignments if a.get("id") is not None and str(a.get("id")) not in before_ids]
            if newly_created or time.time() >= deadline:
                break
            time.sleep(POLL_SECONDS)
        # Only consider quizzes created by this import, so an older copy with the same title is never picked.
        assignment = next((a for a in newly_created if norm(assignment_name(a)) == desired_title), None)
        if assignment is None and len(newly_created) == 1:
            assignment = newly_created[0]
        if assignment is not None:
            log(f"Identified New Quiz assignment created by this QTI import: {assignment.get('id')} ({assignment_name(assignment)})")
    else:
        assignments = get_all_pages(f"/courses/{COURSE_ID}/assignments?per_page=100")
        # Reuse a previously saved assignment ID, if it still exists.
        if saved.get("assignment_id"):
            found = api_request("GET", f"/courses/{COURSE_ID}/assignments/{saved['assignment_id']}", allow_404=True)
            if isinstance(found, dict) and found.get("id"):
                assignment = found
        previous_id = str(saved.get("previous_assignment_id"))
        candidates = [a for a in assignments if str(a.get("id")) != previous_id]
        if assignment is None:
            assignment = next((a for a in candidates if norm(assignment_name(a)) == desired_title), None)
        # Recovery path for an import completed by a previous run that stopped before module placement.
        if assignment is None and saved.get("completed_at"):
            try:
                completed_epoch = time.mktime(time.strptime(saved["completed_at"], "%Y-%m-%d %H:%M:%S"))
                recent = []
                for a in candidates:
                    created_at = a.get("created_at")
                    if not a.get("id") or not created_at:
                        continue
                    try:
                        created_epoch = time.mktime(time.strptime(created_at[:19], "%Y-%m-%dT%H:%M:%S"))
                    except (ValueError, TypeError):
                        continue
                    title = norm(assignment_name(a))
                    if created_epoch >= completed_epoch - 300 and ((number and f"{wkey('lesson')} {number}" in title) or wkey("checkpoint") in title):
                        recent.append(a)
                if len(recent) == 1:
                    assignment = recent[0]
            except (ValueError, TypeError, OSError):
                pass

    if assignment is None or not assignment.get("id"):
        available = "; ".join(f"{a.get('id')}: {assignment_name(a)}" for a in assignments)
        raise CanvasError(
            f"QTI import is complete for {path.name}, but Canvas did not return a safely identifiable New Quiz assignment. "
            f"Available course assignments: {available or '(none returned)'}. Run the program again in a few minutes; "
            f"it will not re-import the ZIP while the import is marked complete."
        )

    assignment_id = assignment["id"]
    # Rename the actual Canvas assignment as well as the module item. QTI imports
    # can preserve the title embedded in the ZIP, which may omit the agreed emoji
    # or checkpoint naming convention.
    # The gradebook group is set in the same request, so a new quiz never needs a second change.
    current_assignment_title = assignment_name(assignment)
    changes = {}
    if current_assignment_title != quiz_title:
        log(f"Renaming Canvas assignment {assignment_id} from '{current_assignment_title}' to '{quiz_title}'")
        changes["name"] = quiz_title
    changes.update(planned_group_change(assignment, title_kind, quiz_title, new_import=imported_now))
    if changes and not DRY_RUN:
        renamed = api_request(
            "PUT", f"/courses/{COURSE_ID}/assignments/{assignment_id}",
            json_body={"assignment": changes},
        )
        if isinstance(renamed, dict) and renamed.get("id"):
            assignment = renamed
            if "name" in changes and assignment_name(assignment) != quiz_title:
                raise CanvasError(f"Canvas did not confirm the expected assignment title '{quiz_title}' for assignment {assignment_id}.")
        confirm_group_change(renamed, changes, quiz_title)
    STATE[key] = {
        **STATE.get(key, {}),
        "sha256": digest,
        "workflow_state": "completed",
        "assignment_id": assignment_id,
        "canvas_assignment_title": assignment_name(assignment),
    }
    save_state()

    # A New Quiz belongs in a module as an Assignment item, not a legacy Quiz item.
    # Activities and exams sit one level below their header; checkpoints keep
    # their established indentation.
    kind = assessment_kind(path, lesson_folder)
    quiz_indent = ind("assessment") if kind in {"activity", "exam"} else ind("checkpoint")
    existing_item = find_module_item(material_module_id, "Assignment", quiz_title)
    if existing_item:
        old_id = existing_item.get("content_id")
        replace_module_item(material_module_id, existing_item, "Assignment", assignment_id, quiz_title, quiz_indent)
        if str(old_id) != str(assignment_id):
            log(f"NOTE: the old version of this checkpoint (assignment {old_id}) is still in Assignments. "
                f"Check it for student attempts, then delete it in Canvas if it is no longer needed.")
    else:
        add_module_item(material_module_id, "Assignment", quiz_title, content_id=assignment_id, indent=quiz_indent)
    log(f"Linked New Quiz assignment {assignment_id} ('{assignment_name(assignment)}') in module as: {quiz_title}")




# ----------------------- modules -----------------------

_PLANNED_PREFIX = "planned-module:"
_UNIT_MODULE_CACHE: dict[int, tuple[str, str]] = {}
_PLAN_MODULE_CACHE: dict[str, str] = {}
_MODULE_NAMES: dict[str, str] = {}   # module id -> name, for readable log lines


def _is_planned(module_id: Any) -> bool:
    """True for a module a preview *would* create; it has no real Canvas ID yet."""
    return str(module_id).startswith(_PLANNED_PREFIX)


def module_items(module_id: str) -> list[dict[str, Any]]:
    if _is_planned(module_id):
        return []  # A module that does not exist yet is empty.
    return get_all_pages(f"/courses/{COURSE_ID}/modules/{module_id}/items?per_page=100")


def create_module(name: str) -> tuple[str, str]:
    """Add one UNPUBLISHED module at the end of the course and return (id, name).

    Only ever adds a module; never deletes, renames, or reorders existing ones.
    """
    if DRY_RUN:
        log(f"DRY RUN: would create a new module '{name}' (unpublished, so students can't see it) at the end of the module list")
        return _PLANNED_PREFIX + name, name
    created = api_request("POST", f"/courses/{COURSE_ID}/modules", json_body={"module": {"name": name}})
    if not isinstance(created, dict) or not created.get("id"):
        raise CanvasError(f"Canvas did not confirm the new module '{name}': {created}")
    if created.get("published"):
        api_request("PUT", f"/courses/{COURSE_ID}/modules/{created['id']}", json_body={"module": {"published": False}})
    log(f"Created new unpublished module: {name}")
    return str(created["id"]), str(created.get("name", name))


def get_material_module(folder: Path) -> tuple[str, str]:
    """Find the Canvas module for the unit folder holding this item.

    Module titles may be "Unit 4" or "Unit 4 - Data and Information"; the leading
    number is the stable match. A missing module is created (unpublished) when
    "Create missing modules" is ticked.
    """
    unit_number = None
    for candidate in (folder, folder.parent):
        unit_number = leading_number(candidate.name, "unit")
        if unit_number is not None:
            break
    if unit_number is None:
        raise CanvasError(
            f"Can't tell which {word('unit').lower()} '{folder.name}' belongs to. Lesson folders need to sit "
            f"inside a folder named like '{word('unit')} 1'."
        )
    if unit_number in _UNIT_MODULE_CACHE:
        return _UNIT_MODULE_CACHE[unit_number]

    modules = get_all_pages(f"/courses/{COURSE_ID}/modules?per_page=100")
    candidates = [m for m in modules if leading_number(str(m.get("name", "")), "unit") == unit_number]
    if len(candidates) > 1:
        raise CanvasError(
            f"More than one Canvas module starts with '{word('unit')} {unit_number}'. Rename the duplicates so the "
            f"program can choose safely: " + "; ".join(str(m.get("name")) for m in candidates)
        )
    if not candidates:
        wanted = f"{word('unit')} {unit_number}"
        if not CREATE_MISSING_MODULES:
            available = "; ".join(str(m.get("name", "(untitled)")) for m in modules)
            raise CanvasError(
                f"No Canvas module named '{wanted}' was found. Create it in Canvas, or tick "
                f"'Create missing modules' on the Course setup tab. Modules found: {available or '(none)'}."
            )
        result = create_module(wanted)
    else:
        result = (str(candidates[0]["id"]), str(candidates[0].get("name", f"{word('unit')} {unit_number}")))
    _UNIT_MODULE_CACHE[unit_number] = result
    _MODULE_NAMES[result[0]] = result[1]
    return result


def get_plan_module_id() -> str:
    """Find the lesson-plans module by name; create it (unpublished) if allowed."""
    if PLAN_MODULE_NAME in _PLAN_MODULE_CACHE:
        return _PLAN_MODULE_CACHE[PLAN_MODULE_NAME]
    modules = get_all_pages(f"/courses/{COURSE_ID}/modules?per_page=100")
    matches = [m for m in modules if str(m.get("name", "")).strip().casefold() == PLAN_MODULE_NAME.casefold()]
    if len(matches) > 1:
        raise CanvasError(f"More than one Canvas module is named '{PLAN_MODULE_NAME}'; rename one so the program doesn't have to guess.")
    if matches:
        module_id = str(matches[0]["id"])
    elif CREATE_MISSING_MODULES:
        module_id = create_module(PLAN_MODULE_NAME)[0]
    else:
        available = "; ".join(str(m.get("name", "(untitled)")) for m in modules)
        raise CanvasError(
            f"No Canvas module named '{PLAN_MODULE_NAME}' was found. Create it in Canvas, change the name on the "
            f"Names & look tab, or tick 'Create missing modules'. Modules found: {available or '(none)'}."
        )
    _PLAN_MODULE_CACHE[PLAN_MODULE_NAME] = module_id
    _MODULE_NAMES[module_id] = PLAN_MODULE_NAME
    return module_id


# ----------------------- recognizing files and folders -----------------------

def identify_files(folder: Path) -> dict[str, list[Path]]:
    """Sort a folder's files into plans, HTML lessons, quiz packages, and answer keys.

    ZIPs count as quiz packages when their name contains the checkpoint, activity,
    or exam word, or when they sit in an activity/exam folder. Other ZIPs are ignored.
    """
    files = [p for p in folder.iterdir() if p.is_file() and not p.name.startswith(("~$", "."))]
    folder_kind = norm(folder.name)
    activity_or_exam_folder = wkey("activity") in folder_kind or wkey("exam") in folder_kind
    quiz_words = (wkey("checkpoint"), wkey("activity"), wkey("exam"))
    qti_files = [
        p for p in files
        if p.suffix.lower() == ".zip"
        and (any(w in norm(p.stem) for w in quiz_words) or activity_or_exam_folder)
    ]
    by_name = lambda p: natural_key(p.name)  # noqa: E731
    return {
        "plans": sorted([p for p in files if "plan" in p.stem.lower() and p.suffix.lower() in {".doc", ".docx"}], key=by_name),
        "html": sorted([p for p in files if p.suffix.lower() in {".html", ".htm"}], key=by_name),
        "qti": sorted(qti_files, key=by_name),
        "keys": sorted([p for p in files if ("answer key" in norm(p.stem) or "answerkey" in p.stem.lower())
                        and p.suffix.lower() == ".pdf"], key=by_name),
        "decks": sorted([p for p in files if p.suffix.lower() in PPT_SUFFIXES], key=by_name),
        "links": sorted([p for p in files if p.suffix.lower() in {".url", ".webloc"}
                         or (p.suffix.lower() == ".txt" and any(w in norm(p.stem) for w in
                             (wkey("slides"), "slide", "presentation", "link")))], key=by_name),
    }


def assessment_kind(path: Path, folder: Path) -> str:
    """Classify a quiz package as checkpoint, activity, exam, or unknown."""
    filename = norm(path.stem)
    folder_name = norm(folder.name)
    if wkey("checkpoint") in filename:
        return "checkpoint"
    if wkey("exam") in filename or wkey("exam") in folder_name:
        return "exam"
    if wkey("activity") in filename or wkey("activity") in folder_name:
        return "activity"
    return "unknown"


def lesson_details(lesson_folder: Path, html_path: Path | None = None) -> tuple[str, str]:
    """Return (lesson number, lesson name), preferring the HTML file name."""
    source = html_path.stem if html_path else lesson_folder.name
    full_rx = _word_rx("lesson") + r"[\s_-]*(\d+)\s*(?:[_-]+\s*)?(.*)$"
    match = re.search(full_rx, source, re.IGNORECASE)
    if not match:
        match = re.search(full_rx, lesson_folder.name, re.IGNORECASE)
    if match:
        number = match.group(1)
        name = re.sub(r"\s+", " ", re.sub(r"[_]+", " ", match.group(2).strip(" _-"))).strip(" -_")
        if not name and html_path is not None:
            fallback = re.search(full_rx, lesson_folder.name, re.IGNORECASE)
            if fallback:
                name = re.sub(r"[_]+", " ", fallback.group(2)).strip(" -_")
        return number, name
    return "", source.replace("_", " ").strip()


def lesson_title(lesson_folder: Path, html_path: Path | None = None) -> str:
    number, name = lesson_details(lesson_folder, html_path)
    return fmt_lesson_title(number, name) or capitalize_words(lesson_folder.name)


def _unit_label_for_folder(folder: Path) -> str:
    """'Unit N' (in the chosen word) for this folder or its parent unit folder."""
    for candidate in (folder, folder.parent):
        number = leading_number(candidate.name, "unit")
        if number is not None:
            return f"{word('unit')} {number}"
    number = leading_number(UNIT_NAME, "unit")
    if number is not None:
        return f"{word('unit')} {number}"
    raise CanvasError(f"Can't tell which {word('unit').lower()} the folder '{folder}' belongs to.")


# ----------------------- uploading each kind of item -----------------------

def ensure_plan(path: Path, lesson_folder: Path, html_path: Path | None = None) -> None:
    number, _ = lesson_details(lesson_folder, html_path)
    title = plan_title(number, lesson_folder.name)
    plan_module_id = get_plan_module_id()
    existing = find_module_item(plan_module_id, "File", title)
    if existing:
        ensure_module_item_title(plan_module_id, existing, title)
    sync_file_item(path, plan_module_id, title, ind("plan"), existing, "Lesson plan")


def _answer_key_assessment_kind(path: Path, lesson_folder: Path) -> str:
    """Decide whether an answer key belongs to a unit activity, unit exam, or lesson quiz."""
    file_name = norm(path.stem)
    folder_names = [norm(lesson_folder.name), norm(lesson_folder.parent.name)]
    if wkey("activity") in file_name:
        return "activity"
    if wkey("exam") in file_name:
        return "exam"
    for folder_name in folder_names:
        if wkey("activity") in folder_name:
            return "activity"
        if wkey("exam") in folder_name:
            return "exam"
    qti_files = identify_files(lesson_folder)["qti"] if lesson_folder.is_dir() else []
    kinds = {assessment_kind(qti_path, lesson_folder) for qti_path in qti_files}
    kinds.discard("unknown")
    if len(kinds) == 1 and next(iter(kinds)) in {"activity", "exam"}:
        return next(iter(kinds))
    return "checkpoint"


def ensure_answer_key(path: Path, lesson_folder: Path, html_path: Path | None = None) -> None:
    material_module_id, material_module_name = get_material_module(lesson_folder)
    number, _ = lesson_details(lesson_folder, html_path)
    key_kind = _answer_key_assessment_kind(path, lesson_folder)
    if key_kind in {"activity", "exam"}:
        title = assessment_key_title(_unit_label_for_folder(lesson_folder), key_kind)
        indent = ind("assessment_key")
    else:
        title = checkpoint_key_title(number) if number else path.stem.replace("_", " ")
        indent = ind("checkpoint_key")
    existing = find_module_item(material_module_id, "File", title) or find_module_item(material_module_id, "File", path.name)
    # Answer keys are always hidden from students. This is not a setting.
    sync_file_item(path, material_module_id, title, indent, existing, "Answer key", hide=True)


def ensure_assessment_subheader(folder: Path, assessment_type: str) -> str:
    """Make sure a unit activity/exam has its heading, e.g. '🪛Unit 3 Activity'."""
    if assessment_type not in {"activity", "exam"}:
        return ""
    material_module_id, _ = get_material_module(folder)
    title = assessment_header_title(_unit_label_for_folder(folder), assessment_type)
    existing = find_module_item(material_module_id, "SubHeader", title)
    if existing:
        ensure_module_item_title(material_module_id, existing, title)
        log(f"Heading already exists: {title}")
    elif DRY_RUN:
        log(f"DRY RUN: would add heading '{title}' (indent {ind('assessment_header')})")
    else:
        api_request("POST", f"/courses/{COURSE_ID}/modules/{material_module_id}/items",
                    json_body={"module_item": {"type": "SubHeader", "title": title, "indent": ind("assessment_header")}})
        log(f"Added heading: {title}")
    return title


def process_lesson(folder: Path) -> None:
    if not folder.is_dir():
        raise CanvasError(f"Folder not found: {folder}")
    material_module_id, material_module_name = get_material_module(folder)
    log("=" * 60)
    log(f"Processing {folder.name}  →  Canvas module '{material_module_name}'")
    groups = identify_files(folder)
    html_path = groups["html"][0] if groups["html"] else None
    number, name = lesson_details(folder, html_path)
    header = lesson_header_title(number, name) if number else with_emoji("lesson_header", lesson_title(folder, html_path))

    # Keep only the kinds of files chosen on the Run tab.
    if not UPLOAD_LESSON_PLANS:
        groups["plans"] = []
    if not UPLOAD_HTML_LESSONS:
        groups["html"] = []
    if not UPLOAD_ANSWER_KEYS:
        groups["keys"] = []
    if not UPLOAD_SLIDES:
        groups["decks"], groups["links"] = [], []
    enabled_qti: list[Path] = []
    for qti_path in groups["qti"]:
        kind = assessment_kind(qti_path, folder)
        if {"checkpoint": UPLOAD_CHECKPOINTS, "activity": UPLOAD_ACTIVITIES, "exam": UPLOAD_EXAMS}.get(kind, False):
            enabled_qti.append(qti_path)
        elif kind == "unknown":
            log(f"Skipping quiz package (can't tell if it is a {word('checkpoint').lower()}, "
                f"{word('activity').lower()}, or {word('exam').lower()}): {qti_path.name}")
    groups["qti"] = enabled_qti

    counts = ", ".join(f"{label} {len(groups[key])}" for key, label in
                       (("plans", "plans"), ("html", "HTML pages"), ("decks", "PowerPoints"),
                        ("links", "presentation links"), ("qti", "quizzes"), ("keys", "answer keys")))
    log(f"Files selected: {counts}")
    make_header = UPLOAD_LESSON_HEADERS and is_lesson_name(folder.name)
    drive_slides = UPLOAD_SLIDES and bool(DRIVE_FOLDER) and is_lesson_name(folder.name)
    if not any(groups.values()) and not make_header and not drive_slides:
        log("Nothing selected to upload in this folder.")
        return

    for path in groups["plans"]:
        ensure_plan(path, folder, html_path)

    if make_header:
        existing_header = find_module_item(material_module_id, "SubHeader", header)
        if existing_header:
            ensure_module_item_title(material_module_id, existing_header, header)
            log(f"Lesson heading already exists: {header}")
        elif DRY_RUN:
            log(f"DRY RUN: would add lesson heading '{header}' (indent {ind('lesson_header')})")
        else:
            api_request("POST", f"/courses/{COURSE_ID}/modules/{material_module_id}/items",
                        json_body={"module_item": {"type": "SubHeader", "title": header, "indent": ind("lesson_header")}})
            log(f"Added lesson heading: {header}")

    for path in groups["html"]:
        ensure_page(path, folder)
    if UPLOAD_SLIDES:
        ensure_presentations(folder, groups["links"], groups["decks"], html_path)
    for path in groups["qti"]:
        kind = assessment_kind(path, folder)
        if kind in {"activity", "exam"}:
            ensure_assessment_subheader(folder, kind)
        ensure_quiz(path, folder, html_path)
    if any(assessment_kind(p, folder) in {"activity", "exam"} for p in groups["qti"]) and not groups["keys"] and UPLOAD_ANSWER_KEYS:
        log("WARNING: this activity/exam folder has no answer key PDF. Expected a PDF with 'Answer Key' in its name in the same folder.")
    for path in groups["keys"]:
        ensure_answer_key(path, folder, html_path)


# ----------------------- presentations (four ways) -----------------------
# 1. A link in the lesson folder (.url / .webloc shortcut, or a "slides link" .txt file).
#    Google Slides, Canva, and other embeddable links become a page; others open in a new tab.
# 2. A PowerPoint in the lesson folder, shown on a Canvas page with Canvas's own viewer.
# 3. A PowerPoint converted to Google Slides in your Drive, published, and embedded.
# 4. Google Slides already in your Drive course folder, matched to each lesson and embedded.
# A link in the folder always wins over a PowerPoint. Drive slides (#4) are only used for
# lessons with no link and no PowerPoint. Nothing is ever deleted in Canvas or in Drive.

GOOGLE_CLIENT_PATH = APP_DIR / "google_client_secret.json"   # Made once by whoever shares the program.
GOOGLE_TOKEN_PATH = APP_DIR / "google_token.json"            # Each person's own sign-in. PRIVATE.
GOOGLE_SCOPE_FILE = "https://www.googleapis.com/auth/drive.file"   # Only files this program makes.
GOOGLE_SCOPE_FULL = "https://www.googleapis.com/auth/drive"        # Needed to read your own Drive folder.
DRIVE_API = "https://www.googleapis.com/drive/v3"
SLIDES_MIME = "application/vnd.google-apps.presentation"
FOLDER_MIME = "application/vnd.google-apps.folder"
EMBED_OPTIONS = "start=false&loop=false&delayms=3000"
SLIDES_IFRAME_HEIGHT = "569"
PPT_SUFFIXES = {".ppt", ".pptx", ".pps", ".ppsx"}
OLD_VERSION_PREFIX = "OLD - "

_DRIVE_INDEX: list[dict[str, Any]] | None = None
_DRIVE_FOLDERS: dict[str, str] = {}
_GOOGLE_TOKEN: dict[str, Any] = {}


class GoogleError(CanvasError):
    pass


def slides_title(number: str, name: str = "", folder_name: str = "", label: str = "") -> str:
    base = f"{word('lesson')} {number} {word('slides')}" if number else f"{capitalize_words(folder_name)} {word('slides')}"
    if label:
        base += f" - {capitalize_words(label)}"
    return with_emoji("slides", base)


def drive_folder_id(text: str) -> str:
    """Accept a pasted Drive folder link (…/folders/ID or ?id=ID) or a bare folder ID."""
    text = (text or "").strip()
    match = re.search(r"/folders/([\w-]{10,})", text) or re.search(r"[?&]id=([\w-]{10,})", text)
    if match:
        return match.group(1)
    return text if re.fullmatch(r"[\w-]{10,}", text) else ""


# ---------- links (way 1) ----------

def read_link_file(path: Path) -> list[tuple[str, str]]:
    """Return (label, url) pairs from a .url, .webloc, or .txt link file."""
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return []
    suffix = path.suffix.lower()
    if suffix == ".url":
        match = re.search(r"^\s*URL\s*=\s*(\S+)", text, re.IGNORECASE | re.MULTILINE)
        return [(path.stem, match.group(1))] if match else []
    if suffix == ".webloc":
        match = re.search(r"<string>\s*(https?://[^<\s]+)\s*</string>", text)
        return [(path.stem, match.group(1))] if match else []
    pairs: list[tuple[str, str]] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        iframe = re.search(r"<iframe[^>]*\ssrc\s*=\s*[\"']([^\"']+)", line, re.IGNORECASE)
        match = iframe or re.search(r"https?://\S+", line)
        if not match:
            continue
        url = match.group(1) if iframe else match.group(0)
        label = "" if iframe else line[:match.start()].strip(" :-|\t")
        pairs.append((label, url))
    return pairs


def embed_url_for(url: str) -> str | None:
    """The address to put in an iframe, or None when the link should open in a new tab."""
    url = url.strip().replace("&amp;", "&")
    published = re.search(r"docs\.google\.com/presentation/d/e/([\w-]+)", url)
    if published:
        return f"https://docs.google.com/presentation/d/e/{published.group(1)}/embed?{EMBED_OPTIONS}"
    shared = re.search(r"docs\.google\.com/presentation/d/([\w-]{20,})", url)
    if shared:
        return f"https://docs.google.com/presentation/d/{shared.group(1)}/embed?{EMBED_OPTIONS}"
    if "embed" in url.lower():   # Canva ?embed, OneDrive /embed, and similar ready-made embed links.
        return url
    return None


def slides_iframe(src: str) -> str:
    safe = src.replace("&", "&amp;").replace('"', "&quot;")
    return (f'<p><iframe src="{safe}" width="100%" height="{SLIDES_IFRAME_HEIGHT}" frameborder="0" '
            'allowfullscreen="true" mozallowfullscreen="true" webkitallowfullscreen="true"></iframe></p>')


# ---------- Canvas side shared by all four ways ----------

def _slug(title: str) -> str:
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9_-]+", "-", title.lower())).strip("-")[:200] or "slides"


def upsert_slides_page(lesson_folder: Path, title: str, body: str, marker: str) -> None:
    """Create or update the Canvas page that shows a presentation, and list it in the module.

    `marker` is the address that must be on the page; if it already is, nothing changes.
    """
    module_id, _ = get_material_module(lesson_folder)
    item = find_module_item(module_id, "Page", title)
    if item:
        ensure_module_item_title(module_id, item, title)
    page_url = item.get("page_url") if item else None
    page = None
    if page_url:
        page = api_request("GET", f"/courses/{COURSE_ID}/pages/{urllib.parse.quote(page_url, safe='')}", allow_404=True)
    else:
        # Exact title or address only: "Lesson 3 Slides" must never land on "Lesson 32 Slides".
        same = [p for p in page_list() if norm(p.get("title", "")) == norm(title) or p.get("url") == _slug(title)]
        if len(same) > 1:
            log(f"WARNING: Skipping '{title}': more than one Canvas page has that title, so the program won't guess.")
            return
        if same:
            page = api_request("GET", f"/courses/{COURSE_ID}/pages/{urllib.parse.quote(same[0]['url'], safe='')}", allow_404=True)
    current = str((page or {}).get("body") or "").replace("&amp;", "&")
    if page and marker.replace("&amp;", "&") in current:
        log(f"Presentation page unchanged; skipping: {title}")
    elif DRY_RUN:
        log(f"DRY RUN: would {'update' if page else 'create'} page '{title}' showing the presentation")
    elif page:
        api_request("PUT", f"/courses/{COURSE_ID}/pages/{urllib.parse.quote(page['url'], safe='')}",
                    json_body={"wiki_page": {"title": title, "body": body}})
        log(f"Updated presentation page: {title}")
    else:
        page = api_request("POST", f"/courses/{COURSE_ID}/pages", json_body={"wiki_page": {
            "title": title, "url": _slug(title), "body": body, "editing_roles": "teachers",
            "published": PUBLISH_NEW_PAGES}})
        log(f"Created presentation page: {title}")
    if not item:
        add_module_item(module_id, "Page", title, indent=ind("slides"),
                        page_url=(page or {}).get("url") or _slug(title))


def ensure_link_item(lesson_folder: Path, title: str, url: str) -> None:
    """A link that can't be embedded becomes a module link that opens in a new tab."""
    module_id, _ = get_material_module(lesson_folder)
    item = find_module_item(module_id, "ExternalUrl", title)
    if not item:
        add_module_item(module_id, "ExternalUrl", title, indent=ind("slides"),
                        extra={"external_url": url, "new_tab": True})
        return
    ensure_module_item_title(module_id, item, title)
    if (item.get("external_url") or "").strip() == url:
        log(f"Presentation link unchanged; skipping: {title}")
    elif DRY_RUN:
        log(f"DRY RUN: would point '{title}' at the new link")
    else:
        api_request("PUT", f"/courses/{COURSE_ID}/modules/{module_id}/items/{item['id']}",
                    json_body={"module_item": {"external_url": url}})
        log(f"Updated presentation link: {title}")


def canvas_preview_body(file_id: Any, name: str) -> str:
    """Canvas's own file link with 'auto_open', so the slides open in Canvas's viewer on the page."""
    href = f"{BASE_URL}/courses/{COURSE_ID}/files/{file_id}?wrap=1"
    api = f"{BASE_URL}/api/v1/courses/{COURSE_ID}/files/{file_id}"
    return (f'<p><a class="instructure_file_link instructure_scribd_file auto_open" title="{name}" '
            f'href="{href}" data-api-endpoint="{api}" data-api-returntype="File">{name}</a></p>')


# ---------- Google sign-in (standard library only) ----------

def google_ready() -> bool:
    return GOOGLE_CLIENT_PATH.exists()


def _google_client() -> dict[str, str]:
    try:
        data = json.loads(GOOGLE_CLIENT_PATH.read_text(encoding="utf-8"))
        client = data.get("installed") or data.get("web") or {}
        if not client.get("client_id"):
            raise ValueError
        return client
    except (OSError, ValueError, AttributeError) as exc:
        raise GoogleError(f"{GOOGLE_CLIENT_PATH.name} is missing or damaged. Ask whoever shared this "
                          "program for a fresh copy and put it next to the program.") from exc


def _google_token_post(fields: dict[str, str]) -> dict[str, Any]:
    client = _google_client()
    fields = {"client_id": client["client_id"], "client_secret": client.get("client_secret", ""), **fields}
    request = urllib.request.Request(client.get("token_uri", "https://oauth2.googleapis.com/token"),
                                     data=urllib.parse.urlencode(fields).encode(), method="POST")
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise GoogleError(f"Google sign-in was refused: {exc.read().decode('utf-8', 'replace')[:300]}") from exc
    except urllib.error.URLError as exc:
        raise GoogleError(f"Couldn't reach Google: {exc}") from exc


def _save_google_token(token: dict[str, Any]) -> None:
    token["expires_at"] = time.time() + int(token.get("expires_in", 3600)) - 60
    _GOOGLE_TOKEN.clear()
    _GOOGLE_TOKEN.update(token)
    GOOGLE_TOKEN_PATH.write_text(json.dumps(token, indent=2), encoding="utf-8")


def google_sign_in(scope: str) -> None:
    """Open the browser once so the person can sign in and click Allow."""
    import base64
    import http.server
    import secrets
    import webbrowser

    client = _google_client()
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
    state = secrets.token_urlsafe(16)
    result: dict[str, str] = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802
            query = {k: v[0] for k, v in urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).items()}
            if "code" in query or "error" in query:
                result.update(query)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            message = "Signed in. You can close this tab and go back to Canvas Course Manager." \
                if "code" in query else "Sign-in didn't finish. You can close this tab."
            self.wfile.write(f"<html><body style='font-family:sans-serif;padding:40px'><h2>{message}</h2>"
                             "</body></html>".encode("utf-8"))

        def log_message(self, *args: Any) -> None:
            pass

    server = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    server.timeout = 1
    redirect = f"http://127.0.0.1:{server.server_port}"
    auth = client.get("auth_uri", "https://accounts.google.com/o/oauth2/auth") + "?" + urllib.parse.urlencode({
        "client_id": client["client_id"], "redirect_uri": redirect, "response_type": "code", "scope": scope,
        "access_type": "offline", "prompt": "consent", "state": state,
        "code_challenge": challenge, "code_challenge_method": "S256"})
    log("A browser window is opening. Sign in with your school Google account and click Allow.")
    webbrowser.open(auth)
    deadline = time.time() + 300
    try:
        while not result and time.time() < deadline:
            server.handle_request()
    finally:
        server.server_close()
    if not result:
        raise GoogleError("Google sign-in timed out after 5 minutes. Run again when you're ready to sign in.")
    if result.get("error") or result.get("state") != state:
        raise GoogleError("Google sign-in was cancelled or blocked. If Google said your school blocks this app, "
                          "use presentation links or 'Show in Canvas' instead.")
    token = _google_token_post({"code": result["code"], "code_verifier": verifier,
                                "redirect_uri": redirect, "grant_type": "authorization_code"})
    token.setdefault("scope", scope)
    _save_google_token(token)
    log("Signed in to Google.")


def google_access_token(scope: str) -> str:
    """A working access token with at least `scope`, signing in or refreshing as needed."""
    if not _GOOGLE_TOKEN and GOOGLE_TOKEN_PATH.exists():
        try:
            _GOOGLE_TOKEN.update(json.loads(GOOGLE_TOKEN_PATH.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass
    granted = str(_GOOGLE_TOKEN.get("scope", "")).split()
    covered = scope in granted or (scope == GOOGLE_SCOPE_FILE and GOOGLE_SCOPE_FULL in granted)
    if not covered or not _GOOGLE_TOKEN.get("refresh_token"):
        google_sign_in(scope)
    elif time.time() > float(_GOOGLE_TOKEN.get("expires_at", 0)):
        fresh = _google_token_post({"refresh_token": _GOOGLE_TOKEN["refresh_token"], "grant_type": "refresh_token"})
        _save_google_token({**_GOOGLE_TOKEN, **fresh})
    return str(_GOOGLE_TOKEN["access_token"])


def google_scope() -> str:
    """Reading your own Drive folder needs the wider permission; otherwise only the program's own files."""
    return GOOGLE_SCOPE_FULL if DRIVE_FOLDER else GOOGLE_SCOPE_FILE


def google_http(method: str, url: str, *, json_body: Any = None, data: bytes | None = None,
                headers: dict[str, str] | None = None) -> tuple[dict[str, Any], Any]:
    """Call Google; returns (response headers, parsed JSON or {})."""
    body = json.dumps(json_body).encode("utf-8") if json_body is not None else data
    all_headers = {"Authorization": f"Bearer {google_access_token(google_scope())}", **(headers or {})}
    if json_body is not None:
        all_headers["Content-Type"] = "application/json; charset=UTF-8"
    request = urllib.request.Request(url, data=body, headers=all_headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=300) as response:
            raw = response.read()
            return dict(response.headers), (json.loads(raw.decode("utf-8")) if raw else {})
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        raise GoogleError(f"Google Drive said no (code {exc.code}) to {method} {url.split('?')[0]}: "
                          f"{detail[:600]}") from exc
    except urllib.error.URLError as exc:
        raise GoogleError(f"Couldn't reach Google Drive: {exc}") from exc


def drive_children(folder_id: str) -> list[dict[str, Any]]:
    files: list[dict[str, Any]] = []
    token = ""
    while True:
        params = {"q": f"'{folder_id}' in parents and trashed=false", "pageSize": "200",
                  "fields": "nextPageToken,files(id,name,mimeType)",
                  "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}
        if token:
            params["pageToken"] = token
        _, data = google_http("GET", f"{DRIVE_API}/files?" + urllib.parse.urlencode(params))
        files += data.get("files", [])
        token = data.get("nextPageToken", "")
        if not token:
            return files


def drive_folder(parent: str, name: str) -> str:
    """Find or create a Drive folder by name inside `parent` ('root' = My Drive)."""
    key = f"{parent}/{name}"
    if key in _DRIVE_FOLDERS:
        return _DRIVE_FOLDERS[key]
    safe = name.replace("\\", "\\\\").replace("'", "\\'")
    params = {"q": f"'{parent}' in parents and name='{safe}' and mimeType='{FOLDER_MIME}' and trashed=false",
              "fields": "files(id)", "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}
    _, data = google_http("GET", f"{DRIVE_API}/files?" + urllib.parse.urlencode(params))
    found = data.get("files", [])
    if found:
        folder_id = found[0]["id"]
    else:
        _, made = google_http("POST", f"{DRIVE_API}/files?supportsAllDrives=true&fields=id",
                              json_body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent]})
        folder_id = made["id"]
        log(f"Created Google Drive folder: {name}")
    _DRIVE_FOLDERS[key] = folder_id
    return folder_id


def drive_upload_as_slides(path: Path, parent: str) -> str:
    """Upload a PowerPoint and let Google convert it to Slides (resumable, so big files work)."""
    mime = mimetypes.guess_type(str(path))[0] or "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    head, _ = google_http(
        "POST", "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true",
        json_body={"name": path.stem, "mimeType": SLIDES_MIME, "parents": [parent]},
        headers={"X-Upload-Content-Type": mime, "X-Upload-Content-Length": str(path.stat().st_size)})
    location = next((v for k, v in head.items() if k.lower() == "location"), "")
    if not location:
        raise GoogleError(f"Google didn't accept the upload of {path.name}.")
    _, made = google_http("PUT", location, data=path.read_bytes(), headers={"Content-Type": mime})
    return str(made["id"])


def drive_publish(file_id: str, name: str) -> str:
    """Publish to the web (if it isn't already) and return the embed address."""
    url = f"{DRIVE_API}/files/{file_id}/revisions/head?fields=published,publishedLink,publishedOutsideDomain"
    _, head = google_http("GET", url)
    if not head.get("published"):
        if DRY_RUN:
            log(f"DRY RUN: would publish '{name}' to the web")
        else:
            try:
                _, head = google_http("PATCH", url, json_body={"published": True, "publishAuto": True,
                                                               "publishedOutsideDomain": True})
            except GoogleError:
                _, head = google_http("PATCH", url, json_body={"published": True, "publishAuto": True})
                log(f"WARNING: '{name}' is published for your school only. Students must be signed in "
                    "to their school Google account to see it.")
            log(f"Published to the web: {name}")
    link = str(head.get("publishedLink") or "")
    match = re.search(r"/presentation/d/e/([\w-]+)", link)
    if match:
        return f"https://docs.google.com/presentation/d/e/{match.group(1)}/embed?{EMBED_OPTIONS}"
    return f"https://docs.google.com/presentation/d/{file_id}/embed?{EMBED_OPTIONS}"


# ---------- way 3: PowerPoint → Google Slides ----------

def ensure_google_slides(path: Path, lesson_folder: Path, title: str) -> None:
    digest, status = sync_status(path)
    saved = STATE.get(file_state_key(path), {})
    module_id, _ = get_material_module(lesson_folder)
    in_module = find_module_item(module_id, "Page", title)
    if in_module and status == "same":
        if saved.get("embed"):
            upsert_slides_page(lesson_folder, title, slides_iframe(saved["embed"]), saved["embed"])
        else:
            log(f"Presentation unchanged; skipping: {title}")
        return
    if in_module and status == "new":
        log(f"Presentation already in module; recording current version so future changes upload: {title}")
        remember_upload(path, digest, None)
        return
    root_name = DRIVE_FOLDER or f"Canvas Course {COURSE_ID}"
    where = f"{root_name if not DRIVE_FOLDER else 'your Drive course folder'} / {lesson_folder.parent.name} / {lesson_folder.name}"
    if DRY_RUN:
        log(f"DRY RUN: would {'upload the changed' if status == 'changed' else 'upload'} {path.name} to Google Drive "
            f"({where}), convert it to Google Slides, publish it, and show it on page '{title}'")
        if not in_module:
            add_module_item(module_id, "Page", title, indent=ind("slides"), page_url=_slug(title))
        return
    parent = drive_folder_id(DRIVE_FOLDER) if DRIVE_FOLDER else drive_folder("root", root_name)
    for part in (lesson_folder.parent.name, lesson_folder.name):
        parent = drive_folder(parent, part)
    log(f"Uploading {path.name} to Google Drive and converting it to Google Slides…")
    new_id = drive_upload_as_slides(path, parent)
    old_id = saved.get("drive_id")
    if old_id and old_id != new_id:
        # The old version stays in Drive (never deleted); it is renamed so it's easy to tell apart.
        google_http("PATCH", f"{DRIVE_API}/files/{old_id}?supportsAllDrives=true",
                    json_body={"name": f"{OLD_VERSION_PREFIX}{path.stem}"})
        log(f"Kept the previous Google Slides version, renamed '{OLD_VERSION_PREFIX}{path.stem}'")
    embed = drive_publish(new_id, path.stem)
    upsert_slides_page(lesson_folder, title, slides_iframe(embed), embed)
    STATE[file_state_key(path)] = {"sha256": digest, "file_id": None, "drive_id": new_id, "embed": embed,
                                   "recorded_at": time.strftime("%Y-%m-%d %H:%M:%S")}
    save_state()


# ---------- way 4: Google Slides already in a Drive folder ----------

def _number_after(key: str, text: str) -> int | None:
    match = re.search(r"(?<![a-z])" + _word_rx(key) + r"[\s_-]*(\d+)(?!\d)", text, re.IGNORECASE)
    return int(match.group(1)) if match else None


def drive_slides_index() -> list[dict[str, Any]]:
    """Every Google Slides file under the Drive course folder (four folders deep), read once per run."""
    global _DRIVE_INDEX
    if _DRIVE_INDEX is not None:
        return _DRIVE_INDEX
    root = drive_folder_id(DRIVE_FOLDER)
    if not root:
        raise GoogleError("The Google Drive course folder on the Course setup tab isn't a folder link.")
    log("Reading your Google Drive course folder…")
    found: list[dict[str, Any]] = []
    queue_: list[tuple[str, list[str], int]] = [(root, [], 0)]
    while queue_:
        folder_id, chain, depth = queue_.pop(0)
        for entry in drive_children(folder_id):
            if entry.get("mimeType") == FOLDER_MIME and depth < 4:
                queue_.append((entry["id"], chain + [entry["name"]], depth + 1))
            elif entry.get("mimeType") == SLIDES_MIME and not entry["name"].startswith(OLD_VERSION_PREFIX):
                names = chain + [entry["name"]]
                lesson = _number_after("lesson", entry["name"])
                if lesson is None:
                    lesson = next((n for n in (_number_after("lesson", c) for c in reversed(chain)) if n is not None), None)
                unit = next((n for n in (_number_after("unit", c) for c in names) if n is not None), None)
                found.append({"id": entry["id"], "name": entry["name"], "path": " / ".join(names),
                              "lesson": lesson, "unit": unit})
    _DRIVE_INDEX = found
    log(f"Found {len(found)} Google Slides file(s) in the Drive folder.")
    return found


def ensure_drive_slides(lesson_folder: Path, number: str, title: str) -> None:
    if not number:
        return
    unit = leading_number(lesson_folder.parent.name, "unit")
    matches = [f for f in drive_slides_index()
               if f["lesson"] == int(number) and (f["unit"] is None or unit is None or f["unit"] == unit)]
    if not matches:
        log(f"No Google Slides for {word('lesson')} {number} in the Drive folder.")
        return
    if len(matches) > 1:
        log(f"WARNING: Skipping Drive slides for {word('lesson')} {number}: more than one could match, so the "
            "program won't guess. Rename or move all but one: " + "; ".join(m["path"] for m in matches))
        return
    match = matches[0]
    log(f"Using Google Slides from Drive: {match['path']}")
    embed = drive_publish(match["id"], match["name"])
    upsert_slides_page(lesson_folder, title, slides_iframe(embed), embed)


# ---------- choosing the way for each lesson ----------

def ensure_presentations(lesson_folder: Path, links: list[Path], decks: list[Path],
                         html_path: Path | None) -> None:
    number, name = lesson_details(lesson_folder, html_path)
    pairs = [pair for path in links for pair in read_link_file(path)]
    for path in links:
        if not read_link_file(path):
            log(f"WARNING: No web address found in {path.name}; skipping it.")
    if pairs:
        if decks:
            log("A presentation link is in this folder, so the PowerPoint file(s) here are not uploaded.")
        for index, (label, url) in enumerate(pairs, start=1):
            if len(pairs) > 1 and not label:
                label = f"Part {index}"
            title = slides_title(number, name, lesson_folder.name, label if len(pairs) > 1 else "")
            embed = embed_url_for(url)
            if embed:
                upsert_slides_page(lesson_folder, title, slides_iframe(embed), embed)
            else:
                ensure_link_item(lesson_folder, title, url)
        return

    mode = SLIDES_MODE
    if decks and mode == "google" and not google_ready():
        log(f"WARNING: Google Slides is chosen, but {GOOGLE_CLIENT_PATH.name} isn't next to the program. "
            "Showing the PowerPoint in Canvas instead for this run.")
        mode = "canvas"
    if decks and mode != "off":
        for path in decks:
            title = slides_title(number, name, lesson_folder.name, path.stem if len(decks) > 1 else "")
            if mode == "google":
                ensure_google_slides(path, lesson_folder, title)
                continue
            digest, status = sync_status(path)
            module_id, _ = get_material_module(lesson_folder)
            in_module = find_module_item(module_id, "Page", title)
            if in_module and status in {"new", "same"}:
                log(f"Presentation {'unchanged' if status == 'same' else 'already in module'}; skipping: {title}")
                if status == "new":
                    remember_upload(path, digest, None)
                continue
            uploaded = upload_course_file(path, upload_folder_for(path))
            file_id = uploaded.get("id")
            marker = f"/files/{file_id}?wrap=1" if file_id else "(preview)"
            upsert_slides_page(lesson_folder, title, canvas_preview_body(file_id, path.name), marker)
            remember_upload(path, digest, file_id)
        return

    if DRIVE_FOLDER:
        if not google_ready():
            log(f"WARNING: Can't read your Drive folder without {GOOGLE_CLIENT_PATH.name} next to the program.")
            return
        ensure_drive_slides(lesson_folder, number, slides_title(number, name, lesson_folder.name))


# ----------------------- choosing folders -----------------------

def _is_lesson_activity_or_exam_folder(path: Path) -> bool:
    if not path.is_dir() or path.name.startswith("."):
        return False
    name = norm(path.name)
    return is_lesson_name(path.name) or wkey("activity") in name or wkey("exam") in name


def _has_assessment_zip(folder: Path) -> bool:
    if not folder.is_dir():
        return False
    quiz_words = (wkey("checkpoint"), wkey("activity"), wkey("exam"))
    return any(p.is_file() and p.suffix.lower() == ".zip" and any(w in norm(p.stem) for w in quiz_words)
               for p in folder.iterdir())


def list_units(course_folder: Path) -> list[str]:
    """Unit folder names, in natural order (Unit 2 before Unit 10)."""
    try:
        return sorted((p.name for p in course_folder.iterdir() if p.is_dir() and is_unit_name(p.name)), key=natural_key)
    except OSError:
        return []


def list_unit_folders(course_folder: Path, unit_name: str) -> list[str]:
    """Lesson, activity, and exam folders inside one unit."""
    unit = course_folder / unit_name
    try:
        names = [p.name for p in unit.iterdir() if _is_lesson_activity_or_exam_folder(p)]
    except OSError:
        return []
    return sorted(names, key=natural_key)


def choose_folders() -> list[Path]:
    if SCOPE == "lesson":
        return [MAIN_DIR / UNIT_NAME / LESSON_NAME]
    units = [MAIN_DIR / UNIT_NAME] if SCOPE == "unit" else [MAIN_DIR / name for name in list_units(MAIN_DIR)]
    found: list[Path] = []
    for unit in units:
        if not unit.is_dir():
            raise CanvasError(f"{word('unit')} folder not found: {unit}")
        found.extend(sorted((p for p in unit.iterdir() if _is_lesson_activity_or_exam_folder(p)),
                            key=lambda p: natural_key(p.name)))
        # Some courses keep the activity/exam ZIP directly inside the unit folder.
        if _has_assessment_zip(unit):
            found.append(unit)
    return found


def validate_canvas_targets() -> None:
    """Read-only check, so it also runs during a preview."""
    course = api_request("GET", f"/courses/{COURSE_ID}")
    if not course or str(course.get("id")) != str(COURSE_ID):
        raise CanvasError(f"Could not open Canvas course {COURSE_ID}. Check the course ID on the Course setup tab.")
    log(f"Connected to Canvas course {COURSE_ID}: {course.get('name', '')}")
    if UPLOAD_LESSON_PLANS:
        plan_id = get_plan_module_id()
        if not _is_planned(plan_id):
            log(f"Lesson plans go to module '{PLAN_MODULE_NAME}' (ID {plan_id}).")
    kinds = {kind for kind, on in (("checkpoint", UPLOAD_CHECKPOINTS), ("activity", UPLOAD_ACTIVITIES),
                                   ("exam", UPLOAD_EXAMS)) if on}
    if kinds:
        resolve_group_targets(kinds)


# ----------------------- gradebook groups (Canvas "assignment groups") -----------------------
# Canvas puts imported quizzes in a group named "Imported Assignments". In a course with
# weighted grading that group usually counts 0%, so each course folder remembers which group
# new checkpoints, activities, and exams belong in (saved by group ID, plus the name so the
# file is readable). Safety rules:
#   * Groups are created only when a person asks in the window, sees a preview, and types APPLY.
#   * A quiz this run imports goes straight into its chosen group.
#   * A quiz that was already in Canvas is moved only while it is still in "Imported Assignments",
#     so a group someone picked by hand is never overridden.
#   * A chosen group that can't be found is never guessed at; those quizzes stay where Canvas put them.

QUIZ_KINDS = ("checkpoint", "activity", "exam")
IMPORTED_GROUP_NAME = "Imported Assignments"
GROUP_TARGETS: dict[str, dict[str, Any]] = {}   # kind -> Canvas group, worked out at the start of a run.
_GROUP_NAMES: dict[str, str] = {}               # group ID -> name, for readable log lines.


def percent_text(value: Any) -> str:
    """20 -> '20%', 12.5 -> '12.5%'."""
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    return f"{number:g}%"


def read_assignment_groups(course_id: str) -> tuple[list[dict[str, Any]], bool]:
    """The course's gradebook groups (in Canvas order) and whether weighted grading is on. Reads only."""
    course = api_request("GET", f"/courses/{course_id}")
    groups = get_all_pages(f"/courses/{course_id}/assignment_groups?per_page=100")
    groups.sort(key=lambda g: int(g.get("position") or 0))
    return groups, bool(isinstance(course, dict) and course.get("apply_assignment_group_weights"))


def create_assignment_group(course_id: str, name: str, weight: float) -> dict[str, Any]:
    """Add one gradebook group. Only called after a person previewed it and typed APPLY."""
    created = api_request("POST", f"/courses/{course_id}/assignment_groups",
                          form_body={"name": name, "group_weight": weight})
    if not isinstance(created, dict) or not created.get("id"):
        raise CanvasError(f"Canvas did not confirm the new gradebook group '{name}': {created}")
    log(f"Created gradebook group '{created.get('name', name)}' ({percent_text(created.get('group_weight', weight))}) "
        f"in course {course_id}.")
    return created


def group_choice_for(kind: str) -> dict[str, Any] | None:
    """This course's saved choice for one kind of quiz, or None when nothing was chosen yet."""
    if str(GROUP_CHOICES.get("course_id") or "") != str(COURSE_ID):
        return None   # Saved for a different course ID: group IDs don't carry over.
    choice = GROUP_CHOICES.get(kind)
    return choice if isinstance(choice, dict) else None


def resolve_group_targets(kinds: set[str]) -> None:
    """Look up the chosen gradebook groups in Canvas (read-only) before anything is uploaded."""
    GROUP_TARGETS.clear()
    _GROUP_NAMES.clear()
    choices = {kind: group_choice_for(kind) for kind in QUIZ_KINDS if kind in kinds}
    groups: list[dict[str, Any]] = []
    weighted = False
    if any(c and c.get("id") is not None for c in choices.values()):
        groups, weighted = read_assignment_groups(COURSE_ID)
        _GROUP_NAMES.update({str(g.get("id")): _cell_text(g.get("name")) for g in groups})
    by_id = {str(g.get("id")): g for g in groups}
    for kind, choice in choices.items():
        label = plural(word(kind)).lower()
        if not choice:
            log(f"New {label} stay where Canvas puts them (usually '{IMPORTED_GROUP_NAME}'). "
                "Choose their gradebook group on the Course setup tab to change this.")
            continue
        if choice.get("leave"):
            log(f"New {label} stay where Canvas puts them, as chosen on the Course setup tab.")
            continue
        group = by_id.get(str(choice.get("id")))
        if group is None:
            same_name = [g for g in groups if norm(_cell_text(g.get("name"))) == norm(_cell_text(choice.get("name")))]
            if len(same_name) == 1:
                group = same_name[0]
                log(f"NOTE: the gradebook group '{choice.get('name')}' has a new ID in Canvas ({group.get('id')}). "
                    "Using it. Click Load groups from Canvas on the Course setup tab to save the new ID.")
            else:
                log(f"WARNING: the gradebook group chosen for {label} ('{choice.get('name')}') isn't in Canvas any more. "
                    f"New {label} will stay where Canvas puts them. Choose a group again on the Course setup tab.")
                continue
        GROUP_TARGETS[kind] = group
        log(f"New {label} go into gradebook group '{group.get('name')}'"
            + (f" ({percent_text(group.get('group_weight'))} of the grade)." if weighted else "."))
        if weighted and not float(group.get("group_weight") or 0):
            log(f"WARNING: '{group.get('name')}' counts 0% of the grade, so {label} in it won't affect grades. "
                "Change its percentage in Canvas if that isn't what you want.")


def planned_group_change(assignment: dict[str, Any], kind: str, title: str, *, new_import: bool) -> dict[str, Any]:
    """{'assignment_group_id': ID} when this quiz should move into its chosen group, otherwise {}.

    A quiz this run just imported always goes into the chosen group. A quiz that was already in
    Canvas moves only while it is still in "Imported Assignments".
    """
    target = GROUP_TARGETS.get(kind)
    if not target or target.get("id") is None:
        return {}
    current_id = assignment.get("assignment_group_id")
    if str(current_id) == str(target["id"]):
        return {}
    if not new_import and norm(_GROUP_NAMES.get(str(current_id), "")) != norm(IMPORTED_GROUP_NAME):
        return {}   # Someone chose this quiz's group by hand: leave it.
    if DRY_RUN:
        log(f"DRY RUN: would move '{title}' into gradebook group '{target.get('name')}'")
    else:
        log(f"Moving '{title}' into gradebook group '{target.get('name')}'")
    return {"assignment_group_id": int(target["id"])}


def confirm_group_change(updated: Any, changes: dict[str, Any], title: str) -> None:
    """Check Canvas's reply really shows the quiz in its new group."""
    if "assignment_group_id" not in changes:
        return
    name = _GROUP_NAMES.get(str(changes["assignment_group_id"]), f"ID {changes['assignment_group_id']}")
    if isinstance(updated, dict) and str(updated.get("assignment_group_id")) == str(changes["assignment_group_id"]):
        log(f"Verified '{title}' is in gradebook group '{name}'.")
    else:
        log(f"WARNING: Canvas didn't confirm that '{title}' moved into gradebook group '{name}'. "
            "Check it on the Assignments page.")


def sort_existing_quizzes() -> None:
    """Move quizzes this program uploaded out of "Imported Assignments" into their chosen groups.

    Only quizzes in this course's upload history are looked at (by their saved assignment IDs),
    and only ones still in "Imported Assignments" are moved. Runs as a preview first, like every run.
    """
    global STATE
    if not API_ROOT:
        raise CanvasError("Not connected to Canvas yet. Fill in the Course setup tab first.")
    STATE = read_state()
    for key in COUNTS:
        COUNTS[key] = 0
    mode = "PREVIEW (nothing in Canvas will change)" if DRY_RUN else "LIVE (Canvas will be changed)"
    log(f"Starting | sort existing quizzes | course {COURSE_ID} | {mode}")
    validate = api_request("GET", f"/courses/{COURSE_ID}")
    log(f"Connected to Canvas course {COURSE_ID}: {validate.get('name', '') if isinstance(validate, dict) else ''}")
    resolve_group_targets(set(QUIZ_KINDS))
    if not GROUP_TARGETS:
        log("No gradebook groups are chosen for this course, so there is nothing to sort. "
            "Choose them on the Course setup tab first.")
        return
    prefix = f"{COURSE_ID}:"
    found: dict[str, tuple[str, str]] = {}   # assignment ID -> (kind, file name)
    for key, saved in STATE.items():
        if not key.startswith(prefix) or not isinstance(saved, dict) or not saved.get("assignment_id"):
            continue
        folder_text, _, file_name = key[len(prefix):].rpartition(":")
        if not file_name.lower().endswith(".zip"):
            continue
        kind = assessment_kind(Path(file_name), Path(folder_text))
        if kind in GROUP_TARGETS:
            found[str(saved["assignment_id"])] = (kind, file_name)
    log(f"Quizzes in this course's upload history with a chosen group: {len(found)}")
    for assignment_id, (kind, file_name) in found.items():
        assignment = api_request("GET", f"/courses/{COURSE_ID}/assignments/{assignment_id}", allow_404=True)
        if not isinstance(assignment, dict) or not assignment.get("id"):
            log(f"Not in Canvas any more (assignment {assignment_id}, from {file_name}); nothing to move.")
            continue
        title = assignment_name(assignment)
        target = GROUP_TARGETS[kind]
        current_id = str(assignment.get("assignment_group_id"))
        current = _GROUP_NAMES.get(current_id, f"group ID {current_id}")
        if current_id == str(target.get("id")):
            log(f"Already in '{target.get('name')}': {title}")
            continue
        if norm(current) != norm(IMPORTED_GROUP_NAME):
            log(f"Left alone (it's in '{current}', which someone chose): {title}")
            continue
        changes = planned_group_change(assignment, kind, title, new_import=False)
        if changes and not DRY_RUN:
            updated = api_request("PUT", f"/courses/{COURSE_ID}/assignments/{assignment_id}",
                                  json_body={"assignment": changes})
            confirm_group_change(updated, changes, title)
    log("Finished.")


# -------------------- spreadsheet-based Canvas settings (unchanged) --------------------

def _cell_text(value: Any) -> str:
    """Convert spreadsheet values to trimmed text without turning IDs into '1075.0'."""
    if value is None:
        return ""
    try:
        if pd is not None and pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _yes(value: Any) -> bool:
    return _cell_text(value).upper() == "YES"


def _blank(value: Any) -> bool:
    return _cell_text(value) == ""


def _course_id_from_row(value: Any) -> str:
    return _cell_text(value) or str(COURSE_ID)


def _normalize_item_type(value: Any) -> str:
    raw = re.sub(r"[\s_-]+", "", _cell_text(value)).lower()
    aliases = {
        "subheader": "SubHeader",
        "textheader": "SubHeader",
        "link": "ExternalUrl",
        "contextmodulesubheader": "SubHeader",
        "contextmoduleheader": "SubHeader",
        "externalurl": "ExternalUrl",
        "page": "Page",
        "attachment": "File",
        "file": "File",
        "assignment": "Assignment",
        "newquiz": "Assignment",  # Canvas New Quizzes appear as Assignment module items.
        "quiz": "Quiz",
        "discussion": "Discussion",
        "externaltool": "ExternalTool",
    }
    return aliases.get(raw, _cell_text(value))


def _module_cache(course_id: str, cache: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    if course_id not in cache:
        modules = get_all_pages(f"/courses/{course_id}/modules?per_page=100")
        by_name: dict[str, dict[str, Any]] = {}
        for module in modules:
            name = _cell_text(module.get("name"))
            if name:
                by_name.setdefault(name, {"duplicates": []})
                if "module" not in by_name[name]:
                    by_name[name]["module"] = module
                else:
                    by_name[name]["duplicates"].append(module)
        cache[course_id] = (modules, by_name)
    return cache[course_id]


def _get_exact_module(course_id: str, name: Any, module_cache: dict[str, Any]) -> dict[str, Any] | None:
    """Find a module by exact name first, then by a confident unique normalized match."""
    module_name = _cell_text(name)
    if not module_name:
        return None
    modules, by_name = _module_cache(course_id, module_cache)
    entry = by_name.get(module_name)
    if entry:
        if entry.get("duplicates"):
            raise CanvasError(f"More than one module is named '{module_name}' in course {course_id}; refusing to guess.")
        return entry["module"]
    normalized_matches = [m for m in modules if norm(_cell_text(m.get("name"))) == norm(module_name)]
    if len(normalized_matches) == 1:
        log(f"Module name variation matched: spreadsheet '{module_name}' -> Canvas '{normalized_matches[0].get('name')}'.")
        return normalized_matches[0]
    if len(normalized_matches) > 1:
        raise CanvasError(f"Multiple modules match '{module_name}' after normalization in course {course_id}; refusing to guess.")
    matched, reason = _unique_best_match(module_name, modules, lambda m: m.get("name"), threshold=0.86, margin=0.10)
    if matched:
        log(f"Module name variation matched ({reason}): spreadsheet '{module_name}' -> Canvas '{matched.get('name')}'.")
        return matched
    if reason == "ambiguous match":
        raise CanvasError(f"Module name '{module_name}' has multiple plausible matches in course {course_id}; refusing to guess.")
    return None


def _module_items(course_id: str, module_id: Any, item_cache: dict[tuple[str, str], list[dict[str, Any]]]) -> list[dict[str, Any]]:
    key = (str(course_id), str(module_id))
    if key not in item_cache:
        item_cache[key] = get_all_pages(
            f"/courses/{course_id}/modules/{module_id}/items?include[]=content_details&per_page=100"
        )
    return item_cache[key]


def _find_exact_module_item(course_id: str, module: dict[str, Any], title: Any,
                            item_type: Any, item_cache: dict[tuple[str, str], list[dict[str, Any]]]) -> dict[str, Any] | None:
    """Find one module item by normalized title and optional normalized Canvas type.

    Emoji, case, punctuation, and spacing differences are ignored. If normalization
    makes multiple items match, raise CanvasError rather than guessing.
    """
    wanted_title = _cell_text(title)
    wanted_type = _normalize_item_type(item_type)
    if not wanted_title:
        return None
    eligible = [
        item for item in _module_items(course_id, module["id"], item_cache)
        if (not wanted_type or _normalize_item_type(item.get("type")) == wanted_type)
    ]
    exact_matches = [item for item in eligible if norm(_cell_text(item.get("title"))) == norm(wanted_title)]
    if len(exact_matches) == 1:
        return exact_matches[0]
    if len(exact_matches) > 1:
        raise CanvasError(
            f"More than one module item matches '{wanted_title}' ({wanted_type}) in module '{module.get('name')}'; refusing to guess."
        )
    matched, reason = _unique_best_match(wanted_title, eligible, lambda item: item.get("title"), threshold=0.82, margin=0.08)
    if matched:
        log(f"Item title variation matched ({reason}): spreadsheet '{wanted_title}' -> Canvas '{matched.get('title')}' in module '{module.get('name')}'.")
        return matched
    if reason == "ambiguous match":
        raise CanvasError(f"Item title '{wanted_title}' has multiple plausible matches in module '{module.get('name')}'; refusing to guess.")
    return None


def _parse_canvas_datetime(value: Any, timezone_name: str) -> str | None:
    raw = _cell_text(value)
    if not raw:
        return None
    if raw.upper() == "CLEAR":
        return None
    try:
        parsed = pd.to_datetime(value, errors="raise")
        if hasattr(parsed, "to_pydatetime"):
            parsed = parsed.to_pydatetime()
        if not isinstance(parsed, datetime):
            raise ValueError(f"not a date/time value: {value!r}")
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=ZoneInfo(timezone_name))
        return parsed.isoformat()
    except Exception as exc:
        raise ValueError(f"Could not interpret date/time value {value!r}: {exc}") from exc


def _get_course_timezone(course_id: str, cache: dict[str, str]) -> str:
    if course_id not in cache:
        course = api_request("GET", f"/courses/{course_id}")
        cache[course_id] = course.get("time_zone") or "UTC"
    return cache[course_id]


def _same_setting_value(key: str, expected: Any, actual: Any) -> bool:
    if key in {"due_at", "unlock_at", "lock_at"}:
        if expected is None or actual is None:
            return expected is actual
        try:
            left = pd.to_datetime(expected, utc=True)
            right = pd.to_datetime(actual, utc=True)
            return left == right
        except Exception:
            return str(expected) == str(actual)
    if key == "points_possible":
        try:
            return float(expected) == float(actual)
        except (TypeError, ValueError):
            return expected == actual
    return expected == actual


def _log_setting_action(message: str) -> None:
    prefix = "SETTINGS DRY RUN: would " if SETTINGS_DRY_RUN else "SPREADSHEET SETTINGS: "
    log(prefix + message)


def _apply_assignment_rows(frame: "pd.DataFrame", module_cache: dict[str, Any],
                           timezone_cache: dict[str, str]) -> None:
    required = {"Apply changes", "Course ID", "Assignment name (exact)"}
    missing = required - set(frame.columns)
    if missing:
        raise CanvasError(f"Assignments tab is missing required column(s): {', '.join(sorted(missing))}")

    # Assignment Group is optional for older workbooks, but when present it moves
    # the existing Canvas assignment into the named group (it never creates a duplicate).
    group_cache: dict[str, dict[str, dict[str, Any]]] = {}
    assignment_cache: dict[str, list[dict[str, Any]]] = {}
    for row_num, row in frame.iterrows():
        if not _yes(row.get("Apply changes")):
            continue
        course_id = _course_id_from_row(row.get("Course ID"))
        name = _cell_text(row.get("Assignment name (exact)"))
        if not name:
            log(f"Assignments row {row_num + 2}: skipped — assignment name is blank.")
            continue
        if course_id not in assignment_cache:
            assignment_cache[course_id] = get_all_pages(f"/courses/{course_id}/assignments?per_page=100")
        assignments = assignment_cache[course_id]
        # Exact normalized match first; otherwise allow a conservative, unique fuzzy match.
        # Similar names are never guessed when the top candidates are too close.
        wanted_name_key = norm(name)
        matches = [a for a in assignments if norm(_cell_text(a.get("name"))) == wanted_name_key]
        if len(matches) == 1:
            assignment = matches[0]
        elif len(matches) > 1:
            log(f"Assignments row {row_num + 2}: skipped — normalized assignment name '{name}' matches {len(matches)} items in course {course_id}; review required.")
            continue
        else:
            assignment, reason = _unique_best_match(name, assignments, lambda a: a.get("name"), threshold=0.82, margin=0.08)
            if not assignment:
                log(f"Assignments row {row_num + 2}: skipped — no unique confident match for '{name}' in course {course_id} ({reason}); review the Canvas title and spreadsheet row.")
                continue
            log(f"Assignments row {row_num + 2}: title variation matched ({reason}): spreadsheet '{name}' -> Canvas '{assignment.get('name')}' (ID {assignment.get('id')}).")
        changes: dict[str, Any] = {}

        group_name = _cell_text(row.get("Assignment Group"))
        if group_name:
            if course_id not in group_cache:
                raw_groups = get_all_pages(f"/courses/{course_id}/assignment_groups?per_page=100")
                group_cache[course_id] = {
                    norm(_cell_text(group.get("name"))): group
                    for group in raw_groups
                    if _cell_text(group.get("name"))
                }
            target_group = group_cache[course_id].get(norm(group_name))
            if not target_group:
                available = ", ".join(sorted(str(g.get("name")) for g in group_cache[course_id].values()))
                log(f"Assignments row {row_num + 2}: skipped — Assignment Group '{group_name}' was not found in Canvas course {course_id}. Available groups: {available or '(none)'}.")
                continue
            target_group_id = target_group.get("id")
            if target_group_id is None:
                log(f"Assignments row {row_num + 2}: skipped — Canvas did not return an ID for assignment group '{group_name}'.")
                continue
            current_group_id = assignment.get("assignment_group_id")
            if str(current_group_id) != str(target_group_id):
                current_group = next((g for g in group_cache[course_id].values()
                                      if str(g.get("id")) == str(current_group_id)), None)
                current_group_name = current_group.get("name") if current_group else f"group ID {current_group_id}"
                log(f"Assignments row {row_num + 2}: assignment group change planned: '{assignment.get('name')}' — '{current_group_name}' -> '{target_group.get('name')}'.")
                changes["assignment_group_id"] = int(target_group_id)
            else:
                log(f"Assignments row {row_num + 2}: assignment '{assignment.get('name')}' is already in assignment group '{target_group.get('name')}'.")

        timezone_name = _get_course_timezone(course_id, timezone_cache)
        for col, api_field in (("Due date", "due_at"), ("Unlock date", "unlock_at"), ("Lock date", "lock_at")):
            value = row.get(col)
            if not _blank(value):
                raw = _cell_text(value)
                changes[api_field] = None if raw.upper() == "CLEAR" else _parse_canvas_datetime(value, timezone_name)

        points = row.get("Points possible")
        if not _blank(points):
            try:
                changes["points_possible"] = float(points)
            except (TypeError, ValueError):
                log(f"Assignments row {row_num + 2}: skipped — Points possible is not numeric.")
                continue

        published = _cell_text(row.get("Published"))
        if published:
            if published.upper() not in ("YES", "NO"):
                log(f"Assignments row {row_num + 2}: skipped — Published must be YES or NO.")
                continue
            changes["published"] = published.upper() == "YES"

        if not _blank(row.get("Available to")):
            log(f"Assignments row {row_num + 2}: NOTE — 'Available to' is not automated by this script and will be left unchanged.")

        # Only send settings that are actually different from Canvas right now.
        changes = {key: value for key, value in changes.items()
                   if key == "assignment_group_id" or not _same_setting_value(key, value, assignment.get(key))}

        if not changes:
            log(f"Assignments row {row_num + 2}: no update needed — no supported settings differ from the current Canvas values.")
            continue

        detail = ", ".join(f"{key}={value!r}" for key, value in changes.items())
        _log_setting_action(f"update assignment {assignment.get('id')} '{name}' in course {course_id}: {detail}")
        if not SETTINGS_DRY_RUN:
            api_request("PUT", f"/courses/{course_id}/assignments/{assignment['id']}",
                        json_body={"assignment": changes})
            verified = api_request("GET", f"/courses/{course_id}/assignments/{assignment['id']}")
            failed = [key for key, value in changes.items() if not _same_setting_value(key, value, verified.get(key))]
            if failed:
                log(f"WARNING: assignment '{name}' update returned, but verification differs for: {', '.join(failed)}")
            else:
                log(f"Verified assignment update: '{name}' (ID {assignment['id']}).")


def _apply_module_item_rows(frame: "pd.DataFrame", module_cache: dict[str, Any],
                            item_cache: dict[tuple[str, str], list[dict[str, Any]]]) -> None:
    required = {"Apply changes", "Course ID", "Current module name", "Current item title", "Canvas Item Type", "Action"}
    missing = required - set(frame.columns)
    if missing:
        raise CanvasError(f"Module Items tab is missing required column(s): {', '.join(sorted(missing))}")

    has_published_column = "Published" in frame.columns

    for row_num, row in frame.iterrows():
        if not _yes(row.get("Apply changes")):
            continue
        course_id = _course_id_from_row(row.get("Course ID"))
        module_name = _cell_text(row.get("Current module name"))
        title = _cell_text(row.get("Current item title"))
        item_type = _cell_text(row.get("Canvas Item Type"))
        actions = [a.strip().upper() for a in re.split(r"[,;+]+", _cell_text(row.get("Action"))) if a.strip()]
        published_raw = _cell_text(row.get("Published")) if has_published_column else ""
        published_value: bool | None = None

        if published_raw:
            if published_raw.upper() not in ("YES", "NO"):
                log(f"Module Items row {row_num + 2}: skipped — Published must be YES, NO, or blank.")
                continue
            # The Canvas Modules API supports publishing/unpublishing module items
            # of multiple types, including SubHeader and ExternalUrl.
            published_value = published_raw.upper() == "YES"

        if not actions and published_value is None:
            log(f"Module Items row {row_num + 2}: skipped — Action and Published are both blank.")
            continue
        unsupported = [a for a in actions if a not in {"RENAME", "MOVE", "REORDER"}]
        if unsupported:
            log(f"Module Items row {row_num + 2}: skipped — unsupported action(s): {', '.join(unsupported)}. DELETE is intentionally not supported.")
            continue
        try:
            module = _get_exact_module(course_id, module_name, module_cache)
        except CanvasError as exc:
            log(f"Module Items row {row_num + 2}: skipped — {exc}")
            continue
        if not module:
            log(f"Module Items row {row_num + 2}: skipped — module '{module_name}' not found in course {course_id}.")
            continue
        try:
            item = _find_exact_module_item(course_id, module, title, item_type, item_cache)
        except CanvasError as exc:
            log(f"Module Items row {row_num + 2}: skipped — {exc}")
            continue
        if not item:
            log(f"Module Items row {row_num + 2}: skipped — item '{title}' ({item_type}) not found in module '{module_name}'.")
            continue

        # Published is the module item's own published state. Use the Modules API
        # for all item types rather than the Pages API, which only handles Pages.
        # Keep answer-key items protected regardless of their Canvas item type.
        if published_value is not None:
            if re.search(r"answer[ _-]*key", title, re.IGNORECASE):
                log(f"Module Items row {row_num + 2}: SAFETY SKIP — item title '{title}' looks like an answer key; it will not be published or unpublished by this row.")
            else:
                current_published = item.get("published")
                if current_published is published_value:
                    log(f"Module Items row {row_num + 2}: module-item publication already matches Published={published_raw.upper()}: '{title}' ({item_type}).")
                else:
                    action_word = "publish" if published_value else "unpublish"
                    _log_setting_action(f"{action_word} module item '{title}' ({item_type}) in module '{module_name}', course {course_id}")
                    if not SETTINGS_DRY_RUN:
                        item_endpoint = f"/courses/{course_id}/modules/{module['id']}/items/{item['id']}"
                        updated_item = api_request("PUT", item_endpoint,
                                                   json_body={"module_item": {"published": published_value}})
                        verified_item = api_request("GET", item_endpoint)
                        verified_value = verified_item.get("published") if isinstance(verified_item, dict) else None
                        if verified_value is published_value:
                            log(f"Verified module-item publication update: '{title}' ({item_type}) -> {published_value}.")
                        else:
                            log(f"WARNING: module item '{title}' ({item_type}) publication update returned, but verification shows published={verified_value!r}.")
                        if isinstance(updated_item, dict):
                            item["published"] = updated_item.get("published", published_value)
                        else:
                            item["published"] = published_value
                    else:
                        item["published"] = published_value

        module_item_changes: dict[str, Any] = {}
        new_title = _cell_text(row.get("New Name Title"))
        destination_name = _cell_text(row.get("Destination Module"))
        position_value = row.get("New Position")
        if "RENAME" in actions:
            if not new_title:
                log(f"Module Items row {row_num + 2}: skipped module-item changes — RENAME requires New Name Title.")
                continue
            module_item_changes["title"] = new_title
        if "MOVE" in actions:
            if not destination_name:
                log(f"Module Items row {row_num + 2}: skipped module-item changes — MOVE requires Destination Module.")
                continue
            try:
                destination = _get_exact_module(course_id, destination_name, module_cache)
            except CanvasError as exc:
                log(f"Module Items row {row_num + 2}: skipped module-item changes — {exc}")
                continue
            if not destination:
                log(f"Module Items row {row_num + 2}: skipped module-item changes — destination module '{destination_name}' not found.")
                continue
            module_item_changes["module_id"] = destination["id"]
        if "REORDER" in actions or (("MOVE" in actions) and not _blank(position_value)):
            try:
                position = int(float(position_value))
                if position < 1:
                    raise ValueError
                module_item_changes["position"] = position
            except (TypeError, ValueError):
                log(f"Module Items row {row_num + 2}: skipped module-item changes — REORDER requires New Position as a whole number of 1 or greater.")
                continue

        if not module_item_changes:
            if actions:
                log(f"Module Items row {row_num + 2}: no module-item rename/move/reorder changes supplied; publication setting (if any) was handled separately.")
            continue

        endpoint = f"/courses/{course_id}/modules/{module['id']}/items/{item['id']}"
        _log_setting_action(f"update module item {item['id']} '{title}' in '{module_name}' with {module_item_changes}")
        if not SETTINGS_DRY_RUN:
            api_request("PUT", endpoint, json_body={"module_item": module_item_changes})
            verify_module_id = module_item_changes.get("module_id", module["id"])
            verified_items = get_all_pages(f"/courses/{course_id}/modules/{verify_module_id}/items?per_page=100")
            verified = next((x for x in verified_items if str(x.get("id")) == str(item["id"])), None)
            if not verified:
                log(f"WARNING: could not verify module item {item['id']} after update.")
            else:
                ok = True
                if "title" in module_item_changes and verified.get("title") != module_item_changes["title"]:
                    ok = False
                if "position" in module_item_changes and int(verified.get("position", 0)) != module_item_changes["position"]:
                    ok = False
                if ok:
                    log(f"Verified module item update: ID {item['id']}.")
                else:
                    log(f"WARNING: module item {item['id']} update returned, but verification differs.")


def _apply_module_requirement_rows(frame: "pd.DataFrame", module_cache: dict[str, Any],
                                   item_cache: dict[tuple[str, str], list[dict[str, Any]]]) -> None:
    required = {"Apply changes", "Course ID", "Module name (exact)", "Item title (exact)", "Requirement"}
    missing = required - set(frame.columns)
    if missing:
        raise CanvasError(f"Module Requirements tab is missing required column(s): {', '.join(sorted(missing))}")
    for row_num, row in frame.iterrows():
        if not _yes(row.get("Apply changes")):
            continue
        course_id = _course_id_from_row(row.get("Course ID"))
        module_name = _cell_text(row.get("Module name (exact)"))
        title = _cell_text(row.get("Item title (exact)"))
        requirement = _cell_text(row.get("Requirement")).lower()
        if not requirement:
            log(f"Module Requirements row {row_num + 2}: skipped — no Requirement chosen for '{title}'.")
            continue
        if requirement not in {"must_view", "must_submit", "must_contribute", "must_mark_done", "min_score"}:
            log(f"Module Requirements row {row_num + 2}: skipped — unsupported Requirement '{requirement}'.")
            continue
        valid_for_type = {
            "must_view": {"SubHeader", "ExternalUrl", "Page", "File", "Assignment", "Quiz", "Discussion", "ExternalTool"},
            "must_submit": {"Assignment", "Quiz"},
            "must_contribute": {"Assignment", "Discussion", "Page"},
            "must_mark_done": {"Assignment", "Page"},
            "min_score": {"Assignment", "Quiz"},
        }
        try:
            module = _get_exact_module(course_id, module_name, module_cache)
            if not module:
                log(f"Module Requirements row {row_num + 2}: skipped — module '{module_name}' not found.")
                continue
            # A lesson heading and its page can have the same title once emoji and
            # punctuation are normalized. Choose among items that support the
            # requested requirement, preferring the actual content item over a
            # SubHeader for must_view when both are present.
            all_module_items = _module_items(course_id, module["id"], item_cache)
            title_matches = [
                candidate for candidate in all_module_items
                if norm(_cell_text(candidate.get("title"))) == norm(title)
            ]
            if not title_matches:
                # Use fuzzy matching only among item types that can support the requested requirement.
                eligible_for_requirement = [
                    candidate for candidate in all_module_items
                    if _normalize_item_type(candidate.get("type")) in valid_for_type[requirement]
                ]
                fuzzy_item, fuzzy_reason = _unique_best_match(
                    title, eligible_for_requirement, lambda candidate: candidate.get("title"),
                    threshold=0.82, margin=0.08,
                )
                if fuzzy_item:
                    title_matches = [fuzzy_item]
                    log(f"Module Requirements row {row_num + 2}: title variation matched ({fuzzy_reason}): spreadsheet '{title}' -> Canvas '{fuzzy_item.get('title')}'.")
                elif fuzzy_reason == "ambiguous match":
                    raise CanvasError(f"Item title '{title}' has multiple plausible matches in module '{module.get('name')}'; refusing to guess.")
            valid_matches = [
                candidate for candidate in title_matches
                if _normalize_item_type(candidate.get("type")) in valid_for_type[requirement]
            ]
            if requirement == "must_view" and any(
                _normalize_item_type(candidate.get("type")) != "SubHeader" for candidate in valid_matches
            ):
                valid_matches = [
                    candidate for candidate in valid_matches
                    if _normalize_item_type(candidate.get("type")) != "SubHeader"
                ]
            if len(valid_matches) > 1:
                details = ", ".join(
                    f"ID {candidate.get('id')} ({_normalize_item_type(candidate.get('type'))})"
                    for candidate in valid_matches
                )
                raise CanvasError(
                    f"More than one eligible module item matches '{title}' in module '{module.get('name')}': {details}; refusing to guess."
                )
            item = valid_matches[0] if valid_matches else None
        except CanvasError as exc:
            log(f"Module Requirements row {row_num + 2}: skipped — {exc}")
            continue
        if not item:
            if title_matches:
                found_types = ", ".join(sorted({_normalize_item_type(candidate.get("type")) for candidate in title_matches}))
                log(f"Module Requirements row {row_num + 2}: skipped — title '{title}' matched item type(s) {found_types}, but none support '{requirement}'.")
            else:
                log(f"Module Requirements row {row_num + 2}: skipped — item '{title}' not found in '{module_name}'.")
            continue

        item_type = _normalize_item_type(item.get("type"))

        completion: dict[str, Any] = {"type": requirement}
        if requirement == "min_score":
            percent_value = row.get("Minimum score (%)")
            if _blank(percent_value):
                log(f"Module Requirements row {row_num + 2}: skipped — min_score requires Minimum score (%).")
                continue
            try:
                percent = float(percent_value)
                if not 0 <= percent <= 100:
                    raise ValueError
                details = item.get("content_details") or {}
                points_possible = details.get("points_possible")
                if points_possible is None and item_type == "Assignment" and item.get("content_id"):
                    assignment = api_request("GET", f"/courses/{course_id}/assignments/{item['content_id']}", allow_404=True)
                    if assignment:
                        points_possible = assignment.get("points_possible")
                if points_possible is None:
                    log(f"Module Requirements row {row_num + 2}: skipped — could not determine points possible to convert the percentage to a Canvas minimum score.")
                    continue
                minimum_points = int(math.ceil(round(float(points_possible) * percent / 100.0, 6)))
                completion["min_score"] = minimum_points
                log(f"Module Requirements row {row_num + 2}: {percent:g}% of {points_possible} points converts to a Canvas minimum score of {minimum_points} points.")
            except (TypeError, ValueError):
                log(f"Module Requirements row {row_num + 2}: skipped — Minimum score (%) must be between 0 and 100.")
                continue

        current = item.get("completion_requirement") or {}
        if current.get("type") == completion["type"] and (
            "min_score" not in completion or _cell_text(current.get("min_score")) == str(completion["min_score"])
        ):
            log(f"Module Requirements row {row_num + 2}: requirement already set to {completion['type']} for '{title}'.")
            continue
        endpoint = f"/courses/{course_id}/modules/{module['id']}/items/{item['id']}"
        _log_setting_action(f"set completion requirement for '{title}' in '{module_name}' to {completion}")
        if not SETTINGS_DRY_RUN:
            api_request("PUT", endpoint, json_body={"module_item": {"completion_requirement": completion}})
            verified_items = get_all_pages(f"/courses/{course_id}/modules/{module['id']}/items?include[]=content_details&per_page=100")
            verified = next((x for x in verified_items if str(x.get("id")) == str(item["id"])), None)
            actual = (verified or {}).get("completion_requirement") or {}
            if actual.get("type") == completion["type"] and (
                "min_score" not in completion or int(actual.get("min_score", -1)) == completion["min_score"]
            ):
                log(f"Verified module requirement for '{title}'.")
            else:
                log(f"WARNING: requirement update returned, but verification differs for '{title}'.")


def module_completion_mode(module: dict[str, Any]) -> str:
    """'one' when students finish a module by completing any one requirement, otherwise 'all'."""
    mode = _cell_text(module.get("requirement_type")).lower()
    if mode in ("all", "one"):
        return mode
    return "one" if _cell_text(module.get("requirement_count")) == "1" else "all"


def _apply_module_setting_rows(frame: "pd.DataFrame", module_cache: dict[str, Any]) -> None:
    required = {"Apply changes", "Course ID", "Module name (exact)"}
    missing = required - set(frame.columns)
    if missing:
        raise CanvasError(f"Module Settings tab is missing required column(s): {', '.join(sorted(missing))}")
    for row_num, row in frame.iterrows():
        if not _yes(row.get("Apply changes")):
            continue
        course_id = _course_id_from_row(row.get("Course ID"))
        module_name = _cell_text(row.get("Module name (exact)"))
        try:
            module = _get_exact_module(course_id, module_name, module_cache)
        except CanvasError as exc:
            log(f"Module Settings row {row_num + 2}: skipped — {exc}")
            continue
        if not module:
            log(f"Module Settings row {row_num + 2}: skipped — module '{module_name}' not found.")
            continue
        changes: dict[str, Any] = {}
        prerequisites_value = _cell_text(row.get("Prerequisite module names"))
        if prerequisites_value:
            if prerequisites_value.upper() == "CLEAR":
                changes["prerequisite_module_ids"] = []
            else:
                prereq_ids = []
                bad_names = []
                for prerequisite_name in [x.strip() for x in prerequisites_value.split(";") if x.strip()]:
                    try:
                        prerequisite = _get_exact_module(course_id, prerequisite_name, module_cache)
                    except CanvasError as exc:
                        bad_names.append(str(exc))
                        continue
                    if not prerequisite:
                        bad_names.append(f"module '{prerequisite_name}' not found")
                    elif int(prerequisite.get("position", 0)) >= int(module.get("position", 0)):
                        bad_names.append(f"prerequisite '{prerequisite_name}' must appear before '{module_name}'")
                    else:
                        prereq_ids.append(prerequisite["id"])
                if bad_names:
                    log(f"Module Settings row {row_num + 2}: skipped — " + "; ".join(bad_names))
                    continue
                changes["prerequisite_module_ids"] = prereq_ids

        sequential = _cell_text(row.get("Require items in order"))
        if sequential:
            if sequential.upper() not in {"YES", "NO"}:
                log(f"Module Settings row {row_num + 2}: skipped — Require items in order must be YES or NO.")
                continue
            changes["require_sequential_progress"] = sequential.upper() == "YES"

        completion_mode = _cell_text(row.get("Completion mode"))
        if completion_mode:
            if completion_mode.lower() not in {"all", "one"}:
                log(f"Module Settings row {row_num + 2}: skipped — Completion mode must be 'all' or 'one'.")
                continue
            changes["requirement_type"] = completion_mode.lower()
            changes["requirement_count"] = 1 if completion_mode.lower() == "one" else None

        # Leave out settings that already match Canvas.
        if "prerequisite_module_ids" in changes and (
            {str(x) for x in changes["prerequisite_module_ids"]}
            == {str(x) for x in (module.get("prerequisite_module_ids") or [])}
        ):
            del changes["prerequisite_module_ids"]
        if "require_sequential_progress" in changes and \
                changes["require_sequential_progress"] == bool(module.get("require_sequential_progress")):
            del changes["require_sequential_progress"]
        if "requirement_type" in changes and changes["requirement_type"] == module_completion_mode(module):
            del changes["requirement_type"]
            del changes["requirement_count"]
        if not changes and (prerequisites_value or sequential or completion_mode):
            log(f"Module Settings row {row_num + 2}: no update needed — '{module_name}' already matches.")
            continue

        if not changes:
            log(f"Module Settings row {row_num + 2}: skipped — no supported settings supplied.")
            continue
        _log_setting_action(f"update module '{module_name}' (ID {module['id']}) in course {course_id}: {changes}")
        if not SETTINGS_DRY_RUN:
            api_request("PUT", f"/courses/{course_id}/modules/{module['id']}",
                        json_body={"module": changes})
            verified = api_request("GET", f"/courses/{course_id}/modules/{module['id']}")
            failed = []
            for key, value in changes.items():
                if key == "requirement_count":
                    continue
                if key == "requirement_type":
                    actual: Any = module_completion_mode(verified)
                elif key == "prerequisite_module_ids":
                    value = {str(x) for x in value}
                    actual = {str(x) for x in (verified.get(key) or [])}
                else:
                    actual = verified.get(key)
                if actual != value:
                    failed.append(key)
            if failed:
                log(f"WARNING: module '{module_name}' update returned, but verification differs for: {', '.join(failed)}")
            else:
                log(f"Verified module settings for '{module_name}'.")


def _apply_course_setting_rows(frame: "pd.DataFrame") -> None:
    if "Apply changes" not in frame.columns:
        raise CanvasError("Course Settings tab is missing the Apply changes column.")
    for row_num, row in frame.iterrows():
        if _yes(row.get("Apply changes")):
            setting = _cell_text(row.get("Setting name"))
            log(f"Course Settings row {row_num + 2}: skipped — course setting '{setting}' is not yet enabled in the script. No change made.")


# Spreadsheet column names, matched without caring about capitals or extra spaces.
# Each tab maps the names people see to the names the update code uses.
SHEET_COLUMN_ALIASES: dict[str, str] = {
    "apply changes": "Apply changes",
    "setting type": "Setting type",
    "course id": "Course ID",
    "canvas item type": "Canvas Item Type",
    "published": "Published",
    "assignment group": "Assignment Group",
    "unlock date": "Unlock date",
    "due date": "Due date",
    "lock date": "Lock date",
    "points possible": "Points possible",
    "action": "Action",
    "new title": "New Name Title",
    "new name title": "New Name Title",
    "move to module": "Destination Module",
    "destination module": "Destination Module",
    "new position": "New Position",
    "requirement": "Requirement",
    "minimum score (%)": "Minimum score (%)",
    "prerequisite modules": "Prerequisite module names",
    "prerequisite module names": "Prerequisite module names",
    "require items in order": "Require items in order",
    "completion mode": "Completion mode",
    "available to": "Available to",
    "notes": "Notes",
}
_MODULE_NAME_HEADERS = {"module name", "module name (exact)", "module name (exact/current)"}
_ITEM_TITLE_HEADERS = {"item title", "item title (exact)", "item title (exact/current)"}


def _canonical_columns(frame: "pd.DataFrame", workbook_id: str) -> "pd.DataFrame":
    """Rename a consolidated tab's columns to the internal names and fill in the course ID."""
    content_tab = any(" ".join(str(c).lower().split()) == "canvas item type" for c in frame.columns)
    renames: dict[Any, str] = {}
    for column in frame.columns:
        key = " ".join(str(column).lower().split())
        if key in _MODULE_NAME_HEADERS:
            renames[column] = "Module name (exact/current)" if content_tab else "Module name (exact)"
        elif key in _ITEM_TITLE_HEADERS:
            renames[column] = "Item title (exact/current)" if content_tab else "Item title (exact)"
        elif key in SHEET_COLUMN_ALIASES:
            renames[column] = SHEET_COLUMN_ALIASES[key]
    frame = frame.rename(columns=renames)
    if "Course ID" not in frame.columns:
        frame["Course ID"] = None
    if workbook_id:   # The course ID on Start Here fills every row that doesn't name its own.
        frame["Course ID"] = [workbook_id if _blank(value) else value for value in frame["Course ID"]]
    return frame


def workbook_course_id(path: Path) -> str:
    """The course ID typed next to 'Canvas course ID' on the Start Here tab, or ''."""
    if openpyxl is None:
        return ""
    try:
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return ""
    try:
        if "Start Here" not in book.sheetnames:
            return ""
        for row in book["Start Here"].iter_rows(min_row=1, max_row=40, max_col=4, values_only=True):
            for index, value in enumerate(row[:-1]):
                if " ".join(_cell_text(value).lower().split()).startswith("canvas course id"):
                    found = course_id_from_text(_cell_text(row[index + 1]))
                    return found if found.isdigit() else ""
        return ""
    finally:
        book.close()


def run_spreadsheet_settings() -> None:
    """Read workbook rows marked YES and preview/apply supported settings only."""
    if pd is None:
        raise CanvasError("The settings spreadsheet needs the pandas and openpyxl helper packages. Reopen the program and choose Install.")
    if not SPREADSHEET_PATH.exists():
        raise CanvasError(f"Spreadsheet not found: {SPREADSHEET_PATH}")
    try:
        sheets = pd.read_excel(SPREADSHEET_PATH, sheet_name=None, dtype=object)
    except Exception as exc:
        raise CanvasError(f"Could not read spreadsheet {SPREADSHEET_PATH}: {exc}. If it is open in Excel, close it and try again.") from exc

    # The consolidated workbook (Start Here, Content & Assignments, Module Setup) is
    # converted into the internal row formats, so the Canvas update code is shared with
    # older five-tab workbooks. Column names are matched loosely (see SHEET_COLUMN_ALIASES).
    if any(tab in sheets for tab in SHEET_TABS):
        sheets = read_task_tabs(sheets, workbook_course_id(SPREADSHEET_PATH))
    elif {"Content & Assignments", "Module Setup"}.issubset(sheets):
        workbook_id = workbook_course_id(SPREADSHEET_PATH)
        content = _canonical_columns(sheets["Content & Assignments"].copy(), workbook_id)
        setup = _canonical_columns(sheets["Module Setup"].copy(), workbook_id)
        if "Setting type" not in content.columns or "Setting type" not in setup.columns:
            raise CanvasError("Consolidated tabs must include a 'Setting type' column.")

        def rows_of(frame: "pd.DataFrame", kind: str) -> "pd.DataFrame":
            return frame[frame["Setting type"].map(_cell_text).str.lower() == kind].copy()

        def pick(frame: "pd.DataFrame", columns: list[str]) -> "pd.DataFrame":
            for column in columns:
                if column not in frame.columns:
                    frame[column] = None   # Optional columns may be left out of the workbook.
            return frame[columns]

        sheets = {
            "Assignments": pick(rows_of(content, "assignment").rename(columns={
                "Item title (exact/current)": "Assignment name (exact)",
            }), ["Apply changes", "Course ID", "Assignment name (exact)", "Assignment Group", "Due date",
                 "Unlock date", "Lock date", "Points possible", "Published", "Available to", "Notes"]),
            "Module Items": pick(rows_of(content, "module item").rename(columns={
                "Module name (exact/current)": "Current module name",
                "Item title (exact/current)": "Current item title",
            }), ["Apply changes", "Course ID", "Current module name", "Current item title", "Canvas Item Type",
                 "Action", "Published", "New Name Title", "Destination Module", "New Position", "Notes"]),
            "Module Requirements": pick(rows_of(setup, "requirement"), [
                "Apply changes", "Course ID", "Module name (exact)", "Item title (exact)", "Requirement",
                "Minimum score (%)", "Notes"]),
            "Module Settings": pick(rows_of(setup, "module"), [
                "Apply changes", "Course ID", "Module name (exact)", "Prerequisite module names",
                "Require items in order", "Completion mode", "Notes"]),
        }
    else:
        required_sheets = {"Assignments", "Module Items", "Module Requirements", "Module Settings"}
        missing_sheets = required_sheets - set(sheets)
        if missing_sheets:
            raise CanvasError(f"Spreadsheet is missing sheet(s): {', '.join(sorted(missing_sheets))}; use the consolidated workbook or the legacy workbook format.")

    module_cache: dict[str, Any] = {}
    item_cache: dict[tuple[str, str], list[dict[str, Any]]] = {}
    timezone_cache: dict[str, str] = {}
    log(f"Starting spreadsheet settings | workbook={SPREADSHEET_PATH} | SETTINGS_DRY_RUN={SETTINGS_DRY_RUN}")
    for sheet_name, frame in sheets.items():
        if "Apply changes" in frame.columns:
            selected = sum(1 for value in frame["Apply changes"].tolist() if _yes(value))
            log(f"Spreadsheet tab/data group '{sheet_name}': {selected} row(s) marked YES.")
    _apply_assignment_rows(sheets["Assignments"], module_cache, timezone_cache)
    _apply_module_item_rows(sheets["Module Items"], module_cache, item_cache)
    _apply_module_requirement_rows(sheets["Module Requirements"], module_cache, item_cache)
    _apply_module_setting_rows(sheets["Module Settings"], module_cache)
    if SETTINGS_DRY_RUN:
        log("Spreadsheet settings preview complete. No spreadsheet-driven Canvas changes were made.")
        log("To make these changes, run again with \"Apply changes to Canvas\" selected.")
    else:
        log("Spreadsheet settings run complete. Review the log for verification warnings.")




# ===================== CREATE A SETTINGS SPREADSHEET =====================
# Reads the course from Canvas (read-only) and writes a new workbook with one tab per job:
# Dates & Points, Publish & Organize, Completion, and Module Rules. Every row starts as
# Apply changes = NO. Rows are grouped under a colored band for each module, which can be
# folded up with the +/- buttons in Excel's margin. Never overwrites an existing file.

SHEET_FONT = "Arial"
SHEET_NAVY = "1F3864"
SHEET_DATE_FORMAT = "ddd m/d/yy h:mm AM/PM"   # Shows like "Wed 10/21/26 11:59 PM".
SHEET_EXTRA_ROWS = 200                        # Dropdowns reach this far past the last row.
ANSWER_KEY_RX = re.compile(r"answer[ _-]*key", re.IGNORECASE)

# Friendly words in the spreadsheet, and the Canvas codes they stand for.
REQUIREMENT_LABELS = {"View": "must_view", "Submit": "must_submit", "Contribute": "must_contribute",
                      "Mark done": "must_mark_done", "Minimum score": "min_score"}
COMPLETION_LABELS = {"All": "all", "Any one": "one"}
TYPE_LABELS = {"SubHeader": "Text header", "ExternalUrl": "Link", "ExternalTool": "External tool"}

# Each tab: name, tab color, and columns as (heading, width, hover note, kind).
# kind: "edit" = you change it, "ref" = shown for reference, "center"/"date" = edit with that layout.
SHEET_TABS: dict[str, dict[str, Any]] = {
    "Dates & Points": {"color": "2E75B6", "freeze": "D2", "columns": [
        ("Apply changes", 10, "YES = change this assignment in Canvas on the next run. NO = skip this row.", "center"),
        ("Module", 16, "The module the assignment is in (for reference).", "ref"),
        ("Assignment", 46, "The assignment's exact Canvas name, emojis included. Don't change it here; "
                           "to rename, use Publish & Organize.", "ref"),
        ("Assignment group", 17, "The gradebook category, like Daily or Major. Pick from the list.", "edit"),
        ("Unlock date", 25, "When students can first open it. Type CLEAR to remove the date.", "date"),
        ("Due date", 25, "When it's due. Type CLEAR to remove the date.", "date"),
        ("Lock date", 25, "When it closes to students. Type CLEAR to remove the date.", "date"),
        ("Points", 9, "Points possible.", "center"),
        ("Notes", 40, "Your own notes. The program ignores this column.", "edit"),
    ]},
    "Publish & Organize": {"color": "548235", "freeze": "D2", "columns": [
        ("Apply changes", 10, "YES = change this item in Canvas on the next run. NO = skip this row.", "center"),
        ("Module", 16, "The module the item is in now.", "ref"),
        ("Item title", 46, "The item's exact Canvas title, emojis included. To rename it, use Action and New title.", "ref"),
        ("Type", 13, "What kind of item this is (for reference).", "ref"),
        ("Published", 11, "YES = students can see it. NO = hidden from students.", "center"),
        ("Action", 15, "Optional: RENAME, MOVE, or REORDER. Fill in the matching column to the right.", "center"),
        ("New title", 32, "Used with RENAME: the new title.", "edit"),
        ("Move to module", 18, "Used with MOVE: the module to move it into.", "edit"),
        ("New position", 10, "Used with REORDER (or MOVE): 1 = top of the module.", "center"),
        ("Notes", 34, "Your own notes. The program ignores this column.", "edit"),
    ]},
    "Completion": {"color": "BF8F00", "freeze": "D2", "columns": [
        ("Apply changes", 10, "YES = set this requirement in Canvas on the next run. NO = skip this row.", "center"),
        ("Module", 16, "The module the item is in (for reference).", "ref"),
        ("Item title", 46, "The item's exact Canvas title (for reference).", "ref"),
        ("Type", 13, "What kind of item this is (for reference).", "ref"),
        ("Requirement", 16, "What a student must do: View it, Submit it, Contribute (post or reply), "
                            "Mark done, or reach a Minimum score. Blank = no requirement.", "center"),
        ("Minimum score (%)", 11, "Used with Minimum score: the percentage students must reach, like 70.", "center"),
        ("Notes", 40, "Your own notes. The program ignores this column.", "edit"),
    ]},
    "Module Rules": {"color": "7030A0", "freeze": "C2", "columns": [
        ("Apply changes", 10, "YES = change this module in Canvas on the next run. NO = skip this row.", "center"),
        ("Module", 24, "The module's exact Canvas name (for reference).", "ref"),
        ("Prerequisite modules", 32, "Modules students must finish first. Separate several with a semicolon, "
                                     "like  Unit 1; Unit 2. They must come earlier in the course.", "edit"),
        ("Require items in order", 13, "YES = students must go through items in order.", "center"),
        ("Completion mode", 13, "All = finish every requirement. Any one = finish any one of them.", "center"),
        ("Notes", 40, "Your own notes. The program ignores this column.", "edit"),
    ]},
}

# How each tab's headings map to the names the update code uses.
TAB_COLUMN_MAP: dict[str, tuple[str, dict[str, str], list[str]]] = {
    "Dates & Points": ("Assignments", {
        "assignment": "Assignment name (exact)", "assignment group": "Assignment Group", "points": "Points possible",
    }, ["Apply changes", "Course ID", "Assignment name (exact)", "Assignment Group", "Due date", "Unlock date",
        "Lock date", "Points possible", "Published", "Notes"]),
    "Publish & Organize": ("Module Items", {
        "module": "Current module name", "item title": "Current item title", "type": "Canvas Item Type",
        "new title": "New Name Title", "move to module": "Destination Module",
    }, ["Apply changes", "Course ID", "Current module name", "Current item title", "Canvas Item Type", "Action",
        "Published", "New Name Title", "Destination Module", "New Position", "Notes"]),
    "Completion": ("Module Requirements", {
        "module": "Module name (exact)", "item title": "Item title (exact)",
    }, ["Apply changes", "Course ID", "Module name (exact)", "Item title (exact)", "Requirement",
        "Minimum score (%)", "Notes"]),
    "Module Rules": ("Module Settings", {
        "module": "Module name (exact)", "prerequisite modules": "Prerequisite module names",
    }, ["Apply changes", "Course ID", "Module name (exact)", "Prerequisite module names",
        "Require items in order", "Completion mode", "Notes"]),
}


def read_task_tabs(sheets: dict[str, "pd.DataFrame"], workbook_id: str) -> dict[str, "pd.DataFrame"]:
    """Turn the four task tabs into the internal tables the update code reads."""
    requirement_codes = {label.lower(): code for label, code in REQUIREMENT_LABELS.items()}
    completion_codes = {label.lower(): code for label, code in COMPLETION_LABELS.items()} | {"one": "one"}
    result: dict[str, "pd.DataFrame"] = {}
    for tab, (internal, renames, columns) in TAB_COLUMN_MAP.items():
        frame = sheets.get(tab)
        frame = pd.DataFrame() if frame is None else frame.copy()
        names: dict[Any, str] = {}
        for column in frame.columns:
            key = " ".join(str(column).lower().split())
            names[column] = renames.get(key) or SHEET_COLUMN_ALIASES.get(key) or str(column)
        frame = frame.rename(columns=names)
        for column in columns:
            if column not in frame.columns:
                frame[column] = None
        if workbook_id:
            frame["Course ID"] = [workbook_id if _blank(v) else v for v in frame["Course ID"]]
        if "Requirement" in columns:
            frame["Requirement"] = [requirement_codes.get(_cell_text(v).lower(), v) for v in frame["Requirement"]]
        if "Completion mode" in columns:
            frame["Completion mode"] = [completion_codes.get(_cell_text(v).lower(), v) for v in frame["Completion mode"]]
        result[internal] = frame[columns]
    return result


def _local_time(value: Any, timezone_name: str) -> datetime | None:
    """Canvas gives UTC times; the spreadsheet shows them in the course's time zone."""
    text = _cell_text(value)
    if not text:
        return None
    try:
        moment = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return moment.astimezone(ZoneInfo(timezone_name)).replace(tzinfo=None)
    except (ValueError, KeyError):
        return None


def _percent_text(min_score: Any, points: Any) -> float | int | None:
    """Turn a Canvas minimum score (points) back into the percentage the spreadsheet uses.
    Rounded down to 2 decimals so it converts back to exactly the same minimum score."""
    try:
        minimum, total = float(min_score), float(points)
    except (TypeError, ValueError):
        return None
    if total <= 0:
        return None
    percent = math.floor(minimum / total * 10000 + 1e-6) / 100
    return int(percent) if float(percent).is_integer() else percent


def read_course_for_sheet(course_id: str) -> dict[str, Any]:
    """Everything the spreadsheet needs, read from Canvas with GET requests only."""
    course = api_request("GET", f"/courses/{course_id}")
    log(f"Reading '{course.get('name', course_id)}' from Canvas (read-only)…")
    modules = get_all_pages(f"/courses/{course_id}/modules?include[]=items&include[]=content_details&per_page=100")
    modules.sort(key=lambda m: int(m.get("position") or 0))
    for module in modules:
        if "items" not in module:   # Canvas leaves items out of very large modules.
            module["items"] = get_all_pages(
                f"/courses/{course_id}/modules/{module['id']}/items?include[]=content_details&per_page=100")
        module["items"].sort(key=lambda i: int(i.get("position") or 0))
    assignments = get_all_pages(f"/courses/{course_id}/assignments?per_page=100")
    groups = get_all_pages(f"/courses/{course_id}/assignment_groups?per_page=100")
    groups.sort(key=lambda g: int(g.get("position") or 0))
    item_count = sum(len(m["items"]) for m in modules)
    log(f"Found {len(modules)} module(s), {item_count} module item(s), {len(assignments)} assignment(s), "
        f"{len(groups)} assignment group(s).")
    return {"course": course, "modules": modules, "assignments": assignments, "groups": groups}


def build_sheet_rows(data: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    """Rows for each tab. {'_band': name} starts a module; keys starting with '_' are styling hints."""
    timezone_name = data["course"].get("time_zone") or "UTC"
    group_names = {str(g.get("id")): _cell_text(g.get("name")) for g in data["groups"]}
    module_names = {str(m.get("id")): _cell_text(m.get("name")) for m in data["modules"]}
    assignments = data["assignments"]
    by_id = {str(a.get("id")): a for a in assignments}
    by_quiz = {str(a.get("quiz_id")): a for a in assignments if a.get("quiz_id")}
    by_topic = {str(a["discussion_topic"].get("id")): a for a in assignments
                if isinstance(a.get("discussion_topic"), dict)}
    name_counts: dict[str, int] = {}
    for a in assignments:
        key = norm(_cell_text(a.get("name")))
        name_counts[key] = name_counts.get(key, 0) + 1

    def assignment_for(item: dict[str, Any]) -> dict[str, Any] | None:
        kind, content_id = item.get("type"), str(item.get("content_id"))
        found = {"Assignment": by_id, "Quiz": by_quiz, "Discussion": by_topic}.get(kind, {}).get(content_id)
        if found is None and kind in ("Quiz", "Discussion"):
            same = [a for a in assignments if norm(_cell_text(a.get("name"))) == norm(_cell_text(item.get("title")))]
            found = same[0] if len(same) == 1 else None
        return found

    def dates_row(a: dict[str, Any], module_name: str) -> dict[str, Any]:
        name = _cell_text(a.get("name"))
        note = ("Another assignment has this same name, so the program will skip it rather than guess. "
                "Rename one of them in Canvas.") if name_counts.get(norm(name), 0) > 1 else ""
        return {"Apply changes": "NO", "Module": module_name, "Assignment": name,
                "Assignment group": group_names.get(str(a.get("assignment_group_id")), ""),
                "Unlock date": _local_time(a.get("unlock_at"), timezone_name),
                "Due date": _local_time(a.get("due_at"), timezone_name),
                "Lock date": _local_time(a.get("lock_at"), timezone_name),
                "Points": a.get("points_possible"), "Notes": note}

    tabs: dict[str, list[dict[str, Any]]] = {name: [] for name in SHEET_TABS}
    used: set[str] = set()
    for module in data["modules"]:
        module_name = _cell_text(module.get("name"))
        hidden = "" if module.get("published") else "   (unpublished)"
        items = module.get("items") or []
        band = {"_band": module_name + hidden}
        tabs["Publish & Organize"].append(band)
        tabs["Completion"].append(band)
        dates_start = len(tabs["Dates & Points"])
        tabs["Dates & Points"].append(band)
        for item in items:
            title = _cell_text(item.get("title"))
            kind = _cell_text(item.get("type"))
            is_key = bool(ANSWER_KEY_RX.search(title))
            tabs["Publish & Organize"].append({
                "Apply changes": "NO", "Module": module_name, "Item title": title,
                "Type": TYPE_LABELS.get(kind, kind), "Published": "YES" if item.get("published") else "NO",
                "Notes": "🔒 Answer key: always stays hidden from students." if is_key else "",
                "_lock": {"Published"} if is_key else set(),
            })
            graded = assignment_for(item)
            if graded is not None and str(graded.get("id")) not in used:
                used.add(str(graded.get("id")))
                tabs["Dates & Points"].append(dates_row(graded, module_name))
            if kind == "SubHeader":
                continue   # Text headers can't have completion requirements.
            requirement = item.get("completion_requirement") or {}
            code = _cell_text(requirement.get("type"))
            label = next((k for k, v in REQUIREMENT_LABELS.items() if v == code), code)
            percent, note = None, ""
            if code == "min_score":
                points = (item.get("content_details") or {}).get("points_possible")
                if points is None and graded is not None:
                    points = graded.get("points_possible")
                percent = _percent_text(requirement.get("min_score"), points)
                if percent is None:
                    note = f"Canvas minimum score is {requirement.get('min_score')} point(s); enter it as a percentage."
            tabs["Completion"].append({
                "Apply changes": "NO", "Module": module_name, "Item title": title,
                "Type": TYPE_LABELS.get(kind, kind), "Requirement": label,
                "Minimum score (%)": percent, "Notes": note,
            })
        if len(tabs["Dates & Points"]) == dates_start + 1:
            tabs["Dates & Points"].pop()   # No graded items in this module: leave its band out.
        prerequisites = [module_names.get(str(x), "") for x in (module.get("prerequisite_module_ids") or [])]
        has_requirements = any(i.get("completion_requirement") for i in items)
        mode = module_completion_mode(module)
        tabs["Module Rules"].append({
            "Apply changes": "NO", "Module": module_name,
            "Prerequisite modules": "; ".join(p for p in prerequisites if p),
            "Require items in order": "YES" if module.get("require_sequential_progress") else "NO",
            "Completion mode": next(k for k, v in COMPLETION_LABELS.items() if v == mode) if has_requirements else "",
            "Notes": "" if module.get("published") else "This module is unpublished (hidden from students).",
        })
    leftovers = [a for a in assignments if str(a.get("id")) not in used]
    if leftovers:
        tabs["Dates & Points"].append({"_band": "Not in any module"})
        for a in sorted(leftovers, key=lambda x: natural_key(_cell_text(x.get("name")))):
            tabs["Dates & Points"].append(dates_row(a, ""))
    return tabs


def _example_rows() -> dict[str, list[dict[str, Any]]]:
    """A small worked example for the blank template (every row marked NO)."""
    note = "Example. Replace or delete."
    return {
        "Dates & Points": [
            {"_band": "Unit 1"},
            {"Apply changes": "NO", "Module": "Unit 1", "Assignment": "📝Lesson 1 Checkpoint",
             "Assignment group": "Daily", "Unlock date": datetime(2026, 8, 17, 8, 0),
             "Due date": datetime(2026, 8, 21, 23, 59), "Points": 10, "Notes": note},
        ],
        "Publish & Organize": [
            {"_band": "Unit 1"},
            {"Apply changes": "NO", "Module": "Unit 1", "Item title": "📚Lesson 1 Welcome", "Type": "Page",
             "Published": "YES", "Notes": note},
            {"Apply changes": "NO", "Module": "Unit 1", "Item title": "📝Lesson 1 Checkpoint", "Type": "Assignment",
             "Published": "NO", "Notes": note},
        ],
        "Completion": [
            {"_band": "Unit 1"},
            {"Apply changes": "NO", "Module": "Unit 1", "Item title": "📚Lesson 1 Welcome", "Type": "Page",
             "Requirement": "View", "Notes": note},
            {"Apply changes": "NO", "Module": "Unit 1", "Item title": "📝Lesson 1 Checkpoint", "Type": "Assignment",
             "Requirement": "Minimum score", "Minimum score (%)": 70, "Notes": note},
        ],
        "Module Rules": [
            {"Apply changes": "NO", "Module": "Unit 1", "Require items in order": "NO", "Completion mode": "All",
             "Notes": note},
            {"Apply changes": "NO", "Module": "Unit 2", "Prerequisite modules": "Unit 1",
             "Require items in order": "NO", "Completion mode": "All", "Notes": note},
        ],
    }


def _write_start_here(sheet: Any, course_name: str, course_id: str, made_line: str) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    navy, gray_text = SHEET_NAVY, "595959"

    def fill(color: str) -> PatternFill:
        return PatternFill("solid", fgColor=color)

    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = SHEET_NAVY
    sheet.column_dimensions["A"].width = 3
    sheet.column_dimensions["B"].width = 24
    sheet.column_dimensions["C"].width = 88
    row = 1

    def line(label: str, text: str, *, label_fill: str = "", label_font: Font | None = None,
             text_fill: str = "", height: float = 21) -> None:
        nonlocal row
        left = sheet.cell(row=row, column=2, value=label or None)
        left.font = label_font or Font(name=SHEET_FONT, size=11, bold=True, color=navy)
        left.alignment = Alignment(vertical="center", horizontal="left", indent=1, wrap_text=True)
        right = sheet.cell(row=row, column=3, value=text or None)
        right.font = Font(name=SHEET_FONT, size=11, color="262626")
        right.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if label_fill:
            left.fill = fill(label_fill)
        if text_fill:
            right.fill = fill(text_fill)
        sheet.row_dimensions[row].height = height
        row += 1

    def space(height: float = 12) -> None:
        nonlocal row
        sheet.row_dimensions[row].height = height
        row += 1

    def section(title: str) -> None:
        nonlocal row
        space(16)
        sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        cell = sheet.cell(row=row, column=2, value=title)
        cell.font = Font(name=SHEET_FONT, size=12, bold=True, color="FFFFFF")
        cell.fill = fill(navy)
        cell.alignment = Alignment(vertical="center", indent=1)
        sheet.row_dimensions[row].height = 26
        row += 1
        space(6)

    title = sheet.cell(row=row, column=2, value="Canvas Settings Spreadsheet")
    title.font = Font(name=SHEET_FONT, size=20, bold=True, color=navy)
    sheet.row_dimensions[row].height = 34
    row += 1
    sub = sheet.cell(row=row, column=2, value="Change your Canvas course settings from one place.")
    sub.font = Font(name=SHEET_FONT, size=11, italic=True, color=gray_text)
    row += 1
    space(10)
    line("Course", course_name or "(type your course name)", text_fill="F2F2F2")
    line("Canvas course ID", course_id or "(type the number after /courses/ in your course's web address)",
         text_fill="FFF2CC")
    sheet.cell(row=row - 1, column=3).font = Font(name=SHEET_FONT, size=12, bold=True, color="262626")
    line("Created", made_line, text_fill="F2F2F2")

    section("HOW IT WORKS")
    big = Font(name=SHEET_FONT, size=14, bold=True, color="2E75B6")
    for number, text in enumerate((
            "Find the row you want and type the new setting.",
            "Change Apply changes to YES on that row. Rows marked NO are skipped.",
            "Save and close this file.",
            "In Canvas Course Manager, run a Preview. Nothing changes yet; the blue lines show what would.",
            "Happy with it? Run again with \"Apply changes to Canvas\" and type APPLY."), start=1):
        line(str(number), text, label_font=big, height=24)
        sheet.cell(row=row - 1, column=2).alignment = Alignment(horizontal="center", vertical="center")

    section("WHICH TAB DO I USE?")
    for tab, text in (("Dates & Points", "Due, unlock, and lock dates; points; assignment group."),
                      ("Publish & Organize", "Show or hide items; rename, move, or reorder them."),
                      ("Completion", "What students must do to complete each item."),
                      ("Module Rules", "Prerequisites, items in order, and finish all or any one.")):
        line(tab, text, label_fill=SHEET_TABS[tab]["color"], height=24,
             label_font=Font(name=SHEET_FONT, size=11, bold=True, color="FFFFFF"))
        space(4)

    section("WHAT THE COLORS MEAN")
    for swatch, text in (("E2EFDA", "This row will be applied (Apply changes is YES)."),
                         ("FCE4E4", "Hidden from students (Published is NO)."),
                         ("DCE6F2", "A module. Click − or + in the left margin to fold it up or open it."),
                         ("E7E6E6", "Locked. Answer keys always stay hidden from students.")):
        line("", text, label_fill=swatch, height=24)
        space(4)

    section("GOOD TO KNOW")
    for label, text in (
            ("Titles", "Must match Canvas, emojis included. Look-alike titles are skipped, never guessed."),
            ("Dates", "Use your course's time zone. Type CLEAR in a date cell to remove that date."),
            ("Deleting", "Nothing is ever deleted. To hide something, set Published to NO."),
            ("Not in here", "\"Assign to\" dates for certain sections or students, module unlock dates, and "
                            "publishing assignments that aren't in a module. Set those in Canvas."),
            ("Help", "Point at any column heading for a quick explanation."),
            ("New copy", "Click \"Create settings spreadsheet\" again anytime. It never replaces this file.")):
        line(label, text, height=36 if len(text) > 90 else 22)


def write_settings_workbook(path: Path, *, course_name: str, course_id: str, made_line: str,
                            tabs: dict[str, list[dict[str, Any]]], groups: list[str], modules: list[str]) -> None:
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.properties import Outline

    def fill(color: str) -> PatternFill:
        return PatternFill("solid", fgColor=color)

    edit_font = Font(name=SHEET_FONT, size=11, color="262626")
    ref_font = Font(name=SHEET_FONT, size=11, color="595959")
    header_font = Font(name=SHEET_FONT, size=11, bold=True, color="FFFFFF")
    band_font = Font(name=SHEET_FONT, size=12, bold=True, color=SHEET_NAVY)
    line_side = Side(style="thin", color="D9DEE5")
    row_border = Border(bottom=line_side)
    band_border = Border(top=Side(style="thin", color="8EA9DB"), bottom=Side(style="thin", color="8EA9DB"))
    zebra = ("FFFFFF", "F6F8FB")

    book = openpyxl.Workbook()
    _write_start_here(book.active, course_name, course_id, made_line)
    book.active.title = "Start Here"

    lists = book.create_sheet("Lists")
    lists.sheet_state = "hidden"
    for r, name in enumerate(groups, start=1):
        lists.cell(row=r, column=1, value=name)
    for r, name in enumerate(modules, start=1):
        lists.cell(row=r, column=2, value=name)

    yes_no = '"YES,NO"'
    dropdowns: dict[str, dict[str, str]] = {
        "Dates & Points": {"Apply changes": yes_no},
        "Publish & Organize": {"Apply changes": yes_no, "Published": yes_no,
                               "Action": '"RENAME,MOVE,REORDER,MOVE + REORDER"'},
        "Completion": {"Apply changes": yes_no, "Requirement": '"' + ",".join(REQUIREMENT_LABELS) + '"'},
        "Module Rules": {"Apply changes": yes_no, "Require items in order": yes_no,
                         "Completion mode": '"' + ",".join(COMPLETION_LABELS) + '"'},
    }
    if groups:
        dropdowns["Dates & Points"]["Assignment group"] = f"=Lists!$A$1:$A${len(groups)}"
    if modules:
        dropdowns["Publish & Organize"]["Move to module"] = f"=Lists!$B$1:$B${len(modules)}"

    for tab, spec in SHEET_TABS.items():
        sheet = book.create_sheet(tab)
        sheet.sheet_properties.tabColor = spec["color"]
        sheet.sheet_properties.outlinePr = Outline(summaryBelow=False)
        sheet.sheet_view.zoomScale = 110
        columns = spec["columns"]
        headers = [c[0] for c in columns]
        kinds = {c[0]: c[3] for c in columns}
        width = len(columns)
        last_letter = get_column_letter(width)

        for col, (name, col_width, note, kind) in enumerate(columns, start=1):
            cell = sheet.cell(row=1, column=col, value=name)
            cell.font, cell.fill = header_font, fill(spec["color"])
            centered = kind in ("center", "date")
            cell.alignment = Alignment(wrap_text=True, vertical="center",
                                       horizontal="center" if centered else "left", indent=0 if centered else 1)
            comment = Comment(note, "Canvas Course Manager")
            comment.width, comment.height = 260, 90
            cell.comment = comment
            sheet.column_dimensions[get_column_letter(col)].width = col_width
        sheet.row_dimensions[1].height = 38

        data_rows: list[int] = []
        r = 2
        stripe = 0
        for entry in tabs.get(tab, []):
            if "_band" in entry:
                sheet.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
                cell = sheet.cell(row=r, column=1, value="▸  " + entry["_band"])
                cell.font, cell.fill = band_font, fill("DCE6F2")
                cell.alignment = Alignment(vertical="center", indent=1)
                for col in range(1, width + 1):
                    sheet.cell(row=r, column=col).border = band_border
                sheet.row_dimensions[r].height = 26
                stripe = 0
                r += 1
                continue
            background = zebra[stripe % 2]
            stripe += 1
            for col, name in enumerate(headers, start=1):
                value = entry.get(name)
                cell = sheet.cell(row=r, column=col, value=None if value in ("", None) else value)
                kind = kinds[name]
                cell.font = ref_font if kind == "ref" else edit_font
                cell.border = row_border
                cell.fill = fill("E7E6E6") if name in entry.get("_lock", ()) else fill(background)
                cell.alignment = Alignment(vertical="center", indent=0 if kind in ("center", "date") else 1,
                                           horizontal="center" if kind in ("center", "date") else "left")
                if kind == "date":
                    cell.number_format = SHEET_DATE_FORMAT
            if col := headers.index("Apply changes") + 1:
                sheet.cell(row=r, column=col).font = Font(name=SHEET_FONT, size=11, bold=True, color="595959")
            sheet.row_dimensions[r].height = 22
            sheet.row_dimensions[r].outline_level = 1 if any("_band" in e for e in tabs.get(tab, [])) else 0
            data_rows.append(r)
            r += 1

        # Rows below the table: same layout, ready for new entries.
        end = r + SHEET_EXTRA_ROWS
        for row in range(r, end + 1):
            for col, name in enumerate(headers, start=1):
                cell = sheet.cell(row=row, column=col)
                cell.font = edit_font
                if kinds[name] in ("center", "date"):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                if kinds[name] == "date":
                    cell.number_format = SHEET_DATE_FORMAT

        # Dropdowns on table rows only (not on the module bands).
        blocks: list[tuple[int, int]] = []
        for row in data_rows + list(range(r, end + 1)):
            if blocks and row == blocks[-1][1] + 1:
                blocks[-1] = (blocks[-1][0], row)
            else:
                blocks.append((row, row))
        for name, formula in dropdowns[tab].items():
            letter = get_column_letter(headers.index(name) + 1)
            check = DataValidation(type="list", formula1=formula, allow_blank=True)
            for first, last in blocks:
                check.add(f"{letter}{first}:{letter}{last}")
            sheet.add_data_validation(check)

        # Colors that follow what you type.
        sheet.conditional_formatting.add(
            f"A2:{last_letter}{end}",
            FormulaRule(formula=['$A2="YES"'], fill=fill("E2EFDA"), font=Font(color="1E4620")))
        if "Published" in headers:
            letter = get_column_letter(headers.index("Published") + 1)
            sheet.conditional_formatting.add(
                f"{letter}2:{letter}{end}",
                CellIsRule(operator="equal", formula=['"NO"'], fill=fill("FCE4E4"), font=Font(color="9C0006")))
        sheet.freeze_panes = spec["freeze"]

    book.move_sheet("Lists", offset=len(SHEET_TABS))   # Keep the hidden list last.
    book.active = 0
    book.save(path)


def new_sheet_path(folder: Path, course_name: str, blank: bool) -> Path:
    """A new, unused file name in the course folder. Existing files are never replaced."""
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", course_name).strip(" .")[:60] or "Course"
    stem = "Canvas settings - blank template" if blank else f"Canvas settings - {safe} - {time.strftime('%Y-%m-%d')}"
    path = folder / f"{stem}.xlsx"
    number = 2
    while path.exists():
        path = folder / f"{stem} ({number}).xlsx"
        number += 1
    return path


def create_settings_spreadsheet(folder: Path, course_id: str, blank: bool) -> Path:
    """Make a new settings workbook in the course folder and return its path."""
    global READ_ONLY
    if openpyxl is None:
        raise CanvasError("Making a spreadsheet needs the openpyxl helper package. Reopen the program and choose Install.")
    today = datetime.now().strftime("%B %d, %Y").replace(" 0", " ")
    if blank:
        tabs = _example_rows()
        course_name, groups, modules = "", [], []
        made = f"Blank template made {today}. Every row starts as NO."
    else:
        READ_ONLY = True
        try:
            data = read_course_for_sheet(course_id)
        finally:
            READ_ONLY = False
        tabs = build_sheet_rows(data)
        course_name = _cell_text(data["course"].get("name")) or f"Course {course_id}"
        groups = [_cell_text(g.get("name")) for g in data["groups"] if _cell_text(g.get("name"))]
        modules = [_cell_text(m.get("name")) for m in data["modules"] if _cell_text(m.get("name"))]
        made = f"From Canvas on {today}. Every row starts as NO."
    path = new_sheet_path(folder, course_name or "Course", blank)
    write_settings_workbook(path, course_name=course_name, course_id=str(course_id or ""), made_line=made,
                            tabs=tabs, groups=groups, modules=modules)
    counts = ", ".join(f"{tab}: {sum(1 for e in rows if '_band' not in e)}" for tab, rows in tabs.items())
    log(f"Saved a new settings spreadsheet: {path.name} ({counts}). Nothing in Canvas was changed.")
    return path


# ----------------------- running a job -----------------------

UPLOAD_TYPES = ["plans", "headers", "html", "slides", "checkpoints", "activities", "exams", "keys"]
TASK_PRESETS: dict[str, tuple[set[str], bool]] = {
    # task key: (upload types, also run spreadsheet settings)
    "lesson": ({"plans", "headers", "html", "slides", "checkpoints", "keys"}, False),
    "checkpoints": ({"checkpoints"}, False),
    "activities": ({"activities", "keys"}, False),
    "exams": ({"exams", "keys"}, False),
    "everything": (set(UPLOAD_TYPES), False),
    "spreadsheet": (set(), True),
    "both": (set(UPLOAD_TYPES), True),
}


def spreadsheet_course_ids() -> tuple[set[str], int]:
    """Course IDs in spreadsheet rows marked YES (they override the course ID box)."""
    if pd is None:
        raise CanvasError("The spreadsheet helper (pandas) isn't installed yet.")
    sheets = pd.read_excel(SPREADSHEET_PATH, sheet_name=None, dtype=object)
    ids: set[str] = set()
    enabled_rows = 0
    workbook_id = workbook_course_id(SPREADSHEET_PATH)
    if workbook_id:
        ids.add(workbook_id)
    for frame in sheets.values():
        frame = frame.rename(columns={c: SHEET_COLUMN_ALIASES.get(" ".join(str(c).lower().split()), c)
                                      for c in frame.columns})
        if "Apply changes" not in frame.columns:
            continue
        for _, row in frame.iterrows():
            if not _yes(row.get("Apply changes")):
                continue
            enabled_rows += 1
            row_id = _cell_text(row.get("Course ID"))
            if row_id:
                ids.add(row_id)
    return ids, enabled_rows


def configure_run(options: dict[str, Any], live: bool) -> None:
    """Load the Run tab choices into the uploader."""
    global SCOPE, UNIT_NAME, LESSON_NAME, RUN_LESSON_UPLOAD, RUN_SPREADSHEET_SETTINGS
    global UPLOAD_LESSON_PLANS, UPLOAD_LESSON_HEADERS, UPLOAD_HTML_LESSONS, UPLOAD_CHECKPOINTS
    global UPLOAD_ACTIVITIES, UPLOAD_EXAMS, UPLOAD_ANSWER_KEYS, DRY_RUN, SETTINGS_DRY_RUN, UPLOAD_SLIDES
    types = effective_types(options)
    RUN_LESSON_UPLOAD = bool(types)
    RUN_SPREADSHEET_SETTINGS = bool(options["spreadsheet"])
    SCOPE = options["scope"]
    UNIT_NAME = options["unit"]
    LESSON_NAME = options["lesson"]
    UPLOAD_LESSON_PLANS = "plans" in types
    UPLOAD_LESSON_HEADERS = "headers" in types
    UPLOAD_HTML_LESSONS = "html" in types
    UPLOAD_SLIDES = "slides" in types
    UPLOAD_CHECKPOINTS = "checkpoints" in types
    UPLOAD_ACTIVITIES = "activities" in types
    UPLOAD_EXAMS = "exams" in types
    UPLOAD_ANSWER_KEYS = "keys" in types or UPLOAD_ACTIVITIES or UPLOAD_EXAMS
    DRY_RUN = not live
    SETTINGS_DRY_RUN = not live


def configure_sort(live: bool) -> None:
    """Sorting existing quizzes uses the same preview / live switch as a normal run."""
    global DRY_RUN, SETTINGS_DRY_RUN
    DRY_RUN = not live
    SETTINGS_DRY_RUN = not live


def quiz_kinds_in(options: dict[str, Any]) -> list[str]:
    """Which kinds of quiz a Run tab choice would upload."""
    types = effective_types(options)
    return [kind for kind, key in (("checkpoint", "checkpoints"), ("activity", "activities"), ("exam", "exams"))
            if key in types]


def effective_types(options: dict[str, Any]) -> set[str]:
    """Whole-unit and whole-course runs always include the unit activity, exam, and keys."""
    types = set(options["types"])
    if types and options["scope"] in {"unit", "course"}:
        types |= {"activities", "exams", "keys"}
    if types & {"activities", "exams"}:
        types.add("keys")
    return types


def run_job() -> None:
    """Run the upload and/or spreadsheet steps chosen on the Run tab."""
    global STATE
    if not API_ROOT:
        raise CanvasError("Not connected to Canvas yet. Fill in the Course setup tab first.")
    _UNIT_MODULE_CACHE.clear()
    _PLAN_MODULE_CACHE.clear()
    GROUP_TARGETS.clear()
    global _DRIVE_INDEX
    _DRIVE_INDEX = None
    _DRIVE_FOLDERS.clear()
    STATE = read_state()
    for key in COUNTS:
        COUNTS[key] = 0
    mode = "PREVIEW (nothing in Canvas will change)" if DRY_RUN else "LIVE (Canvas will be changed)"
    log(f"Starting | course {COURSE_ID} | {mode}")
    if RUN_LESSON_UPLOAD:
        validate_canvas_targets()
        folders = choose_folders()
        if not folders:
            raise CanvasError("No lesson, activity, or exam folders were found for this choice. Check the folder names.")
        log(f"Folders to process: {len(folders)}")
        for folder in folders:
            process_lesson(folder)
        log("Upload step finished.")
    if RUN_SPREADSHEET_SETTINGS:
        log("=" * 60)
        run_spreadsheet_settings()
    log("Finished.")


# =====================================================================
#                        THE WINDOW (tkinter)
# =====================================================================
# Colors: deep slate header, quiet surfaces, one green/red mode badge.
# The header and badge look the same in light and dark mode; everything else
# comes from the palette for the current appearance.
HEADER_BG = "#22394b"
HEADER_SUB = "#b9c9d6"
PREVIEW_GREEN = "#2e7d5b"     # Badge colors (white text on top).
LIVE_RED = "#b3362b"
PALETTES: dict[str, dict[str, str]] = {
    "light": {
        "ink": "#1c1c1c", "muted": "#5d6b75", "surface": "#f4f6f8", "field": "#ffffff",
        "border": "#dcdfe3", "module_bar": "#e6ebf0", "kind": "#8a969f", "time": "#9aa5ad",
        "would": "#1f5fa8", "done": "#2e7d5b", "warn": "#a35b00", "error": "#b3362b",
    },
    "dark": {
        "ink": "#f2f2f2", "muted": "#a9b4bd", "surface": "#1c1c1c", "field": "#262626",
        "border": "#3a3a3a", "module_bar": "#33404a", "kind": "#8f9aa3", "time": "#7d8790",
        "would": "#78acff", "done": "#5fc49a", "warn": "#f0a64a", "error": "#ff7b6e",
    },
}

TASK_ORDER = ["lesson", "checkpoints", "activities", "exams", "everything", "spreadsheet", "both", "custom"]


def plural(text: str) -> str:
    low = text.lower()
    if low.endswith("quiz"):
        return text + "zes"
    if low.endswith(("s", "x", "z", "ch", "sh")):
        return text + "es"
    if low.endswith("y") and len(low) > 1 and low[-2] not in "aeiou":
        return text[:-1] + "ies"
    return text + "s"


def open_with_system(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def friendly_error(exc: Exception) -> str:
    text = str(exc)
    if "(401)" in text:
        return "Canvas didn't accept the access token. Make a new token in Canvas and paste it again."
    if "(403)" in text:
        return "Your Canvas account isn't allowed to do this in that course. Check that you are a teacher in it."
    if "(404)" in text:
        return "Canvas couldn't find that. Check the course ID."
    if "Network error" in text:
        return "Couldn't reach that address. Check the Canvas address and your internet connection."
    return text


# ========================= PROGRAM UPDATES =========================
# Updates come from published GitHub *releases* only, so work-in-progress pushed to the
# repository never reaches colleagues. Only this one program file is ever replaced.
class UpdateError(RuntimeError):
    pass


def version_tuple(text: str) -> tuple[int, ...] | None:
    """'v1.2.0' or '1.2' -> (1, 2, 0). Anything else -> None."""
    match = re.fullmatch(r"v?(\d+)(?:\.(\d+))?(?:\.(\d+))?", str(text).strip())
    if not match:
        return None
    return tuple(int(part or 0) for part in match.groups())


def is_newer(candidate: str, current: str = __version__) -> bool:
    new, old = version_tuple(candidate), version_tuple(current)
    return bool(new and old and new > old)


def is_developer_copy() -> bool:
    """A git checkout is the author's working copy: never auto-update it."""
    return (APP_DIR / ".git").exists()


def read_update_status() -> dict[str, Any]:
    try:
        data = json.loads(UPDATE_STATUS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def write_update_status(**changes: Any) -> None:
    data = read_update_status()
    data.update(changes)
    try:
        UPDATE_STATUS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass  # Not being able to remember a choice is never worth an error message.


def _github_get(url: str, timeout: int = 10) -> bytes:
    request = urllib.request.Request(url, headers={
        "User-Agent": f"CanvasCourseManager/{__version__}",
        "Accept": "application/vnd.github+json",
    })
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise UpdateError("No published version was found online yet.") from exc
        if exc.code in (403, 429):
            raise UpdateError("The update site is busy right now. Try again in an hour.") from exc
        raise UpdateError(f"The update site answered with an error ({exc.code}).") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise UpdateError("Couldn't reach the update site. Check your internet connection.") from exc


def latest_release() -> dict[str, str]:
    """The newest published (non-draft, non-pre-release) version on GitHub."""
    if not GITHUB_REPO:
        raise UpdateError("Updates aren't turned on for this copy of the program.")
    raw = _github_get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest")
    try:
        data = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise UpdateError("The update site sent something unexpected.") from exc
    tag = str(data.get("tag_name") or "")
    if not version_tuple(tag):
        raise UpdateError(f"The newest release has an unusual version name ({tag or 'blank'}).")
    return {
        "tag": tag,
        "version": tag.lstrip("vV"),
        "notes": str(data.get("body") or "").strip(),
        "page": str(data.get("html_url") or ""),
    }


def download_program(release: dict[str, str]) -> bytes:
    """Download the program file for a release and prove it is a complete, matching program."""
    tag = urllib.parse.quote(release["tag"], safe="")
    url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{tag}/{PROGRAM_FILE_NAME}"
    data = _github_get(url, timeout=30)
    try:
        source = data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise UpdateError("The downloaded file isn't a readable program file.") from exc
    if len(data) < 20_000 or "def launch(" not in source or "class App(" not in source:
        raise UpdateError("The downloaded file looks incomplete, so it was not installed.")
    try:
        compile(source, PROGRAM_FILE_NAME, "exec")   # Checks it is valid Python without running it.
    except SyntaxError as exc:
        raise UpdateError(f"The downloaded file has a problem on line {exc.lineno}, so it was not installed.") from exc
    found = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', source, re.MULTILINE)
    if not found or version_tuple(found.group(1)) != version_tuple(release["version"]):
        raise UpdateError("The downloaded file's version number doesn't match the release, so it was not installed.")
    return data


def install_program(data: bytes) -> Path:
    """Back up the current program, then swap in the new one in a single step."""
    try:
        BACKUP_DIR.mkdir(exist_ok=True)
        stamp = time.strftime("%Y%m%d-%H%M%S")
        backup = BACKUP_DIR / f"{PROGRAM_FILE.stem}_v{__version__}_{stamp}.py"
        backup.write_bytes(PROGRAM_FILE.read_bytes())
        temp = PROGRAM_FILE.with_name(f".{PROGRAM_FILE.stem}.update.tmp")
        temp.write_bytes(data)
        os.replace(temp, PROGRAM_FILE)   # All-or-nothing: never leaves a half-written program.
    except OSError as exc:
        raise UpdateError(f"Couldn't replace the program file ({exc.strerror or exc}). "
                          "Make sure the program folder isn't read-only, then try again.") from exc
    old = sorted(BACKUP_DIR.glob(f"{PROGRAM_FILE.stem}_v*.py"), key=lambda p: p.stat().st_mtime)
    for extra in old[:-BACKUPS_TO_KEEP]:
        try:
            extra.unlink()
        except OSError:
            pass
    return backup


def restart_program() -> None:
    """Start a fresh copy the same way this one was started (launcher or program file)."""
    entry = Path(sys.argv[0]).resolve() if sys.argv and sys.argv[0] else PROGRAM_FILE
    if not entry.exists():
        entry = PROGRAM_FILE
    subprocess.Popen([sys.executable, str(entry)], cwd=str(APP_DIR))


class App(tk.Tk if tk else object):  # type: ignore[misc]
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE}  (version {__version__})")
        width = min(1240, self.winfo_screenwidth() - 40)
        height = min(840, self.winfo_screenheight() - 90)
        self.geometry(f"{width}x{height}")
        self.minsize(min(940, width), min(620, height))
        self.program = load_program_settings()
        folder = starting_course_folder(self.program)
        self.settings, self.course_saved = load_course_settings(folder)
        self.current_folder = folder
        self.creds, self.creds_path = read_credentials(self.settings)
        self._move_old_credentials()
        self.events: queue.Queue = queue.Queue()
        self.running = False
        self.job = "run"                     # "run" or "sort": what the progress panel is showing.
        self.checking_update = False
        self.course_name = ""
        self.canvas_groups: list[dict[str, Any]] = []   # Gradebook groups read from Canvas.
        self.groups_weighted = False
        self.groups_course = ""              # Course ID the groups above belong to.
        self.loading_groups = False
        self.group_by_label: dict[str, dict[str, Any] | None] = {}
        self._preview_job: str | None = None
        self.task_radios: dict[str, ttk.Radiobutton] = {}
        self.type_checks: dict[str, ttk.Checkbutton] = {}
        self.scope_radios: dict[str, ttk.Radiobutton] = {}
        self.item_labels: dict[str, ttk.Label] = {}
        self.text_widgets: list[tk.Text] = []

        modern = sv_ttk is not None
        self.appearance = self.program.get("appearance", "light") if modern else "light"
        if self.appearance not in PALETTES:
            self.appearance = "light"
        self.colors = PALETTES[self.appearance]
        self._fonts()
        self._apply_theme()
        self._make_vars()
        self._build_header()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=(10, 0))
        self.page_canvases: list[tk.Canvas] = []
        self.page_run, self.tab_run = self._scroll_page()
        self.page_setup, self.tab_setup = self._scroll_page()
        self.page_look, self.tab_look = self._scroll_page()
        self.notebook.add(self.page_run, text="  Run  ")
        self.notebook.add(self.page_setup, text="  Course setup  ")
        self.notebook.add(self.page_look, text="  Names & look  ")
        self._build_run_tab()
        self._build_setup_tab()
        self._build_look_tab()
        self._build_statusbar()

        self._refresh_words()
        self._refresh_units()
        self._on_task_change()
        self._on_mode_change()
        self._update_course_label()
        self._render_sample()

        global LOG_HOOK
        LOG_HOOK = lambda message: self.events.put(("log", message))  # noqa: E731
        self.after(80, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(250, self._startup)

    # ------------------------------------------------------------ setup
    def _fonts(self) -> None:
        """Pick fonts once. With the modern look, text sizes follow the screen's scaling."""
        style = ttk.Style(self)
        mono = "Consolas" if sys.platform == "win32" else "Menlo" if sys.platform == "darwin" else "DejaVu Sans Mono"
        if sv_ttk is not None:
            sv_ttk.set_theme(self.appearance, self)   # Loads the theme and creates its fonts.
            family = self._fit_sun_valley_fonts()
            size = 10
        else:
            if sys.platform not in ("win32", "darwin") and "clam" in style.theme_names():
                style.theme_use("clam")
            base = tkfont.nametofont("TkDefaultFont")
            family = base.actual()["family"]
            size = abs(int(base.actual()["size"])) or 10
        self.family = family
        self.font_title = (family, size + 6, "bold")
        self.font_head = (family, size + 1, "bold")
        self.font_body = (family, size + 1)
        self.font_small = (family, max(size - 1, 8))
        self.font_run = (family, size + 2, "bold")
        self.font_mono = (mono, size)

    def _fit_sun_valley_fonts(self) -> str:
        """Sun Valley asks for Windows 11's 'Segoe UI Variable' in fixed pixel sizes.

        Older Windows and every Mac don't have that font, and pixel sizes stay tiny on
        high-resolution screens. Swap in the computer's own font when needed and use
        point sizes, so text is crisp and readable everywhere.
        """
        default_family = tkfont.nametofont("TkDefaultFont").actual()["family"]
        available = set(tkfont.families(self))
        have_segoe_variable = "Segoe UI Variable Text" in available or "Segoe UI Variable Display" in available
        for name in tkfont.names(self):
            if not name.startswith("SunValley"):
                continue
            font = tkfont.nametofont(name, root=self)
            pixels = abs(int(font.cget("size")))
            font.configure(size=max(8, round(pixels * 0.75)))
            if not have_segoe_variable:
                bold = "Semibold" in font.cget("family") or "Strong" in name
                font.configure(family=default_family, weight="bold" if bold else "normal")
        return tkfont.nametofont("SunValleyBodyFont", root=self).actual()["family"]

    def _apply_theme(self) -> None:
        """Paint the window for the current appearance. Safe to call again to switch."""
        style = ttk.Style(self)
        if sv_ttk is not None:
            sv_ttk.set_theme(self.appearance, self)
            # Sun Valley normally colors label backgrounds a moment later, when the theme-change
            # event arrives. On Windows that event can be used up before the window exists,
            # leaving grey boxes behind every label. Apply the colors right now instead.
            try:
                self.tk.call("configure_colors")
            except tk.TclError:
                pass
        c = self.colors = PALETTES[self.appearance]
        # Theme settings belong to one theme, so these are re-applied after every switch.
        style.configure("TLabelframe.Label", font=self.font_head, foreground=c["ink"])
        style.configure("Hint.TLabel", foreground=c["muted"], font=self.font_small)
        style.configure("Muted.TLabel", foreground=c["muted"])
        style.configure("Good.TLabel", foreground=c["done"])
        style.configure("Bad.TLabel", foreground=c["error"])
        style.configure("Warn.TLabel", foreground=c["warn"])
        style.configure("Run.TButton", font=self.font_run, padding=(22, 9))
        style.configure("Run.Accent.TButton", font=self.font_run, padding=(22, 9))
        surface = style.lookup(".", "background") or c["surface"]
        self.configure(bg=surface)
        for canvas in getattr(self, "page_canvases", []):
            canvas.configure(background=surface)
        for widget in self.text_widgets:
            self._paint_text(widget)
        # The title bar can only be colored once the window is on screen.
        self.after(150, lambda: self._title_bar(dark=self.appearance == "dark"))

    def _scroll_page(self) -> tuple[ttk.Frame, ttk.Frame]:
        """A tab page that fills the window, and scrolls up and down only when the
        window is too short to show everything (small laptop screens)."""
        outer = ttk.Frame(self.notebook)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0,
                           background=ttk.Style(self).lookup(".", "background") or self.colors["surface"])
        bar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=bar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        inner = ttk.Frame(canvas, padding=12)
        window = canvas.create_window(0, 0, window=inner, anchor="nw")
        self.page_canvases.append(canvas)

        def fit(_event: Any = None) -> None:
            needed = inner.winfo_reqheight()
            shown = canvas.winfo_height()
            height = max(shown, needed)
            canvas.itemconfigure(window, width=canvas.winfo_width(), height=height)
            canvas.configure(scrollregion=(0, 0, canvas.winfo_width(), height))
            if needed > shown + 1:
                bar.grid(row=0, column=1, sticky="ns")
            else:
                bar.grid_remove()
                canvas.yview_moveto(0)

        def wheel(event: Any) -> None:
            if not bar.winfo_ismapped():
                return
            try:   # Lists, number boxes, and text areas keep their own scrolling.
                if event.widget.winfo_class() in ("Text", "TCombobox", "TSpinbox", "Listbox"):
                    return
            except (AttributeError, tk.TclError):
                return
            if getattr(event, "num", 0) in (4, 5):
                step = -1 if event.num == 4 else 1
            else:
                step = -1 if event.delta > 0 else 1
            canvas.yview_scroll(step * 3, "units")

        def wheel_on(_event: Any) -> None:
            for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.bind_all(sequence, wheel)

        def wheel_off(_event: Any) -> None:
            for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.unbind_all(sequence)

        canvas.bind("<Configure>", fit)
        inner.bind("<Configure>", fit)
        outer.bind("<Enter>", wheel_on)
        outer.bind("<Leave>", wheel_off)
        return outer, inner

    @property
    def accent(self) -> str:
        """Style for the one main button on a screen (blue with the modern look)."""
        return "Accent.TButton" if sv_ttk is not None else "TButton"

    def _text_box(self, parent: tk.Misc, **options: Any) -> tuple[ttk.Frame, tk.Text]:
        """A text area with a themed scrollbar that follows light/dark mode."""
        frame = ttk.Frame(parent)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        text = tk.Text(frame, relief="flat", borderwidth=0, highlightthickness=1, **options)
        bar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=bar.set)
        text.grid(row=0, column=0, sticky="nsew")
        bar.grid(row=0, column=1, sticky="ns")
        self.text_widgets.append(text)
        self._paint_text(text)
        return frame, text

    def _paint_text(self, text: tk.Text) -> None:
        c = self.colors
        try:
            text.configure(background=c["field"], foreground=c["ink"], insertbackground=c["ink"],
                           highlightbackground=c["border"], highlightcolor=c["border"],
                           selectbackground="#2f60d8", selectforeground="#ffffff",
                           inactiveselectbackground="#2f60d8")
            for tag, color in (("would", "would"), ("done", "done"), ("warn", "warn"), ("error", "error"),
                               ("section", "muted"), ("time", "time"), ("kind", "kind"), ("hidden", "error")):
                if tag in text.tag_names():
                    text.tag_configure(tag, foreground=c[color])
            if "module" in text.tag_names():
                text.tag_configure("module", background=c["module_bar"])
        except tk.TclError:   # The widget was closed (for example, an update dialog).
            pass

    def _title_bar(self, dark: bool) -> None:
        """Windows 10/11: match the title bar to dark mode. Other systems: nothing to do."""
        if sys.platform != "win32":
            return
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            value = ctypes.c_int(1 if dark else 0)
            for attribute in (20, 19):   # 20 on current Windows, 19 on early Windows 10 builds.
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value),
                                                              ctypes.sizeof(value)) == 0:
                    break
        except Exception:
            pass

    def _toggle_dark(self) -> None:
        self.appearance = "dark" if self.v_dark.get() else "light"
        self.program["appearance"] = self.appearance
        self._apply_theme()
        self._render_sample()
        try:
            save_program_settings(self.program)
        except OSError:
            pass

    def _make_vars(self) -> None:
        s = self.settings
        self.v_url = tk.StringVar(value=self.creds.get("URL", ""))
        self.v_token = tk.StringVar(value=self.creds.get("TOKEN", ""))
        self.v_show_token = tk.BooleanVar(value=False)
        self.v_course = tk.StringVar(value=str(s["course_id"]))
        self.v_folder = tk.StringVar(value=s["course_folder"])
        self.v_sheet = tk.StringVar(value=s["spreadsheet"])
        self.v_plan_module = tk.StringVar(value=s["plan_module_name"])
        self.v_create_missing = tk.BooleanVar(value=s["create_missing_modules"])
        self.v_publish_pages = tk.BooleanVar(value=s["publish_new_pages"])
        self.v_slides_mode = tk.StringVar(value=s["slides_mode"])
        self.v_drive_folder = tk.StringVar(value=s["drive_folder"])
        self.v_words = {k: tk.StringVar(value=s["words"][k]) for k in DEFAULT_SETTINGS["words"]}
        self.v_emojis = {k: tk.StringVar(value=s["emojis"][k] or "(none)") for k in DEFAULT_SETTINGS["emojis"]}
        self.v_indents = {k: tk.StringVar(value=str(s["indents"][k])) for k in DEFAULT_SETTINGS["indents"]}
        self.v_use_emojis = tk.BooleanVar(value=s["use_emojis"])
        self.v_fix_emojis = tk.BooleanVar(value=s["fix_existing_emojis"])
        self.v_capitalize = tk.BooleanVar(value=s["capitalize_titles"])
        task = s.get("last_task") if s.get("last_task") in TASK_ORDER else "lesson"
        self.v_task = tk.StringVar(value=task)
        self.v_types = {k: tk.BooleanVar(value=False) for k in UPLOAD_TYPES}
        self.v_custom_sheet = tk.BooleanVar(value=False)
        self.v_scope = tk.StringVar(value=s.get("last_scope") if s.get("last_scope") in ("lesson", "unit", "course") else "lesson")
        self.v_unit = tk.StringVar(value=s.get("last_unit", ""))
        self.v_lesson = tk.StringVar(value=s.get("last_lesson", ""))
        self.v_mode = tk.StringVar(value="preview")   # Always opens in preview mode.
        self.v_dark = tk.BooleanVar(value=self.appearance == "dark")
        self.v_group = {kind: tk.StringVar() for kind in QUIZ_KINDS}

        look_vars: list[tk.Variable] = [*self.v_words.values(), *self.v_emojis.values(), *self.v_indents.values(),
                                        self.v_use_emojis, self.v_capitalize, self.v_plan_module]
        for var in look_vars:
            var.trace_add("write", lambda *_: self._schedule_sample())
        for var in self.v_words.values():
            var.trace_add("write", lambda *_: self._refresh_words())
        self.v_task.trace_add("write", lambda *_: self._on_task_change())
        self.v_scope.trace_add("write", lambda *_: self._on_scope_change())
        self.v_unit.trace_add("write", lambda *_: self._refresh_lessons())
        self.v_mode.trace_add("write", lambda *_: self._on_mode_change())
        self.v_course.trace_add("write", lambda *_: self._update_course_label())
        self.v_course.trace_add("write", lambda *_: self._refresh_group_widgets())

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=HEADER_BG)
        header.pack(fill="x")
        left = tk.Frame(header, bg=HEADER_BG)
        left.pack(side="left", padx=20, pady=8)
        tk.Label(left, text=APP_TITLE, font=self.font_title, bg=HEADER_BG, fg="white").pack(anchor="w")
        self.lbl_course = tk.Label(left, text="", font=self.font_body, bg=HEADER_BG, fg=HEADER_SUB)
        self.lbl_course.pack(anchor="w")
        self.badge = tk.Label(header, text="", font=self.font_head, fg="white", padx=14, pady=6)
        self.badge.pack(side="right", padx=20)

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self, padding=(16, 6))
        bar.pack(fill="x", side="bottom", before=self.notebook)   # Always visible, even on small screens.
        self.lbl_status = ttk.Label(bar, text="", style="Hint.TLabel")
        self.lbl_status.pack(side="left")
        ttk.Button(bar, text="Check for updates", command=lambda: self._check_updates(manual=True)).pack(side="right")
        ttk.Label(bar, text=f"Version {__version__}", style="Hint.TLabel").pack(side="right", padx=10)
        if sv_ttk is not None:
            ttk.Checkbutton(bar, text="Dark mode", style="Switch.TCheckbutton", variable=self.v_dark,
                            command=self._toggle_dark).pack(side="right", padx=(0, 14))
        else:
            ttk.Button(bar, text="Get the modern look", command=self._offer_modern_look).pack(side="right",
                                                                                            padx=(0, 14))

    def _status(self, text: str) -> None:
        self.lbl_status.config(text=text)

    # ------------------------------------------------------------ Run tab
    def _build_run_tab(self) -> None:
        tab = self.tab_run
        tab.columnconfigure(0, weight=3, uniform="cols")
        tab.columnconfigure(1, weight=2, uniform="cols")
        tab.rowconfigure(2, weight=1)

        what = ttk.LabelFrame(tab, text="What would you like to do?", padding=(10, 6))
        what.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        what.columnconfigure((0, 1), weight=1)
        for index, key in enumerate(TASK_ORDER):
            radio = ttk.Radiobutton(what, text=key, value=key, variable=self.v_task)
            radio.grid(row=index % 4, column=index // 4, sticky="w", pady=1)
            self.task_radios[key] = radio
        self.lbl_task_hint = ttk.Label(what, text="", style="Hint.TLabel", wraplength=560, justify="left")
        self.lbl_task_hint.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 2))
        inner = ttk.Frame(what)
        inner.grid(row=5, column=0, columnspan=2, sticky="w")
        for index, type_key in enumerate(UPLOAD_TYPES):
            check = ttk.Checkbutton(inner, text=type_key, variable=self.v_types[type_key])
            check.grid(row=index // 3, column=index % 3, sticky="w", padx=(0, 16))
            self.type_checks[type_key] = check
        self.chk_custom_sheet = ttk.Checkbutton(inner, text="Settings spreadsheet", variable=self.v_custom_sheet)
        self.chk_custom_sheet.grid(row=2, column=2, sticky="w")

        where = ttk.LabelFrame(tab, text="Where should it look?", padding=(10, 6))
        where.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        where.columnconfigure(1, weight=1)
        for row, key in enumerate(("lesson", "unit", "course")):
            radio = ttk.Radiobutton(where, text=key, value=key, variable=self.v_scope)
            radio.grid(row=row, column=0, columnspan=2, sticky="w", pady=1)
            self.scope_radios[key] = radio
        self.lbl_unit = ttk.Label(where, text="Unit")
        self.lbl_unit.grid(row=3, column=0, sticky="w", pady=(6, 2))
        self.cmb_unit = ttk.Combobox(where, textvariable=self.v_unit, state="readonly")
        self.cmb_unit.grid(row=3, column=1, sticky="ew", pady=(6, 2), padx=(8, 0))
        self.lbl_lesson = ttk.Label(where, text="Folder")
        self.lbl_lesson.grid(row=4, column=0, sticky="w", pady=2)
        self.cmb_lesson = ttk.Combobox(where, textvariable=self.v_lesson, state="readonly")
        self.cmb_lesson.grid(row=4, column=1, sticky="ew", pady=2, padx=(8, 0))
        self.lbl_scope_hint = ttk.Label(where, text="", style="Hint.TLabel", wraplength=380, justify="left")
        self.lbl_scope_hint.grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))

        action = ttk.Frame(tab, padding=(0, 10, 0, 0))
        action.grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Radiobutton(action, text="Preview only (recommended): changes nothing",
                        value="preview", variable=self.v_mode).pack(side="left")
        ttk.Radiobutton(action, text="Apply changes to Canvas (you'll type APPLY)",
                        value="live", variable=self.v_mode).pack(side="left", padx=(18, 0))
        ttk.Button(action, text="Open full log", command=self._open_log).pack(side="right")
        self.btn_run = ttk.Button(action, text="Run preview", style="Run.Accent.TButton" if sv_ttk else "Run.TButton",
                                  command=self._start_run)
        self.btn_run.pack(side="right", padx=(0, 10))

        progress = ttk.LabelFrame(tab, text="Progress", padding=8)
        progress.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        progress.columnconfigure(0, weight=1)
        progress.rowconfigure(1, weight=1)
        top = ttk.Frame(progress)
        top.grid(row=0, column=0, sticky="ew")
        self.lbl_summary = ttk.Label(top, text="Nothing has run yet.", font=self.font_body)
        self.lbl_summary.pack(side="left")
        ttk.Button(top, text="Clear", command=self._clear_log).pack(side="right")
        ttk.Button(top, text="Copy", command=self._copy_log).pack(side="right", padx=(0, 6))
        self.bar = ttk.Progressbar(top, mode="indeterminate", length=140)
        self.bar.pack(side="right", padx=12)
        log_box, self.txt_log = self._text_box(progress, height=3, wrap="word", font=self.font_mono,
                                               padx=10, pady=8, state="disabled")
        log_box.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        for tag in ("would", "done", "warn", "section", "time"):
            self.txt_log.tag_configure(tag)
        self.txt_log.tag_configure("error", font=(self.font_mono[0], self.font_mono[1], "bold"))
        self._paint_text(self.txt_log)
        legend = ttk.Label(progress, style="Hint.TLabel",
                           text="Blue: a change a live run would make.   Green: done.   Orange: worth a look.   Red: the run stopped.")
        legend.grid(row=2, column=0, sticky="w", pady=(4, 0))

    # ------------------------------------------------------------ Setup tab
    def _build_setup_tab(self) -> None:
        tab = self.tab_setup
        tab.columnconfigure((0, 1), weight=1, uniform="setup")

        conn = ttk.LabelFrame(tab, text="Connect to Canvas (one time)", padding=10)
        conn.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        conn.columnconfigure(1, weight=1)
        ttk.Label(conn, text="Canvas address").grid(row=0, column=0, sticky="w")
        ttk.Entry(conn, textvariable=self.v_url).grid(row=0, column=1, columnspan=2, sticky="ew", padx=8, pady=2)
        ttk.Label(conn, style="Hint.TLabel",
                  text="The start of your Canvas web address, like https://yourschool.instructure.com",
                  wraplength=320
                  ).grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
        ttk.Label(conn, text="Access token").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.ent_token = ttk.Entry(conn, textvariable=self.v_token, show="•")
        self.ent_token.grid(row=2, column=1, sticky="ew", padx=8, pady=(8, 2))
        ttk.Checkbutton(conn, text="Show", variable=self.v_show_token,
                        command=lambda: self.ent_token.config(show="" if self.v_show_token.get() else "•")
                        ).grid(row=2, column=2, sticky="w", pady=(8, 2))
        ttk.Label(conn, style="Hint.TLabel", justify="left",
                  text="In Canvas: Account (your picture) > Settings > + New Access Token. Copy the long code and "
                       "paste it here. It is saved privately in canvas_config.txt next to this program (never in a "
                       "course folder) and works for all your courses. Never share that file.", wraplength=320
                  ).grid(row=3, column=1, columnspan=2, sticky="w", padx=8)
        row = ttk.Frame(conn)
        row.grid(row=4, column=1, columnspan=2, sticky="w", padx=8, pady=(8, 0))
        ttk.Button(row, text="Test connection", command=self._test_connection).pack(side="left")
        self.lbl_conn = ttk.Label(row, text="", wraplength=300, justify="left")
        self.lbl_conn.pack(side="left", padx=10)

        course = ttk.LabelFrame(tab, text="Your course", padding=10)
        course.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        course.columnconfigure(1, weight=1)
        ttk.Label(course, text="Course ID").grid(row=0, column=0, sticky="w")
        ttk.Entry(course, textvariable=self.v_course, width=16).grid(row=0, column=1, sticky="w", padx=8, pady=2)
        ttk.Label(course, style="Hint.TLabel",
                  text="The number after /courses/ in your course's web address. Pasting the whole address works too.",
                  wraplength=320
                  ).grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
        ttk.Label(course, text="Course folder").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.cmb_folder = ttk.Combobox(course, textvariable=self.v_folder,
                                       values=recent_course_folders(self.program))
        self.cmb_folder.grid(row=2, column=1, sticky="ew", padx=8, pady=(8, 2))
        self.cmb_folder.bind("<<ComboboxSelected>>", lambda _e: self._commit_folder())
        self.cmb_folder.bind("<Return>", lambda _e: self._commit_folder())
        self.cmb_folder.bind("<FocusOut>", lambda _e: self._commit_folder())
        ttk.Button(course, text="Browse…", command=self._browse_folder).grid(row=2, column=2, pady=(8, 2))
        self.lbl_folder = ttk.Label(course, text="", style="Hint.TLabel", wraplength=320)
        self.lbl_folder.grid(row=3, column=1, columnspan=2, sticky="w", padx=8)
        ttk.Label(course, style="Hint.TLabel", wraplength=320, justify="left",
                  text="The folder that holds your Unit folders. Each course folder remembers its own course ID, "
                       "spreadsheet, and look. Use the arrow to switch between courses you've used before."
                  ).grid(row=6, column=1, columnspan=2, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(course, text="Settings spreadsheet").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(course, textvariable=self.v_sheet).grid(row=4, column=1, sticky="ew", padx=8, pady=(8, 2))
        ttk.Button(course, text="Browse…", command=self._browse_sheet).grid(row=4, column=2, pady=(8, 2))
        ttk.Label(course, style="Hint.TLabel",
                  text="Optional. Only needed for due dates, points, publishing, and completion requirements.",
                  wraplength=320
                  ).grid(row=5, column=1, columnspan=2, sticky="w", padx=8)
        ttk.Button(course, text="Create settings spreadsheet…", command=self._create_sheet_dialog
                   ).grid(row=7, column=1, columnspan=2, sticky="w", padx=8, pady=(12, 2))
        ttk.Label(course, style="Hint.TLabel", wraplength=320, justify="left",
                  text="Makes a new spreadsheet filled in from your Canvas course, ready to edit. "
                       "It only reads Canvas and never replaces an existing file."
                  ).grid(row=8, column=1, columnspan=2, sticky="w", padx=8)

        modules = ttk.LabelFrame(tab, text="Modules and pages", padding=10)
        modules.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(12, 0))
        ttk.Checkbutton(modules, text="Create missing modules for me", variable=self.v_create_missing).pack(anchor="w")
        ttk.Label(modules, style="Hint.TLabel",
                  text="If a unit module or the lesson plans module doesn't exist yet, it is added as unpublished "
                       "(hidden from students) at the end of the module list.", wraplength=420, justify="left").pack(anchor="w", padx=22)
        ttk.Checkbutton(modules, text="Publish new HTML lesson pages right away",
                        variable=self.v_publish_pages).pack(anchor="w", pady=(6, 0))
        ttk.Label(modules, style="Hint.TLabel",
                  text="Leave this off to keep new pages hidden until you publish them yourself.").pack(anchor="w", padx=22)

        slides = ttk.LabelFrame(tab, text="Presentations", padding=10)
        slides.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(12, 0))
        slides.columnconfigure(1, weight=1)
        ttk.Label(slides, text="PowerPoints in lesson folders:").grid(row=0, column=0, columnspan=3, sticky="w")
        for row, (value, text) in enumerate((("canvas", "Show them in Canvas (no Google needed)"),
                                             ("google", "Turn them into Google Slides, publish, and show them"),
                                             ("off", "Don't upload them")), start=1):
            ttk.Radiobutton(slides, text=text, value=value, variable=self.v_slides_mode
                            ).grid(row=row, column=0, columnspan=3, sticky="w", padx=(12, 0))
        ttk.Label(slides, text="Drive folder").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(slides, textvariable=self.v_drive_folder).grid(row=4, column=1, sticky="ew", padx=8, pady=(8, 2))
        ttk.Button(slides, text="Connect Google", command=self._connect_google).grid(row=4, column=2, pady=(8, 2))
        ttk.Label(slides, style="Hint.TLabel", justify="left", wraplength=440,
                  text="Optional. Paste the link to your course folder in Google Drive to show Google Slides "
                       "already there, matched by their Unit and Lesson numbers. A link file in a lesson folder "
                       "(drag a link in, or a text file named 'slides link') always works.").grid(
            row=5, column=0, columnspan=3, sticky="w")
        self.lbl_google = ttk.Label(slides, text="", wraplength=440, justify="left")
        self.lbl_google.grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))

        always = ttk.Frame(modules, padding=(0, 14, 0, 0))
        always.pack(fill="x", anchor="w")
        ttk.Label(always, wraplength=460, justify="left", style="Muted.TLabel",
                  text="Always on: answer keys are hidden from students, nothing in Canvas is ever deleted, "
                       "and if two Canvas items look alike the program skips them instead of guessing.").pack(anchor="w")
        ttk.Button(always, text="Save setup", style=self.accent, command=self._save_setup).pack(anchor="w", pady=(10, 0))
        self.v_folder.trace_add("write", lambda *_: self._refresh_units())

        grades = ttk.LabelFrame(tab, text="Gradebook groups for quizzes", padding=10)
        grades.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        grades.columnconfigure(2, weight=1)
        ttk.Label(grades, style="Hint.TLabel", wraplength=860, justify="left",
                  text=f"Canvas puts uploaded quizzes in a group called \"{IMPORTED_GROUP_NAME}\". If your course "
                       "uses weighted grading, that group usually counts 0%. Choose where each kind of quiz belongs "
                       "and the program puts new quizzes there as it uploads them. Your choices save right away."
                  ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        self.group_kind_labels: dict[str, ttk.Label] = {}
        self.cmb_groups: dict[str, ttk.Combobox] = {}
        for r, kind in enumerate(QUIZ_KINDS, start=1):
            label = ttk.Label(grades, text="")
            label.grid(row=r, column=0, sticky="w", pady=2)
            self.group_kind_labels[kind] = label
            combo = ttk.Combobox(grades, textvariable=self.v_group[kind], state="readonly", width=42)
            combo.grid(row=r, column=1, sticky="w", padx=8, pady=2)
            combo.bind("<<ComboboxSelected>>", lambda _e, k=kind: self._group_picked(k))
            self.cmb_groups[kind] = combo
        buttons = ttk.Frame(grades)
        buttons.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Button(buttons, text="Load groups from Canvas",
                   command=self._load_groups).pack(side="left")
        ttk.Button(buttons, text="Sort existing quizzes…", command=self._start_sort).pack(side="left", padx=(8, 0))
        self.lbl_groups = ttk.Label(grades, text="", wraplength=860, justify="left")
        self.lbl_groups.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # Missing a group? Add one right here: name, percentage, and which quizzes it's for.
        ttk.Label(grades, text="Add a new group to Canvas", font=self.font_body
                  ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(14, 2))
        add = ttk.Frame(grades)
        add.grid(row=7, column=0, columnspan=3, sticky="w")
        self.v_new_group_name = tk.StringVar()
        self.v_new_group_weight = tk.StringVar()
        self.v_new_group_for = tk.StringVar()
        self.new_group_kinds: dict[str, str | None] = {}
        ttk.Label(add, text="Name").pack(side="left")
        ttk.Entry(add, textvariable=self.v_new_group_name, width=22).pack(side="left", padx=(6, 14))
        ttk.Label(add, text="Percentage").pack(side="left")
        ttk.Entry(add, textvariable=self.v_new_group_weight, width=6).pack(side="left", padx=(6, 2))
        ttk.Label(add, text="%").pack(side="left", padx=(0, 14))
        ttk.Label(add, text="Use it for").pack(side="left")
        self.cmb_new_group_for = ttk.Combobox(add, textvariable=self.v_new_group_for, state="readonly", width=18)
        self.cmb_new_group_for.pack(side="left", padx=(6, 14))
        self.cmb_new_group_for.bind("<<ComboboxSelected>>", lambda _e: self._suggest_group_name())
        ttk.Button(add, text="Add to Canvas…", command=self._add_group_clicked).pack(side="left")
        self.lbl_new_group = ttk.Label(grades, text="", wraplength=860, justify="left")
        self.lbl_new_group.grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(grades, style="Hint.TLabel", wraplength=860, justify="left",
                  text="Load groups only reads Canvas. Adding a group shows you a preview and asks you to type "
                       "APPLY; nothing in Canvas is ever deleted. Sort existing quizzes moves quizzes this program "
                       f"already uploaded, but only ones still in \"{IMPORTED_GROUP_NAME}\", and it runs as a "
                       "preview first. The Assignment group column in your settings spreadsheet still works for "
                       "one-off exceptions."
                  ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(8, 0))

    # ------------------------------------------------------------ Names & look tab
    def _build_look_tab(self) -> None:
        tab = self.tab_look
        tab.columnconfigure(2, weight=1)
        tab.rowconfigure(0, weight=1)

        # Column 1: words used for folders and titles.
        left = ttk.Frame(tab)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        words = ttk.LabelFrame(left, text="Your words", padding=(10, 6))
        words.pack(fill="x")
        for index, (key, label, examples) in enumerate(WORD_FIELDS):
            ttk.Label(words, text=label).grid(row=index * 2, column=0, sticky="w", pady=(2, 0))
            ttk.Entry(words, textvariable=self.v_words[key], width=18).grid(row=index * 2, column=1, sticky="w",
                                                                            padx=(8, 0), pady=(2, 0))
            ttk.Label(words, text=f"e.g. {examples}", style="Hint.TLabel").grid(row=index * 2 + 1, column=1,
                                                                               sticky="w", padx=(8, 0))
        base = len(WORD_FIELDS) * 2
        ttk.Label(words, text="Lesson plans module").grid(row=base, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(words, textvariable=self.v_plan_module, width=18).grid(row=base, column=1, sticky="w",
                                                                         padx=(8, 0), pady=(8, 0))
        ttk.Checkbutton(words, text="Capitalize Every Word in lesson names", variable=self.v_capitalize
                        ).grid(row=base + 1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.lbl_word_hint = ttk.Label(words, text="", style="Hint.TLabel", wraplength=300, justify="left")
        self.lbl_word_hint.grid(row=base + 2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        actions = ttk.Frame(left)
        actions.pack(fill="x", pady=(10, 0))
        actions.columnconfigure((0, 1), weight=1)
        ttk.Button(actions, text="Save look", style=self.accent, command=self._save_setup).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Restore defaults", command=self._restore_look).grid(row=0, column=1, sticky="ew")
        ttk.Button(actions, text="Share look…", command=self._export_look).grid(row=1, column=0, sticky="ew",
                                                                               padx=(0, 4), pady=(4, 0))
        ttk.Button(actions, text="Use a colleague's look…", command=self._import_look).grid(row=1, column=1,
                                                                                         sticky="ew", pady=(4, 0))

        # Column 2: emoji and indent for each kind of item.
        items = ttk.LabelFrame(tab, text="Emoji and indent", padding=(10, 6))
        items.grid(row=0, column=1, sticky="nsw", padx=(0, 10))
        ttk.Checkbutton(items, text="Start titles with an emoji", variable=self.v_use_emojis,
                        command=self._toggle_emoji_boxes).grid(row=0, column=0, columnspan=3, sticky="w")
        self.chk_fix_emojis = ttk.Checkbutton(items, text="Update emojis on items already in Canvas",
                                              variable=self.v_fix_emojis)
        self.chk_fix_emojis.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(items, text="Emoji", style="Hint.TLabel").grid(row=2, column=1, sticky="w", padx=8, pady=(6, 0))
        ttk.Label(items, text="Indent", style="Hint.TLabel").grid(row=2, column=2, sticky="w", pady=(6, 0))
        # (row key, emoji key or None, indent key)
        self.item_rows = [
            ("lesson_header", "lesson_header", "lesson_header"),
            ("page", "page", "page"),
            ("slides", "slides", "slides"),
            ("checkpoint", "checkpoint", "checkpoint"),
            ("checkpoint_key", None, "checkpoint_key"),
            ("activity_header", "activity_header", "assessment_header"),
            ("activity", "activity", "assessment"),
            ("exam_header", "exam_header", "assessment_header"),
            ("exam", "exam", "assessment"),
            ("assessment_key", None, "assessment_key"),
            ("plan", "plan", "plan"),
        ]
        self.emoji_boxes: list[ttk.Combobox] = []
        self.emoji_more: list[ttk.Button] = []
        for row, (row_key, emoji_key, indent_key) in enumerate(self.item_rows, start=3):
            label = ttk.Label(items, text=row_key)
            label.grid(row=row, column=0, sticky="w", pady=2)
            self.item_labels[row_key] = label
            if emoji_key:
                cell = ttk.Frame(items)
                cell.grid(row=row, column=1, sticky="w", padx=8, pady=2)
                box = ttk.Combobox(cell, textvariable=self.v_emojis[emoji_key], values=EMOJI_CHOICES,
                                   width=5 if sv_ttk else 7)
                box.pack(side="left")
                more = ttk.Button(cell, text="…", width=2, style="Toolbutton" if sv_ttk else "TButton",
                                  command=lambda k=emoji_key, r=row_key: self._pick_emoji(k, r))
                more.pack(side="left", padx=(2, 0))
                self.emoji_boxes.append(box)
                self.emoji_more.append(more)
            else:
                ttk.Label(items, text="none", style="Hint.TLabel").grid(row=row, column=1, sticky="w", padx=12)
            ttk.Spinbox(items, from_=0, to=5, width=2 if sv_ttk else 3, textvariable=self.v_indents[indent_key],
                        state="readonly").grid(row=row, column=2, sticky="w", pady=2)
        ttk.Label(items, style="Hint.TLabel", wraplength=300, justify="left",
                  text="Choose from the list, or click … to pick from many more emojis or paste one you "
                       "copied. Choose (none) for no emoji. Indent 0 is flush left; "
                       "Canvas allows up to 5. Activity and exam rows share one indent so they line up. Answer keys "
                       "never get an emoji.\n\nLeave \"Update emojis\" off to keep the emojis already in Canvas and "
                       "use yours only on new items.").grid(
            row=len(self.item_rows) + 3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self._toggle_emoji_boxes()

        # Column 3: live sample of the finished module.
        sample = ttk.LabelFrame(tab, text="How it will look in Canvas", padding=8)
        sample.grid(row=0, column=2, sticky="nsew")
        sample.rowconfigure(0, weight=1)
        sample.columnconfigure(0, weight=1)
        self.txt_sample = tk.Text(sample, wrap="word", relief="flat", borderwidth=0, highlightthickness=1,
                                  font=self.font_body, padx=0, pady=6, cursor="arrow", spacing1=4, spacing3=4,
                                  width=40)
        self.txt_sample.grid(row=0, column=0, sticky="nsew")
        self.text_widgets.append(self.txt_sample)
        for level in range(6):
            pad = 14 + level * 22
            self.txt_sample.tag_configure(f"ind{level}", lmargin1=pad, lmargin2=pad + 12)
        self.txt_sample.tag_configure("module", font=self.font_head, lmargin1=10, spacing1=10, spacing3=10)
        self.txt_sample.tag_configure("heading", font=self.font_head)
        self.txt_sample.tag_configure("kind", font=self.font_small)
        self.txt_sample.tag_configure("hidden", font=self.font_small)
        self._paint_text(self.txt_sample)
        sample_hint = ttk.Label(sample, style="Hint.TLabel", wraplength=250, justify="left",
                                text="Bold lines are headings. (hidden) means students never see it. Items already "
                                     "in Canvas are renamed to match the next time you upload them; the preview "
                                     "lists each rename first.")
        sample_hint.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        # Re-wrap the hint to whatever width the column ends up with.
        sample_hint.bind("<Configure>", lambda e: sample_hint.config(wraplength=max(120, e.width - 4)))

    def _toggle_emoji_boxes(self) -> None:
        state = "normal" if self.v_use_emojis.get() else "disabled"
        for box in self.emoji_boxes:
            box.config(state=state)
        for button in getattr(self, "emoji_more", []):
            button.config(state=state)
        self.chk_fix_emojis.config(state=state)

    def _pick_emoji(self, emoji_key: str, row_key: str) -> None:
        """A window of emojis to click, plus a box to paste any other emoji.

        Some computers' emoji keyboards can't type into this program's boxes,
        so this works the same everywhere.
        """
        win = tk.Toplevel(self)
        name = self.item_labels[row_key].cget("text") if row_key in self.item_labels else "this item"
        win.title(f"Emoji for {name}")
        win.transient(self)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Click an emoji for {name}.", font=self.font_head).pack(anchor="w")
        if sys.platform == "win32":
            face = "Segoe UI Emoji"
        elif sys.platform == "darwin":
            face = "Apple Color Emoji"
        else:
            face = self.family
        ttk.Style(self).configure("Emoji.TButton", font=(face, 16), padding=(4, 2), width=3)

        def choose(value: str) -> None:
            self.v_emojis[emoji_key].set(value)
            win.destroy()

        tabs = ttk.Notebook(frame)
        tabs.pack(fill="both", pady=(8, 10))
        for title, emojis in EMOJI_LIBRARY:
            page = ttk.Frame(tabs, padding=8)
            tabs.add(page, text=title)
            for index, value in enumerate(emojis.split()):
                ttk.Button(page, text=value, style="Emoji.TButton",
                           command=lambda v=value: choose(v)).grid(row=index // 8, column=index % 8, padx=2, pady=2)

        ttk.Label(frame, text="Or paste any emoji you copied (Ctrl+V, or Command+V on a Mac):").pack(anchor="w")
        row = ttk.Frame(frame)
        row.pack(fill="x", pady=(4, 0))
        pasted = tk.StringVar()
        entry = ttk.Entry(row, textvariable=pasted, width=10, font=(face, 14))
        entry.pack(side="left")
        message = ttk.Label(frame, text="", style="Bad.TLabel", wraplength=420, justify="left")

        def use_pasted() -> None:
            value = clean_emoji(pasted.get())
            if not value:
                message.config(text="Paste an emoji into the box first.")
            elif emoji_problem(value):
                message.config(text=f"That has {emoji_problem(value)}. Paste just an emoji.")
            else:
                choose(value)

        ttk.Button(row, text="Use it", command=use_pasted).pack(side="left", padx=(6, 0))
        entry.bind("<Return>", lambda _e: use_pasted())
        message.pack(anchor="w", pady=(4, 0))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="No emoji", command=lambda: choose("(none)")).pack(side="left")
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda _e: win.destroy())
        win.grab_set()

    def _emoji_problems(self, s: dict[str, Any]) -> list[str]:
        """Friendly messages for emoji boxes holding letters or numbers."""
        if not s["use_emojis"]:
            return []
        labels = dict(EMOJI_FIELDS)
        problems = []
        for key, value in s["emojis"].items():
            why = emoji_problem(value)
            if why:
                problems.append(f"The emoji for \"{labels.get(key, key)}\" ({value}) has {why}. Only emojis can go "
                                "at the start of titles; the program finds items it already made by their words, "
                                "so letters there would cause duplicates.")
        return problems

    # ------------------------------------------------------------ settings in/out
    def collect_settings(self) -> dict[str, Any]:
        s = copy.deepcopy(self.settings)
        s["course_id"] = course_id_from_text(self.v_course.get())
        s["course_folder"] = self.current_folder
        s["spreadsheet"] = self.v_sheet.get().strip()
        s["plan_module_name"] = self.v_plan_module.get().strip() or "Lesson Plans"
        s["create_missing_modules"] = self.v_create_missing.get()
        s["publish_new_pages"] = self.v_publish_pages.get()
        s["slides_mode"] = self.v_slides_mode.get()
        s["drive_folder"] = self.v_drive_folder.get().strip()
        s["words"] = {k: v.get().strip() for k, v in self.v_words.items()}
        s["emojis"] = {k: clean_emoji(v.get()) for k, v in self.v_emojis.items()}
        indents: dict[str, int] = {}
        for k, v in self.v_indents.items():
            try:
                indents[k] = int(v.get())
            except ValueError:
                indents[k] = DEFAULT_SETTINGS["indents"][k]
        s["indents"] = indents
        s["use_emojis"] = self.v_use_emojis.get()
        s["fix_existing_emojis"] = self.v_fix_emojis.get()
        s["capitalize_titles"] = self.v_capitalize.get()
        s["last_task"] = self.v_task.get()
        s["last_scope"] = self.v_scope.get()
        s["last_unit"] = self.v_unit.get()
        s["last_lesson"] = self.v_lesson.get()
        return _merge(DEFAULT_SETTINGS, s)

    def _set_look_vars(self, s: dict[str, Any]) -> None:
        for k, v in self.v_words.items():
            v.set(s["words"][k])
        for k, v in self.v_emojis.items():
            v.set(s["emojis"][k] or "(none)")
        for k, v in self.v_indents.items():
            v.set(str(s["indents"][k]))
        self.v_use_emojis.set(s["use_emojis"])
        self.v_capitalize.set(s["capitalize_titles"])
        self.v_plan_module.set(s["plan_module_name"])
        self._toggle_emoji_boxes()

    def _set_course_vars(self, s: dict[str, Any]) -> None:
        """Show one course's saved choices in every tab."""
        self.v_course.set(str(s["course_id"]))
        self.v_sheet.set(s["spreadsheet"])
        self.v_create_missing.set(s["create_missing_modules"])
        self.v_publish_pages.set(s["publish_new_pages"])
        self.v_slides_mode.set(s["slides_mode"])
        self.v_drive_folder.set(s["drive_folder"])
        self.v_fix_emojis.set(s["fix_existing_emojis"])
        self._set_look_vars(s)
        if s.get("last_task") in TASK_ORDER:
            self.v_task.set(s["last_task"])
        if s.get("last_scope") in ("lesson", "unit", "course"):
            self.v_scope.set(s["last_scope"])
        self.v_unit.set(s.get("last_unit", ""))
        self.v_lesson.set(s.get("last_lesson", ""))

    def _save_course(self, s: dict[str, Any] | None = None) -> bool:
        """Save the current course's settings in its folder. False when there's no folder yet."""
        s = s or self.collect_settings()
        if not s["course_folder"] or not Path(s["course_folder"]).is_dir():
            return False
        save_course_settings(s)
        self.settings, self.course_saved = s, True
        remember_course_folder(self.program, s["course_folder"])
        try:
            save_program_settings(self.program)
        except OSError:
            pass
        return True

    def _commit_folder(self) -> None:
        """Switch courses when the course folder box holds a different folder."""
        text = self.v_folder.get().strip()
        if not text or (self.current_folder and _same_folder(text, self.current_folder)):
            return
        self._switch_course(text)

    def _switch_course(self, folder_text: str) -> None:
        folder = Path(folder_text).expanduser()
        if not folder.is_dir():
            messagebox.showwarning("Folder not found", f"This folder doesn't exist on this computer:\n\n{folder}",
                                   parent=self)
            self.v_folder.set(self.current_folder)
            return
        if self.running:
            self.v_folder.set(self.current_folder)
            return
        previous = self.collect_settings()
        try:
            self._save_course(previous)          # Keep the course we're leaving exactly as it was.
        except (OSError, CanvasError):
            pass
        new_settings, found = load_course_settings(str(folder))
        if not found and previous["course_folder"]:
            if messagebox.askyesno(
                    "New course folder",
                    f"{folder.name} doesn't have its own setup yet.\n\n"
                    f"Use the same words, emojis, and indents as {Path(previous['course_folder']).name}?\n\n"
                    "Yes: copy that look.   No: start from the standard look.\n\n"
                    "(You'll still need to type this course's ID on the Course setup tab.)", parent=self):
                for key in LOOK_KEYS:
                    new_settings[key] = copy.deepcopy(previous[key])
        self.current_folder = str(folder)
        self.settings, self.course_saved = new_settings, found
        self.course_name = ""
        self.v_folder.set(str(folder))
        self._set_course_vars(new_settings)
        if found:
            self._save_course(new_settings)      # Also moves an older setup into the course folder.
        else:
            remember_course_folder(self.program, str(folder))
            try:
                save_program_settings(self.program)
            except OSError:
                pass
        self.cmb_folder.config(values=recent_course_folders(self.program))
        self._refresh_units()
        self._update_course_label()
        self._render_sample()
        self.canvas_groups, self.groups_course, self.groups_weighted = [], "", False
        self._refresh_group_widgets()
        if found:
            self._load_groups(quiet=True)
        if found:
            self._status(f"Opened {folder.name}. Its saved setup is loaded.")
        else:
            self.notebook.select(self.page_setup)
            self._status(f"Opened {folder.name}. Add its course ID, then click Save setup.")

    def _move_old_credentials(self) -> None:
        """Older setups kept canvas_config.txt in the course folder. Copy it next to the program."""
        if self.creds_path == CREDENTIALS_PATH or not self.creds.get("URL") or not self.creds.get("TOKEN"):
            return
        try:
            write_credentials(self.creds["URL"], self.creds["TOKEN"], CREDENTIALS_PATH)
            self.old_creds_path = self.creds_path
            self.creds_path = CREDENTIALS_PATH
        except OSError:
            pass

    def _save_credentials(self) -> None:
        url = clean_canvas_url(self.v_url.get())
        token = self.v_token.get().strip()
        if url and token and (url != self.creds.get("URL") or token != self.creds.get("TOKEN")):
            write_credentials(url, token, CREDENTIALS_PATH)   # Always next to the program, never in a course.
            self.creds_path = CREDENTIALS_PATH
            self.creds = {**self.creds, "URL": url, "TOKEN": token}
            self.v_url.set(url)

    def _save_setup(self) -> None:
        self._commit_folder()
        s = self.collect_settings()
        problems = self._emoji_problems(s)
        if problems:
            messagebox.showwarning("Check your emojis", "\n\n".join(problems), parent=self)
            return
        try:
            self._save_credentials()
            if self._save_course(s):
                self._status(f"Saved at {time.strftime('%H:%M')}. This course's choices are in "
                             f"{Path(s['course_folder']).name}/{COURSE_SETTINGS_NAME}.")
            else:
                self._status("Canvas login saved. Choose your course folder to save the course setup.")
                messagebox.showinfo("Choose your course folder",
                                    "Your Canvas login is saved.\n\nNext, choose your course folder (the folder "
                                    "that holds your Unit folders). The course setup is saved inside it.",
                                    parent=self)
        except (OSError, CanvasError) as exc:
            messagebox.showerror("Couldn't save", f"The settings could not be saved:\n\n{exc}", parent=self)

    def _restore_look(self) -> None:
        if messagebox.askyesno("Restore defaults", "Put every word, emoji, and indent back to the standard setup?",
                               parent=self):
            self._set_look_vars(copy.deepcopy(DEFAULT_SETTINGS))

    def _export_look(self) -> None:
        path = filedialog.asksaveasfilename(parent=self, title="Save your look for a colleague",
                                            defaultextension=".json", initialfile="canvas_look.json",
                                            filetypes=[("Look settings", "*.json")])
        if not path:
            return
        s = self.collect_settings()
        Path(path).write_text(json.dumps({k: s[k] for k in LOOK_KEYS}, indent=2, ensure_ascii=False), encoding="utf-8")
        self._status(f"Saved {Path(path).name}. Send it to a colleague; it holds no passwords or course details.")

    def _import_look(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Choose a look file from a colleague",
                                          filetypes=[("Look settings", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("not a settings file")
        except (OSError, ValueError) as exc:
            messagebox.showerror("Couldn't read that file", str(exc), parent=self)
            return
        merged = _merge(DEFAULT_SETTINGS, {**self.collect_settings(), **{k: data[k] for k in LOOK_KEYS if k in data}})
        self._set_look_vars(merged)
        self._status("Look loaded. Check the sample, then click Save look.")

    # ------------------------------------------------------------ live labels
    def _refresh_words(self) -> None:
        if not self.task_radios:
            return
        w = {k: (v.get().strip() or DEFAULT_SETTINGS["words"][k]) for k, v in self.v_words.items()}
        U, L, C, A, E, S = w["unit"], w["lesson"], w["checkpoint"], w["activity"], w["exam"], w["slides"]
        tasks = {
            "lesson": f"{L} materials", "checkpoints": f"{plural(C)} only",
            "activities": f"{U} {plural(A).lower()} only", "exams": f"{U} {plural(E).lower()} only",
            "everything": "Everything in the folders", "spreadsheet": "Settings spreadsheet only",
            "both": "Everything + spreadsheet", "custom": "Let me choose",
        }
        for key, radio in self.task_radios.items():
            radio.config(text=tasks[key])
        self.task_hints = {
            "lesson": f"{L} headings, lesson plans, HTML pages, presentations, {plural(C).lower()}, and their answer keys.",
            "checkpoints": f"Only the {plural(C).lower()} (quiz ZIP files) in each {L.lower()} folder.",
            "activities": f"The {U.lower()} {A.lower()} quiz, its heading, and its answer key.",
            "exams": f"The {U.lower()} {E.lower()} quiz, its heading, and its answer key.",
            "everything": "Every kind of file below.",
            "spreadsheet": "Dates, points, publishing, and completion rules from your settings spreadsheet.",
            "both": "Every kind of file below, then the settings spreadsheet.",
            "custom": "Tick exactly what you want below.",
        }
        self.lbl_task_hint.config(text=self.task_hints.get(self.v_task.get(), ""))
        types = {
            "plans": "Lesson plans", "headers": f"{L} headings", "html": "HTML lesson pages", "slides": "Presentations",
            "checkpoints": plural(C), "activities": plural(A), "exams": plural(E),
            "keys": "Answer keys (hidden)",
        }
        for key, check in self.type_checks.items():
            check.config(text=types[key])
        scopes = {
            "lesson": f"One folder (a {L.lower()}, {A.lower()}, or {E.lower()})",
            "unit": f"One whole {U.lower()}",
            "course": "The entire course",
        }
        for key, radio in self.scope_radios.items():
            radio.config(text=scopes[key])
        self.lbl_unit.config(text=U)
        rows = {
            "lesson_header": f"{L} heading", "page": "HTML lesson page", "slides": f"Presentation ({S})", "checkpoint": C,
            "checkpoint_key": f"{C} answer key", "activity_header": f"{A} heading", "activity": A,
            "exam_header": f"{E} heading", "exam": E, "assessment_key": f"{A} / {E} answer keys",
            "plan": "Lesson plan file",
        }
        for key, label in self.item_labels.items():
            label.config(text=rows[key])
        self.lbl_word_hint.config(
            text=f"Folders look like '{U} 6' and '{L} 32 Topic'; quiz ZIPs have '{C}', '{A}', or '{E}' in the "
                 f"name. Unit modules start with '{U} 6'.")
        self._on_scope_change()
        self._refresh_folder_status()
        self._refresh_group_widgets()

    def _update_course_label(self) -> None:
        cid = course_id_from_text(self.v_course.get())
        if cid.isdigit():
            text = f"Course {cid}" + (f": {self.course_name}" if self.course_name else "")
        else:
            text = "No course chosen yet. Start on the Course setup tab."
        self.lbl_course.config(text=text)

    def _on_mode_change(self) -> None:
        if self.running:
            return
        live = self.v_mode.get() == "live"
        self.badge.config(text="Live mode" if live else "Preview mode", bg=LIVE_RED if live else PREVIEW_GREEN)
        self.btn_run.config(text="Apply changes…" if live else "Run preview")

    def _on_task_change(self) -> None:
        if not self.type_checks:
            return
        task = self.v_task.get()
        custom = task == "custom"
        if not custom:
            types, _ = TASK_PRESETS[task]
            for key, var in self.v_types.items():
                var.set(key in types)
        for check in self.type_checks.values():
            check.config(state="normal" if custom else "disabled")
        self.chk_custom_sheet.config(state="normal" if custom else "disabled")
        if not custom:
            self.v_custom_sheet.set(TASK_PRESETS[task][1])
        if hasattr(self, "task_hints"):
            self.lbl_task_hint.config(text=self.task_hints.get(task, ""))
        self._on_scope_change()

    def _on_scope_change(self) -> None:
        if not self.scope_radios:
            return
        upload = self.v_task.get() != "spreadsheet"
        scope = self.v_scope.get()
        for radio in self.scope_radios.values():
            radio.config(state="normal" if upload else "disabled")
        self.cmb_unit.config(state="readonly" if upload and scope in ("lesson", "unit") else "disabled")
        self.cmb_lesson.config(state="readonly" if upload and scope == "lesson" else "disabled")
        if not upload:
            hint = "The spreadsheet step doesn't use your folders."
        elif scope == "lesson":
            hint = "Tip: try one folder first and check it in Canvas before doing a whole unit."
        else:
            a = plural(self.v_words["activity"].get() or "Activity").lower()
            e = plural(self.v_words["exam"].get() or "Exam").lower()
            hint = f"Whole-unit and whole-course runs always include the unit {a}, {e}, and their answer keys."
        self.lbl_scope_hint.config(text=hint)

    def _refresh_units(self) -> None:
        if not hasattr(self, "cmb_unit"):
            return
        apply_look(self.collect_settings())
        folder = Path(self.v_folder.get().strip()) if self.v_folder.get().strip() else None
        units = list_units(folder) if folder and folder.is_dir() else []
        self.cmb_unit.config(values=units)
        if units and self.v_unit.get() not in units:
            self.v_unit.set(units[0])
        elif not units:
            self.v_unit.set("")
        self._refresh_lessons()
        self._refresh_folder_status()

    def _refresh_lessons(self) -> None:
        if not hasattr(self, "cmb_lesson"):
            return
        folder = self.v_folder.get().strip()
        names = list_unit_folders(Path(folder), self.v_unit.get()) if folder and self.v_unit.get() else []
        self.cmb_lesson.config(values=names)
        if names and self.v_lesson.get() not in names:
            lessons = [n for n in names if is_lesson_name(n)]
            self.v_lesson.set((lessons or names)[0])
        elif not names:
            self.v_lesson.set("")

    def _refresh_folder_status(self) -> None:
        if not hasattr(self, "lbl_folder"):
            return
        text = self.v_folder.get().strip()
        unit_word = self.v_words["unit"].get().strip() or "Unit"
        if not text:
            self.lbl_folder.config(text="Choose the folder that holds your unit folders.", style="Hint.TLabel")
            return
        folder = Path(text)
        if not folder.is_dir():
            self.lbl_folder.config(text="That folder doesn't exist on this computer.", style="Bad.TLabel")
            return
        apply_look(self.collect_settings())
        units = list_units(folder)
        if not units:
            self.lbl_folder.config(text=f"No folders named like '{unit_word} 1' were found here.", style="Bad.TLabel")
            return
        folders = sum(len(list_unit_folders(folder, u)) for u in units)
        self.lbl_folder.config(text=f"Found {len(units)} {unit_word.lower()} folder(s) holding {folders} lesson, "
                                    f"activity, and exam folder(s).", style="Good.TLabel")

    # ------------------------------------------------------------ live sample
    def _schedule_sample(self) -> None:
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(150, self._render_sample)

    def _render_sample(self) -> None:
        self._preview_job = None
        if self.running or not hasattr(self, "txt_sample"):
            return
        apply_look(self.collect_settings())
        unit = f"{word('unit')} 6"
        number, name = "32", "networking basics"
        rows: list[tuple[str, int, str, str]] = [
            # (style, indent, title, kind note)
            ("module", 0, unit, ""),
            ("heading", ind("lesson_header"), lesson_header_title(number, name), "heading"),
            ("item", ind("page"), page_title(number, name), "page"),
            ("item", ind("slides"), slides_title(number, name), "page"),
            ("item", ind("checkpoint"), checkpoint_title(number, name), "quiz"),
            ("key", ind("checkpoint_key"), checkpoint_key_title(number), "file"),
            ("heading", ind("assessment_header"), assessment_header_title(unit, "activity"), "heading"),
            ("item", ind("assessment"), assessment_quiz_title(f"{unit} {word('activity')}", "activity"), "quiz"),
            ("key", ind("assessment_key"), assessment_key_title(unit, "activity"), "file"),
            ("heading", ind("assessment_header"), assessment_header_title(unit, "exam"), "heading"),
            ("item", ind("assessment"), assessment_quiz_title(f"{unit} {word('exam')}", "exam"), "quiz"),
            ("key", ind("assessment_key"), assessment_key_title(unit, "exam"), "file"),
            ("module", 0, PLAN_MODULE_NAME, ""),
            ("item", ind("plan"), plan_title(number), "file"),
        ]
        text = self.txt_sample
        text.config(state="normal")
        text.delete("1.0", "end")
        for style, level, title, kind in rows:
            if style == "module":
                text.insert("end", f"{title}\n", ("module",))
                continue
            tags = (f"ind{level}",) + (("heading",) if style == "heading" else ())
            text.insert("end", title, tags)
            if style == "key":
                text.insert("end", "  (hidden)", (f"ind{level}", "hidden"))
            text.insert("end", "\n", (f"ind{level}",))
        text.config(state="disabled")

    # ------------------------------------------------------------ browsing
    def _browse_folder(self) -> None:
        start = self.v_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(parent=self, title="Choose the folder that holds your unit folders",
                                         initialdir=start if Path(start).is_dir() else str(Path.home()))
        if chosen:
            self.v_folder.set(str(Path(chosen)))
            self._commit_folder()

    def _browse_sheet(self) -> None:
        start = self.v_folder.get().strip() or str(Path.home())
        chosen = filedialog.askopenfilename(parent=self, title="Choose your settings spreadsheet",
                                            initialdir=start if Path(start).is_dir() else str(Path.home()),
                                            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if chosen:
            self.v_sheet.set(str(Path(chosen)))

    # ------------------------------------------------------------ create a spreadsheet
    def _create_sheet_dialog(self) -> None:
        if self.running:
            return
        dialog = tk.Toplevel(self)
        dialog.title("Create a settings spreadsheet")
        dialog.transient(self)
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        choice = tk.StringVar(value="canvas")
        ttk.Label(frame, text="What should the new spreadsheet start with?", font=self.font_body
                  ).pack(anchor="w", pady=(0, 8))
        ttk.Radiobutton(frame, text="My Canvas course, filled in (recommended)", value="canvas",
                        variable=choice).pack(anchor="w")
        ttk.Label(frame, style="Hint.TLabel", wraplength=420, justify="left",
                  text="Lists your modules, items, assignments, dates, points, and requirements as they are now. "
                       "Canvas is only read, never changed.").pack(anchor="w", padx=(24, 0), pady=(0, 6))
        ttk.Radiobutton(frame, text="A blank template with examples", value="blank",
                        variable=choice).pack(anchor="w")
        ttk.Label(frame, style="Hint.TLabel", wraplength=420, justify="left",
                  text="For a course that isn't in Canvas yet.").pack(anchor="w", padx=(24, 0))
        ttk.Label(frame, style="Hint.TLabel", wraplength=420, justify="left",
                  text="Every row starts as Apply changes = NO. The file is saved in your course folder "
                       "with today's date, and existing spreadsheets are never replaced."
                  ).pack(anchor="w", pady=(12, 0))
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e", pady=(14, 0))

        def go() -> None:
            blank = choice.get() == "blank"
            dialog.destroy()
            self._create_sheet(blank)

        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="Create", style=self.accent, command=go).pack(side="right", padx=(0, 8))
        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.grab_set()

    def _create_sheet(self, blank: bool) -> None:
        if self.running:
            return
        self._commit_folder()
        s = self.collect_settings()
        folder = Path(s["course_folder"]) if s["course_folder"] else None
        problems = []
        if not folder or not folder.is_dir():
            problems.append("Choose your course folder first. The new spreadsheet is saved there.")
        if not blank:
            if not clean_canvas_url(self.v_url.get()) or not self.v_token.get().strip():
                problems.append("Connect to Canvas first: add your Canvas address and access token.")
            if not str(s["course_id"]).isdigit():
                problems.append("Add your course ID (the number after /courses/ in the course's web address).")
        if openpyxl is None:
            problems.append("Making a spreadsheet needs two free helper packages. Click Install when asked, "
                            "then reopen the program.")
        if problems:
            messagebox.showwarning("Almost ready", "\n\n".join(problems), parent=self)
            if openpyxl is None:
                self._offer_install()
            return
        try:
            self._save_credentials()
            self._save_course(s)
        except (OSError, CanvasError):
            pass
        apply_settings(s)
        set_connection(self.v_url.get(), self.v_token.get())
        course_id = str(s["course_id"]).strip()

        self.running = True
        self.btn_run.config(state="disabled")
        self.notebook.select(self.page_run)
        self.bar.start(12)
        self.lbl_summary.config(text="Making your spreadsheet…", style="TLabel")
        self._append("", None)
        self._append("Create settings spreadsheet" + (" (blank template)" if blank else " from Canvas"), "section")

        def work() -> None:
            try:
                path = create_settings_spreadsheet(folder, course_id, blank)
                self.events.put(("call", lambda: self._sheet_created(path, "")))
            except Exception as exc:  # Shown to the person; nothing in Canvas was changed.
                problem = friendly_error(exc) if isinstance(exc, (CanvasError, OSError)) else f"{type(exc).__name__}: {exc}"
                log(f"ERROR: {problem}")
                self.events.put(("call", lambda: self._sheet_created(None, problem)))

        threading.Thread(target=work, daemon=True).start()

    def _sheet_created(self, path: Path | None, problem: str) -> None:
        self.running = False
        self.btn_run.config(state="normal")
        self.bar.stop()
        if path is None:
            self.lbl_summary.config(text="The spreadsheet wasn't made. The red line explains why.", style="Bad.TLabel")
            messagebox.showerror("Couldn't make the spreadsheet",
                                 f"{problem}\n\nNothing in Canvas was changed.", parent=self)
            return
        self.lbl_summary.config(text=f"Made {path.name}. Nothing in Canvas was changed.", style="Good.TLabel")
        if messagebox.askyesno("Spreadsheet ready",
                               f"Your new spreadsheet is ready:\n\n{path.name}\n\n"
                               "Every row starts as NO, so nothing changes until you choose YES.\n\n"
                               "Use this spreadsheet for this course from now on?", parent=self):
            self.v_sheet.set(str(path))
            try:
                self._save_course()
            except (OSError, CanvasError):
                pass
        try:
            open_with_system(path)
        except OSError:
            pass

    # ------------------------------------------------------------ Google
    def _connect_google(self) -> None:
        if not google_ready():
            messagebox.showinfo("Google isn't set up yet",
                                f"To use Google Slides, the file {GOOGLE_CLIENT_PATH.name} must be in the same folder "
                                "as this program. Ask whoever shared Canvas Course Manager with you for it.\n\n"
                                "You can still use presentation links and 'Show them in Canvas' without it.", parent=self)
            return
        folder_text = self.v_drive_folder.get().strip()
        if folder_text and not drive_folder_id(folder_text):
            messagebox.showwarning("That isn't a folder link",
                                   "Open the folder in Google Drive, copy the address from the top of the browser, "
                                   "and paste it here.", parent=self)
            return
        self.lbl_google.config(text="Waiting for you to sign in to Google in your browser…", style="TLabel")

        def work() -> None:
            global DRIVE_FOLDER
            ok = True
            try:
                DRIVE_FOLDER = folder_text
                google_access_token(google_scope())
                message = "Connected to Google."
                if folder_text:
                    _, info = google_http("GET", f"{DRIVE_API}/files/{drive_folder_id(folder_text)}"
                                                 "?fields=name&supportsAllDrives=true")
                    message += f" Drive folder: {info.get('name', '?')}."
            except CanvasError as exc:
                ok, message = False, str(exc)
            self.events.put(("call", lambda: self.lbl_google.config(
                text=message, style="Good.TLabel" if ok else "Bad.TLabel")))

        threading.Thread(target=work, daemon=True).start()
        self._save_setup()

    # ------------------------------------------------------------ connection test
    def _test_connection(self) -> None:
        url = clean_canvas_url(self.v_url.get())
        token = self.v_token.get().strip()
        if not url or not token:
            messagebox.showwarning("Almost there", "Fill in both the Canvas address and the access token first.", parent=self)
            return
        self.lbl_conn.config(text="Checking…", style="TLabel")
        cid = course_id_from_text(self.v_course.get())

        def work() -> None:
            ok, message, course_name = True, "", ""
            groups: tuple[list[dict[str, Any]], bool] | None = None
            try:
                set_connection(url, token)
                me = api_request("GET", "/users/self")
                message = f"Connected as {me.get('name', 'you')}."
                if cid.isdigit():
                    course = api_request("GET", f"/courses/{cid}", allow_404=True)
                    if course and course.get("id"):
                        course_name = str(course.get("name", ""))
                        message += f" Course {cid}: {course_name}."
                        groups = read_assignment_groups(cid)   # Read-only, for the gradebook group choices.
                    else:
                        ok = False
                        message += f" Course {cid} wasn't found, or you aren't a teacher in it."
            except CanvasError as exc:
                ok, message = False, friendly_error(exc)
            self.events.put(("call", lambda: self._connection_result(ok, message, course_name, url, token,
                                                                     cid, groups)))

        threading.Thread(target=work, daemon=True).start()

    def _connection_result(self, ok: bool, message: str, course_name: str, url: str, token: str,
                           cid: str = "", groups: tuple[list[dict[str, Any]], bool] | None = None) -> None:
        self.lbl_conn.config(text=message, style="Good.TLabel" if ok else "Bad.TLabel")
        if course_name:
            self.course_name = course_name
            self._update_course_label()
        if ok:
            try:
                self.v_url.set(url)
                self._save_credentials()
                self._save_setup()
            except OSError:
                pass
        if ok and groups is not None:
            self._groups_loaded(cid, groups[0], groups[1], "")

    # ------------------------------------------------------------ gradebook groups
    def _kind_plural(self, kind: str) -> str:
        """'Checkpoints', 'Activities', 'Exams', using the words on the Names & look tab."""
        text = self.v_words[kind].get().strip() or DEFAULT_SETTINGS["words"][kind]
        return plural(text)

    def _group_settings(self) -> dict[str, Any]:
        """Saved group choices, but only when they were saved for the course ID in the box now."""
        saved = self.settings.get("assignment_groups") or {}
        if str(saved.get("course_id") or "") != course_id_from_text(self.v_course.get()):
            return {}
        return saved

    def _set_group_choice(self, kind: str, value: dict[str, Any] | None) -> None:
        cid = course_id_from_text(self.v_course.get())
        saved = self.settings.get("assignment_groups")
        if not isinstance(saved, dict) or str(saved.get("course_id") or "") != cid:
            saved = {"course_id": cid, **{k: None for k in QUIZ_KINDS}}
            self.settings["assignment_groups"] = saved
        saved[kind] = value

    def _save_group_choices(self) -> None:
        try:
            if not self._save_course():
                self._status("Choose your course folder on this tab so the gradebook groups can be saved.")
        except (OSError, CanvasError):
            pass

    def _groups_ready(self) -> bool:
        return bool(self.groups_course) and self.groups_course == course_id_from_text(self.v_course.get())

    def _group_label(self, group: dict[str, Any]) -> str:
        name = _cell_text(group.get("name")) or "(no name)"
        twins = sum(1 for g in self.canvas_groups if norm(_cell_text(g.get("name"))) == norm(name))
        text = name + (f" (ID {group.get('id')})" if twins > 1 else "")
        if self.groups_weighted:
            text += f"  ·  {percent_text(group.get('group_weight'))}"
        return text

    def _refresh_group_widgets(self) -> None:
        if not hasattr(self, "cmb_new_group_for"):
            return
        leave = "Leave where Canvas puts it"
        self.group_by_label = {leave: None}
        if self._groups_ready():
            for group in self.canvas_groups:
                self.group_by_label[self._group_label(group)] = group
        choices = self._group_settings()
        for kind in QUIZ_KINDS:
            self.group_kind_labels[kind].config(text=f"New {self._kind_plural(kind).lower()}")
            values = list(self.group_by_label)
            choice = choices.get(kind)
            label = ""
            if isinstance(choice, dict) and choice.get("leave"):
                label = leave
            elif isinstance(choice, dict) and choice.get("id") is not None:
                label = next((text for text, g in self.group_by_label.items()
                              if g and str(g.get("id")) == str(choice["id"])), "")
                if not label:   # Not loaded yet, or gone from Canvas.
                    label = _cell_text(choice.get("name")) + ("  (not found in Canvas)" if self._groups_ready() else "")
                    values.append(label)
            self.cmb_groups[kind].config(values=values)
            self.v_group[kind].set(label)
        self.new_group_kinds = {f"New {self._kind_plural(k).lower()}": k for k in QUIZ_KINDS}
        self.new_group_kinds["Nothing yet"] = None
        self.cmb_new_group_for.config(values=list(self.new_group_kinds))
        if self.v_new_group_for.get() not in self.new_group_kinds:
            first_unset = next((k for k in QUIZ_KINDS if not isinstance(choices.get(k), dict)), None)
            self.v_new_group_for.set(next(label for label, k in self.new_group_kinds.items() if k == first_unset))
        self._show_group_summary()

    def _show_group_summary(self) -> None:
        if not hasattr(self, "lbl_groups"):
            return
        choices = self._group_settings()
        unset = [self._kind_plural(k).lower() for k in QUIZ_KINDS if not isinstance(choices.get(k), dict)]
        if not self._groups_ready():
            text = "Groups not loaded yet. Click Load groups from Canvas (it only reads your course)."
            if any(isinstance(choices.get(k), dict) for k in QUIZ_KINDS):
                text = "Your saved choices are shown. Click Load groups from Canvas to see every group."
            self.lbl_groups.config(text=text, style="Hint.TLabel")
            return
        text = f"{len(self.canvas_groups)} group(s) in this Canvas course. "
        if self.groups_weighted:
            total = sum(float(g.get("group_weight") or 0) for g in self.canvas_groups)
            text += f"Weighted grading is on; the groups add up to {percent_text(total)}."
        else:
            text += "Weighted grading is off, so percentages don't change grades."
        if unset:
            text += (f"  Not chosen yet: {', '.join(unset)}. Pick a group above, or add a new one below "
                     "(until then they stay where Canvas puts them).")
        self.lbl_groups.config(text=text, style="Warn.TLabel" if unset else "TLabel")

    def _load_groups(self, quiet: bool = False, then: Callable[[], None] | None = None) -> None:
        """Read the course's gradebook groups from Canvas (GET only)."""
        if self.loading_groups or self.running:
            return
        url, token = clean_canvas_url(self.v_url.get()), self.v_token.get().strip()
        cid = course_id_from_text(self.v_course.get())
        if not url or not token or not cid.isdigit():
            if not quiet:
                messagebox.showwarning("Almost there", "Add your Canvas address, access token, and course ID first.",
                                       parent=self)
            return
        self.loading_groups = True
        self.lbl_groups.config(text="Reading gradebook groups from Canvas…", style="TLabel")

        def work() -> None:
            try:
                set_connection(url, token)
                groups, weighted = read_assignment_groups(cid)
                self.events.put(("call", lambda: self._groups_loaded(cid, groups, weighted, "", then)))
            except CanvasError as exc:
                problem = friendly_error(exc)
                self.events.put(("call", lambda: self._groups_loaded(cid, [], False, problem, then)))

        threading.Thread(target=work, daemon=True).start()

    def _groups_loaded(self, cid: str, groups: list[dict[str, Any]], weighted: bool, problem: str,
                       then: Callable[[], None] | None = None) -> None:
        self.loading_groups = False
        if problem:
            self.lbl_groups.config(text=f"Couldn't read the gradebook groups: {problem}", style="Bad.TLabel")
            return
        if cid != course_id_from_text(self.v_course.get()):
            return   # The course ID changed while Canvas was being read.
        self.canvas_groups, self.groups_weighted, self.groups_course = groups, weighted, cid
        choices = self._group_settings()
        by_id = {str(g.get("id")): g for g in groups}
        changed = False
        for kind in QUIZ_KINDS:
            choice = choices.get(kind)
            if isinstance(choice, dict) and choice.get("id") is not None:
                group = by_id.get(str(choice["id"]))
                if group is None:   # Recreated in Canvas? Follow it only if exactly one group has its name.
                    twins = [g for g in groups if norm(_cell_text(g.get("name"))) == norm(_cell_text(choice.get("name")))]
                    group = twins[0] if len(twins) == 1 else None
                if group is not None and (str(group.get("id")) != str(choice["id"])
                                          or _cell_text(group.get("name")) != choice.get("name")):
                    self._set_group_choice(kind, {"id": str(group["id"]), "name": _cell_text(group.get("name"))})
                    changed = True
            elif not isinstance(choice, dict):
                # Never chosen: a group named like the quiz kind ("Checkpoint" or "Checkpoints") is picked.
                names = {norm(self._kind_plural(kind)),
                         norm(self.v_words[kind].get().strip() or DEFAULT_SETTINGS["words"][kind])}
                matches = [g for g in groups if norm(_cell_text(g.get("name"))) in names]
                if len(matches) == 1:
                    self._set_group_choice(kind, {"id": str(matches[0]["id"]), "name": _cell_text(matches[0].get("name"))})
                    changed = True
        if changed:
            self._save_group_choices()
        self._refresh_group_widgets()
        if then is not None:
            self.after(50, then)

    def _group_picked(self, kind: str) -> None:
        label = self.v_group[kind].get()
        if label not in self.group_by_label:
            return   # The "(not found in Canvas)" line: nothing new was chosen.
        group = self.group_by_label[label]
        if group is None:
            self._set_group_choice(kind, {"leave": True})
        else:
            if self.groups_weighted and not float(group.get("group_weight") or 0):
                messagebox.showwarning(
                    "This group counts 0%",
                    f"\"{group.get('name')}\" counts 0% of the grade in this course, so {self._kind_plural(kind).lower()} "
                    "in it won't affect grades.\n\nIf that isn't what you want, change its percentage in Canvas "
                    "or choose another group.", parent=self)
            self._set_group_choice(kind, {"id": str(group["id"]), "name": _cell_text(group.get("name"))})
        self._save_group_choices()
        self._refresh_group_widgets()
        self._status(f"Saved: new {self._kind_plural(kind).lower()} go to {label.split('  ·  ')[0]}.")

    def _show_group_section(self) -> None:
        """Open the Course setup tab at the gradebook groups."""
        self.notebook.select(self.page_setup)
        self.after(50, lambda: self.page_canvases[1].yview_moveto(1.0))
        if not self._groups_ready():
            self._load_groups(quiet=True)

    def _suggest_group_name(self) -> None:
        """Choosing what a new group is for fills in a name, if the name box is empty."""
        kind = self.new_group_kinds.get(self.v_new_group_for.get())
        if kind and not self.v_new_group_name.get().strip():
            self.v_new_group_name.set(self._kind_plural(kind))

    def _add_group_clicked(self) -> None:
        """Add one gradebook group from the boxes on the Course setup tab (preview, then APPLY)."""
        if self.running:
            return
        if not self._groups_ready():   # The current groups are needed to avoid duplicates and add up percentages.
            self.lbl_new_group.config(text="Reading your course's groups first…", style="TLabel")
            self._load_groups(then=self._add_group_clicked)
            return
        name = " ".join(self.v_new_group_name.get().split())
        weight_text = self.v_new_group_weight.get().strip().rstrip("%").strip()
        kind = self.new_group_kinds.get(self.v_new_group_for.get())
        problem = ""
        try:
            weight = float(weight_text)
        except ValueError:
            weight = -1.0
        if not name:
            problem = "Type a name for the new group."
        elif len(name) > 255:
            problem = "That name is too long for Canvas."
        elif not 0 <= weight <= 100 or math.isnan(weight):
            problem = "Type a percentage from 0 to 100, like 20."
        if problem:
            self.lbl_new_group.config(text=problem, style="Bad.TLabel")
            return
        existing = next((g for g in self.canvas_groups if norm(_cell_text(g.get("name"))) == norm(name)), None)
        if existing is not None:   # Never make a second group with the same name.
            if kind:
                self._set_group_choice(kind, {"id": str(existing["id"]), "name": _cell_text(existing.get("name"))})
                self._save_group_choices()
                self._refresh_group_widgets()
            self.lbl_new_group.config(
                text=f"\"{existing.get('name')}\" is already in Canvas, so nothing was added"
                     + (f"; new {self._kind_plural(kind).lower()} will go there." if kind else ".")
                     + " Its percentage wasn't changed.", style="Warn.TLabel")
            return
        cid = course_id_from_text(self.v_course.get())
        lines = [f"This gradebook group will be added to course {cid}"
                 + (f" ({self.course_name})" if self.course_name else "") + ":", "",
                 f"   {name}, {percent_text(weight)}"
                 + (f"   (for new {self._kind_plural(kind).lower()})" if kind else ""),
                 "", "Nothing already in Canvas is changed or deleted."]
        if self.groups_weighted:
            total = sum(float(g.get("group_weight") or 0) for g in self.canvas_groups) + weight
            if abs(total - 100) > 0.01:
                lines += ["", f"Note: all your groups would add up to {percent_text(total)}. Canvas allows this, "
                              "but you'll probably want to adjust the percentages in Canvas so they add up to 100%."]
            if not weight:
                lines += ["", "WARNING: a group set to 0% won't count toward grades."]
        else:
            lines += ["", "This course doesn't use weighted grading, so the percentage won't change grades "
                          "unless weighting is turned on in Canvas. The program never changes that setting."]
        if not messagebox.askokcancel("Preview: new gradebook group", "\n".join(lines), parent=self):
            return
        typed = simpledialog.askstring("Confirm", "Type APPLY (in capitals) to add this group to your Canvas course:",
                                       parent=self)
        if (typed or "").strip() != "APPLY":
            self.lbl_new_group.config(text="Cancelled. Nothing was changed.", style="TLabel")
            return
        apply_settings(self.collect_settings())   # So the record of this lands in this course's log.
        set_connection(self.v_url.get(), self.v_token.get())
        self.lbl_new_group.config(text=f"Adding \"{name}\" in Canvas…", style="TLabel")

        def work() -> None:
            try:
                group = create_assignment_group(cid, name, weight)
                self.events.put(("call", lambda: self._group_created(cid, kind, group, "")))
            except CanvasError as exc:
                problem_text = friendly_error(exc)
                log(f"ERROR: {problem_text}")
                self.events.put(("call", lambda: self._group_created(cid, kind, None, problem_text)))

        threading.Thread(target=work, daemon=True).start()

    def _group_created(self, cid: str, kind: str | None, group: dict[str, Any] | None, problem: str) -> None:
        if group is None:
            self.lbl_new_group.config(text=f"The group wasn't added: {problem}", style="Bad.TLabel")
            return
        if cid == course_id_from_text(self.v_course.get()):
            self.canvas_groups.append(group)
            if kind:
                self._set_group_choice(kind, {"id": str(group["id"]), "name": _cell_text(group.get("name"))})
                self._save_group_choices()
        self.v_new_group_name.set("")
        self.v_new_group_weight.set("")
        self.v_new_group_for.set("")   # Moves on to the next kind still without a group.
        self._refresh_group_widgets()
        self.lbl_new_group.config(
            text=f"Added \"{group.get('name')}\" ({percent_text(group.get('group_weight'))}) to Canvas"
                 + (f" and chose it for new {self._kind_plural(kind).lower()}." if kind else ".")
                 + " To move quizzes uploaded before today, click Sort existing quizzes.", style="Good.TLabel")
        self._load_groups(quiet=True)   # Show Canvas's own view of the new group.

    def _start_sort(self, live: bool = False) -> None:
        """Move already-uploaded quizzes out of "Imported Assignments". Always previewed first."""
        if self.running:
            return
        self._commit_folder()
        s = self.collect_settings()
        problems = []
        if not clean_canvas_url(self.v_url.get()) or not self.v_token.get().strip():
            problems.append("Connect to Canvas first: add your Canvas address and access token.")
        if not str(s["course_id"]).isdigit():
            problems.append("Add your course ID on this tab.")
        if not s["course_folder"] or not Path(s["course_folder"]).is_dir():
            problems.append("Choose your course folder. Its upload history says which quizzes this program made.")
        choices = self._group_settings()
        if not any(isinstance(choices.get(k), dict) and choices[k].get("id") for k in QUIZ_KINDS):
            problems.append("Choose a gradebook group for at least one kind of quiz first.")
        if problems:
            messagebox.showwarning("Almost ready", "\n\n".join(problems), parent=self)
            return
        if live:
            if not messagebox.askokcancel("Move quizzes in Canvas?",
                                          "Quizzes this program uploaded that are still in "
                                          f"\"{IMPORTED_GROUP_NAME}\" will move into the groups you chose "
                                          "(the blue lines in the preview). Nothing else changes.", parent=self):
                return
            typed = simpledialog.askstring("Confirm live changes",
                                           "Type APPLY (in capitals) to change your Canvas course:", parent=self)
            if (typed or "").strip() != "APPLY":
                messagebox.showinfo("Cancelled", "Nothing was changed.", parent=self)
                return
        try:
            self._save_credentials()
            self._save_course(s)
        except (OSError, CanvasError):
            pass
        apply_settings(s)
        set_connection(self.v_url.get(), self.v_token.get())
        configure_sort(live)
        self.job = "sort"
        self.notebook.select(self.page_run)
        self._set_running(True, live)
        self._append("", None)
        self._append("Sort existing quizzes into gradebook groups" + ("" if live else " (preview)"), "section")
        threading.Thread(target=self._worker, args=(live, sort_existing_quizzes), daemon=True).start()

    # ------------------------------------------------------------ running
    def _run_options(self) -> dict[str, Any]:
        task = self.v_task.get()
        if task == "custom":
            types = {k for k, v in self.v_types.items() if v.get()}
            sheet = self.v_custom_sheet.get()
        else:
            preset_types, sheet = TASK_PRESETS[task]
            types = set(preset_types)
        return {"task": task, "types": types, "spreadsheet": sheet, "scope": self.v_scope.get(),
                "unit": self.v_unit.get(), "lesson": self.v_lesson.get()}

    def _problems(self, s: dict[str, Any], options: dict[str, Any]) -> list[str]:
        problems: list[str] = self._emoji_problems(s)
        if not clean_canvas_url(self.v_url.get()) or not self.v_token.get().strip():
            problems.append("Connect to Canvas first: add your Canvas address and access token on the Course setup tab.")
        if not str(s["course_id"]).isdigit():
            problems.append("Add your course ID on the Course setup tab (the number after /courses/ in the course's web address).")
        if not options["types"] and not options["spreadsheet"]:
            problems.append("Choose at least one thing to upload under \"Let me choose\".")
        if options["types"]:
            folder = Path(s["course_folder"]) if s["course_folder"] else None
            if not folder or not folder.is_dir():
                problems.append("Choose your course folder on the Course setup tab.")
            elif options["scope"] in ("lesson", "unit") and not options["unit"]:
                problems.append(f"Choose a {word('unit').lower()} under \"Where should it look?\".")
            elif options["scope"] == "lesson" and not options["lesson"]:
                problems.append("Choose a folder under \"Where should it look?\".")
        if options["spreadsheet"]:
            if pd is None:
                problems.append("The settings spreadsheet needs two free helper packages. Click Install when asked, "
                                "then reopen the program.")
            if not s["spreadsheet"] or not Path(s["spreadsheet"]).is_file():
                problems.append("Choose your settings spreadsheet on the Course setup tab.")
        return problems

    def _summary_text(self, s: dict[str, Any], options: dict[str, Any], live: bool) -> str:
        names = {"plans": "lesson plans", "headers": f"{word('lesson').lower()} headings", "html": "HTML pages",
                 "slides": {"canvas": "presentations (shown in Canvas)", "google": "presentations (as Google Slides)",
                            "off": "presentation links"}[SLIDES_MODE],
                 "checkpoints": plural(word("checkpoint")).lower(), "activities": plural(word("activity")).lower(),
                 "exams": plural(word("exam")).lower(), "keys": "answer keys (hidden)"}
        types = effective_types(options)
        lines = [f"Course: {s['course_id']}" + (f"  ({self.course_name})" if self.course_name else "")]
        if types:
            where = {"lesson": f"{options['unit']} / {options['lesson']}", "unit": f"all of {options['unit']}",
                     "course": "the entire course"}[options["scope"]]
            lines.append("Upload: " + ", ".join(names[k] for k in UPLOAD_TYPES if k in types))
            lines.append(f"Where: {where}")
        lines.append("Settings spreadsheet: " + (Path(s["spreadsheet"]).name if options["spreadsheet"] else "no"))
        lines.append("Mode: " + ("LIVE. Canvas will be changed." if live else "Preview. Nothing in Canvas changes."))
        return "\n".join(lines)

    def _start_run(self, live: bool | None = None, skip_sheet_check: bool = False) -> None:
        if self.running:
            return
        self._commit_folder()
        s = self.collect_settings()
        options = self._run_options()
        apply_look(s)
        problems = self._problems(s, options)
        if problems:
            messagebox.showwarning("Almost ready", "\n\n".join(problems), parent=self)
            if pd is None and options["spreadsheet"]:
                self._offer_install()
            return
        try:
            self._save_credentials()
            self._save_course(s)
        except (OSError, CanvasError):
            pass
        apply_settings(s)
        set_connection(self.v_url.get(), self.v_token.get())
        live = (self.v_mode.get() == "live") if live is None else live

        unset = [k for k in quiz_kinds_in(options) if not isinstance(self._group_settings().get(k), dict)]
        if unset and not skip_sheet_check:
            names = ", ".join(self._kind_plural(k).lower() for k in unset)
            if messagebox.askyesno(
                    "Gradebook groups",
                    f"New {names} don't have a gradebook group yet, so Canvas will put them in "
                    f"\"{IMPORTED_GROUP_NAME}\" (which often counts 0% in a weighted course).\n\n"
                    "Set up their groups first? Nothing has run yet.\n\n"
                    "Yes: take me to the gradebook groups on the Course setup tab.\n"
                    "No: run anyway and leave them where Canvas puts them.",
                    parent=self):
                self._show_group_section()
                return

        if options["spreadsheet"] and not skip_sheet_check:
            try:
                ids, rows = spreadsheet_course_ids()
            except Exception as exc:
                messagebox.showerror("Can't read the spreadsheet", f"{exc}\n\nIs it open in Excel? Close it and try again.",
                                     parent=self)
                return
            message = (f"The spreadsheet has {rows} row(s) marked YES.\n\n"
                       "The course ID in the spreadsheet wins over the course ID on the Course setup tab.\n\n"
                       f"Course setup tab: {s['course_id']}\n"
                       f"Spreadsheet: {', '.join(sorted(ids)) if ids else 'blank (it uses the course setup ID)'}")
            if any(value != str(s["course_id"]) for value in ids):
                message += "\n\nWARNING: at least one row points at a DIFFERENT course."
            message += "\n\nHave you checked the course ID and want to continue?"
            if not messagebox.askyesno("Check the spreadsheet course IDs", message,
                                       icon="warning" if ids - {str(s["course_id"])} else "question", parent=self):
                return

        if live:
            if not messagebox.askokcancel("Apply changes to Canvas?", self._summary_text(s, options, live), parent=self):
                return
            typed = simpledialog.askstring("Confirm live changes", "Type APPLY (in capitals) to change your Canvas course:",
                                           parent=self)
            if (typed or "").strip() != "APPLY":
                messagebox.showinfo("Cancelled", "Nothing was changed.", parent=self)
                return

        configure_run(options, live)
        self.job = "run"
        self._set_running(True, live)
        self._append("", None)
        for line in self._summary_text(s, options, live).splitlines():
            self._append(line, "section")
        threading.Thread(target=self._worker, args=(live,), daemon=True).start()

    def _worker(self, live: bool, job: Callable[[], None] = run_job) -> None:
        ok = True
        try:
            job()
        except (CanvasError, OSError, ValueError, KeyError) as exc:
            ok = False
            log(f"ERROR: {friendly_error(exc)}")
            log("ERROR: Stopped here so nothing is left half-finished. Fix the problem above, then run again.")
        except Exception as exc:  # last-resort safety net, with details for whoever helps
            ok = False
            log(f"ERROR (unexpected): {type(exc).__name__}: {exc}")
            for line in traceback.format_exc().splitlines():
                log("ERROR    " + line)
        self.events.put(("done", ok, live, dict(COUNTS)))

    def _set_running(self, running: bool, live: bool = False) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        self.btn_run.config(state=state)
        for index in (1, 2):
            self.notebook.tab(index, state=state)
        if running:
            self.bar.start(12)
            self.badge.config(text="Applying changes…" if live else "Previewing…", bg=LIVE_RED if live else PREVIEW_GREEN)
            self.lbl_summary.config(text="Working… long quiz imports can take a few minutes.", style="TLabel")
            self._status("Please leave the window open until the run finishes.")
        else:
            self.bar.stop()
            self._on_mode_change()

    def _on_done(self, ok: bool, live: bool, counts: dict[str, int]) -> None:
        self._set_running(False)
        would, warn = counts["would"], counts["warn"]
        if not ok:
            self.lbl_summary.config(text="The run stopped. The red lines explain why.", style="Bad.TLabel")
            self._status("")
            messagebox.showerror("The run stopped",
                                 "The red lines in the progress panel explain what needs fixing.\n\n"
                                 "Steps after that point were not run.", parent=self)
            return
        if live:
            self.lbl_summary.config(text=f"Finished. Changes applied. Warnings: {warn}.", style="Good.TLabel")
            self._status("Check the course in Canvas with Student View to see it the way students will.")
            messagebox.showinfo("Finished", "Your changes were applied to Canvas.\n\n"
                                + (f"{warn} orange line(s) are worth a look." if warn else "No warnings."), parent=self)
            self.v_mode.set("preview")
            if self.job == "sort":
                self._load_groups(quiet=True)
            return
        self.lbl_summary.config(text=f"Preview finished: {would} change(s) would be made, {warn} warning(s). "
                                     "Nothing in Canvas was changed.", style="TLabel")
        self._status("")
        if would == 0:
            messagebox.showinfo("Preview finished", "Everything here is already up to date in Canvas. "
                                "There is nothing to change." + (f"\n\n{warn} orange line(s) are worth a look." if warn else ""),
                                parent=self)
            return
        question = (f"Preview finished. Nothing in Canvas was changed.\n\nA live run would make {would} change(s) "
                    "(the blue lines)." + (f"\n{warn} orange line(s) are worth a look first." if warn else "")
                    + "\n\nApply these changes to Canvas now?")
        if messagebox.askyesno("Preview finished", question, parent=self):
            if self.job == "sort":
                self._start_sort(live=True)
            else:
                self._start_run(live=True, skip_sheet_check=True)

    # ------------------------------------------------------------ progress panel
    def _tag_for(self, message: str) -> str | None:
        if message.startswith("ERROR") or "Stopped" in message:
            return "error"
        if "WARNING" in message or "SAFETY SKIP" in message or "skipped —" in message or message.startswith("Skipping"):
            return "warn"
        if "DRY RUN: would" in message:
            return "would"
        if message.startswith(("=", "Processing", "Starting", "Finished", "Files selected", "Folders to process")):
            return "section"
        if any(word_ in message for word_ in ("Verified", "Added", "Created", "Linked", "Uploading", "Re-uploading",
                                              "Hidden from students", "Renamed", "Updated", "now points")):
            return "done"
        return None

    def _append(self, message: str, tag: str | None) -> None:
        display = re.sub(r"^(SETTINGS )?DRY RUN: would ", "Would ", message)
        display = re.sub(r"^SPREADSHEET SETTINGS: ", "", display)
        self.txt_log.config(state="normal")
        if message:
            self.txt_log.insert("end", time.strftime("%H:%M:%S  "), ("time",))
        self.txt_log.insert("end", display + "\n", (tag,) if tag else ())
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _clear_log(self) -> None:
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self.lbl_summary.config(text="Nothing has run yet.", style="TLabel")

    def _copy_log(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.txt_log.get("1.0", "end"))
        self._status("Progress copied. You can paste it into an email.")

    def _open_log(self) -> None:
        apply_settings(self.collect_settings())
        if LOG_PATH.exists():
            try:
                open_with_system(LOG_PATH)
            except OSError as exc:
                messagebox.showerror("Couldn't open the log", str(exc), parent=self)
        else:
            messagebox.showinfo("No log yet", "The log file appears after your first run.", parent=self)

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    self._append(event[1], self._tag_for(event[1]))
                elif event[0] == "call":
                    event[1]()
                elif event[0] == "done":
                    self._on_done(event[1], event[2], event[3])
        except queue.Empty:
            pass
        self.after(80, self._drain_events)

    # ------------------------------------------------------------ start and close
    def report_callback_exception(self, exc_type: Any, exc: BaseException, tb: Any) -> None:
        """Show surprises in a window; the no-console launcher would otherwise hide them."""
        details = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] WINDOW ERROR\n{details}\n")
        except OSError:
            pass
        messagebox.showerror(APP_TITLE, f"Something unexpected happened:\n\n{exc}\n\n"
                             "The details were added to canvas_lesson_uploader.log in your course folder.", parent=self)

    def _startup(self) -> None:
        if getattr(self, "old_creds_path", None):
            self._append(f"Your Canvas login was copied next to the program. You can delete the old copy in "
                         f"{self.old_creds_path.parent.name} ({self.old_creds_path.name}) so it isn't shared "
                         "by accident.", "section")
        if self.current_folder and self.course_saved and not (Path(self.current_folder) / COURSE_SETTINGS_NAME).is_file():
            try:
                self._save_course()   # First start after updating: move the older setup into the course folder.
                self._append(f"Your course setup now lives in {Path(self.current_folder).name}/"
                             f"{COURSE_SETTINGS_NAME}.", "section")
            except (OSError, CanvasError):
                pass
        if not self.v_url.get() or not self.v_token.get() or not self.v_course.get():
            self.notebook.select(self.page_setup)
            self._append("Welcome! Fill in the Course setup tab first, then choose how titles should look on the "
                         "Names & look tab. Your answers are saved, so this only happens once.", "section")
        if not self.current_folder:
            self.notebook.select(self.page_setup)
            if messagebox.askyesno(APP_TITLE, "Which course folder should I work with?\n\nChoose the folder that "
                                   "holds your Unit folders (for example \"Unit 1\", \"Unit 2\"). You can switch "
                                   "courses any time on the Course setup tab.\n\nChoose it now?", parent=self):
                self._browse_folder()
        if pd is None:
            self._offer_install()
        if self.current_folder and self.v_url.get() and self.v_token.get() and self.v_course.get():
            self.after(600, lambda: self._load_groups(quiet=True))   # Read-only; just fills the Course setup tab.
        if GITHUB_REPO and not is_developer_copy():
            self.after(2000, lambda: self._check_updates(manual=False))

    # ------------------------------------------------------------ updates
    def _check_updates(self, manual: bool) -> None:
        if self.checking_update:
            return
        if self.running:
            if manual:
                messagebox.showinfo("Updates", "Please wait until the current run finishes.", parent=self)
            return
        if not GITHUB_REPO:
            if manual:
                messagebox.showinfo("Updates", "Updates aren't turned on for this copy of the program.\n\n"
                                    "Ask whoever shared it with you for the newest version.", parent=self)
            return
        self.checking_update = True
        if manual:
            self._status("Checking for updates...")

        def work() -> None:
            try:
                release, problem = latest_release(), ""
            except UpdateError as exc:
                release, problem = None, str(exc)
            self.events.put(("call", lambda: self._update_checked(release, problem, manual)))

        threading.Thread(target=work, daemon=True).start()

    def _update_checked(self, release: dict[str, str] | None, problem: str, manual: bool) -> None:
        self.checking_update = False
        write_update_status(last_check=datetime.now().isoformat(timespec="seconds"))
        if release is None:
            if manual:   # Automatic checks stay silent: no internet should never nag anyone.
                self._status("")
                messagebox.showwarning("Updates", problem, parent=self)
            return
        if not is_newer(release["version"]):
            if manual:
                self._status("")
                messagebox.showinfo("Updates", f"You have the newest version ({__version__}).", parent=self)
            return
        if not manual and read_update_status().get("skipped_version") == release["version"]:
            return
        self._status(f"Version {release['version']} is available.")
        self._update_dialog(release)

    def _update_dialog(self, release: dict[str, str]) -> None:
        win = tk.Toplevel(self)
        win.title("A new version is available")
        win.transient(self)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Version {release['version']} is ready (you have {__version__}).",
                  font=self.font_head).pack(anchor="w")
        ttk.Label(frame, wraplength=460, justify="left", text=(
            "Updating replaces only the program file. Your Canvas connection, course settings, spreadsheets, "
            "and lesson folders are not touched, and your current version is saved in the \"backups\" folder "
            "in case you need it again.")).pack(anchor="w", pady=(6, 10))
        if release["notes"]:
            ttk.Label(frame, text="What's new:").pack(anchor="w")
            notes_box, notes = self._text_box(frame, height=9, width=60, wrap="word", font=self.font_small,
                                              padx=8, pady=6)
            notes.insert("1.0", release["notes"])
            notes.config(state="disabled")
            notes_box.pack(fill="both", pady=(2, 10))
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e")

        def choose(action: str) -> None:
            self.text_widgets = [w for w in self.text_widgets if w.winfo_exists() and w.winfo_toplevel() is not win]
            win.destroy()
            if action == "skip":
                write_update_status(skipped_version=release["version"])
                self._status("")
            elif action == "update":
                self._install_update(release)

        ttk.Button(buttons, text="Skip this version", command=lambda: choose("skip")).pack(side="left")
        ttk.Button(buttons, text="Not now", command=lambda: choose("later")).pack(side="left", padx=8)
        ttk.Button(buttons, text="Update now", style=self.accent,
                   command=lambda: choose("update")).pack(side="left")
        win.protocol("WM_DELETE_WINDOW", lambda: choose("later"))
        win.grab_set()

    def _install_update(self, release: dict[str, str]) -> None:
        if self.running:
            messagebox.showinfo("Updates", "Please wait until the current run finishes.", parent=self)
            return
        try:
            self._commit_folder()
            self._save_course()
        except (OSError, CanvasError):
            pass
        self._status(f"Downloading version {release['version']}...")
        self.config(cursor="watch")

        def work() -> None:
            try:
                backup, problem = install_program(download_program(release)), ""
            except UpdateError as exc:
                backup, problem = None, str(exc)
            self.events.put(("call", lambda: self._update_installed(release, backup, problem)))

        threading.Thread(target=work, daemon=True).start()

    def _update_installed(self, release: dict[str, str], backup: Path | None, problem: str) -> None:
        self.config(cursor="")
        if backup is None:
            self._status("")
            messagebox.showerror("Update didn't finish",
                                 f"{problem}\n\nNothing was changed. Your current version still works.", parent=self)
            return
        log(f"Updated program from version {__version__} to {release['version']}. Old copy saved as {backup}.")
        write_update_status(skipped_version="")
        self._status(f"Version {release['version']} is installed.")
        if messagebox.askyesno("Update installed",
                               f"Version {release['version']} is installed.\n\n"
                               "Restart the program now to start using it?", parent=self):
            try:
                restart_program()
            except OSError:
                messagebox.showinfo("Almost done", "Please close this window and open the program again.",
                                    parent=self)
                return
            self.destroy()
        else:
            messagebox.showinfo("Update installed", "The new version starts the next time you open the program.",
                                parent=self)

    def _offer_install(self) -> None:
        packages = ["pandas", "openpyxl"] + ([] if sv_ttk is not None else ["sv-ttk"])
        if not messagebox.askyesno(
                "One-time setup",
                "The settings spreadsheet needs two free helper packages (pandas and openpyxl)"
                + (", and the modern look needs one more (sv-ttk)" if sv_ttk is None else "") + ".\n\n"
                "Install them now? It needs an internet connection and takes about a minute.\n\n"
                "(You'll need them to use settings spreadsheets.)", parent=self):
            return
        self._pip_install(packages)

    def _offer_modern_look(self) -> None:
        if self.running:
            messagebox.showinfo("Modern look", "Please wait until the current run finishes.", parent=self)
            return
        if not messagebox.askyesno(
                "Modern look",
                "A cleaner, modern window (with an optional dark mode) needs one small free helper "
                "package called sv-ttk.\n\nInstall it now? It needs an internet connection and takes "
                "a few seconds. Nothing about your course or settings changes.", parent=self):
            return
        self._pip_install(["sv-ttk"])

    def _pip_install(self, packages: list[str]) -> None:
        names = ", ".join(packages)
        self._append(f"Installing {names}. This can take a minute…", "section")
        self.bar.start(12)

        def work() -> None:
            exe = Path(sys.executable)
            console_python = exe.with_name("python.exe")
            if exe.name.lower() == "pythonw.exe" and console_python.exists():
                exe = console_python
            flags = 0x08000000 if sys.platform == "win32" else 0   # no flashing console window
            output = ""
            ok = False
            for extra in ([], ["--user"]):
                try:
                    result = subprocess.run([str(exe), "-m", "pip", "install", *extra, *packages],
                                            capture_output=True, text=True, creationflags=flags)
                    output = (result.stdout or "") + (result.stderr or "")
                    if result.returncode == 0:
                        ok = True
                        break
                except OSError as exc:
                    output = str(exc)
            self.events.put(("call", lambda: self._install_done(ok, output[-900:], packages)))

        threading.Thread(target=work, daemon=True).start()

    def _install_done(self, ok: bool, output: str, packages: list[str]) -> None:
        self.bar.stop()
        if ok:
            self._append("Helper packages installed.", "done")
            messagebox.showinfo("Installed", "All set. Close Canvas Course Manager and open it again to start "
                                "using them.", parent=self)
        else:
            self._append("ERROR: The helper packages could not be installed.", "error")
            for line in output.splitlines()[-8:]:
                self._append(line, "section")
            messagebox.showerror("Install didn't finish",
                                 "Check your internet connection and try again.\n\nIf it keeps failing, ask whoever "
                                 "shared this program to run this command once on your computer:\n\n"
                                 f"python -m pip install {' '.join(packages)}", parent=self)

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno(
                "A run is still going",
                "Closing now could leave a quiz import or upload half-finished.\n\nClose anyway?",
                icon="warning", parent=self):
            return
        try:
            self._commit_folder()
            self._save_course()
        except (OSError, CanvasError):
            pass
        self.destroy()


def launch() -> None:
    if tk is None:
        _fatal_popup("This copy of Python doesn't include tkinter, which draws the window.\n\n"
                     "Install Python from python.org (it includes tkinter) and try again.\n"
                     "Mac with Homebrew: run  brew install python-tk")
        return
    if sys.platform == "win32":
        try:  # Sharper text on high-resolution Windows screens.
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    try:
        app = App()
    except Exception as exc:
        _fatal_popup(f"Canvas Course Manager couldn't start:\n\n{type(exc).__name__}: {exc}")
        return
    app.mainloop()


if __name__ == "__main__":
    launch()
