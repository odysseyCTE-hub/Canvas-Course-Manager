#!/bin/bash
# Mac launcher: double-click to open Canvas Course Manager.
# Keep this file in the same folder as canvas_course_manager.py.
cd "$(dirname "$0")" || exit 1
APP="canvas_course_manager.py"
finish() { echo; read -n 1 -s -r -p "Press any key to close this window..."; echo; exit 1; }

[ -f "$APP" ] || { echo "Can't find $APP next to this launcher."; finish; }

if ! command -v python3 >/dev/null 2>&1 || \
   ! python3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" 2>/dev/null; then
  echo "Canvas Course Manager needs Python 3.10 or newer from python.org."
  echo "A download page will open. Install it, then double-click this launcher again."
  open "https://www.python.org/downloads/"
  finish
fi

python3 -c "import tkinter" 2>/dev/null || {
  echo "This Python can't draw windows (tkinter missing). Install Python from python.org instead."
  finish; }

if ! python3 -c "import pandas, openpyxl, sv_ttk" 2>/dev/null; then
  echo "First-time setup: installing helper packages (one time only)..."
  python3 -m pip install --user --disable-pip-version-check pandas openpyxl tzdata sv-ttk || {
    echo "The helper packages didn't install. Check your internet and try again."; finish; }
fi

echo "Opening Canvas Course Manager. Leave this Terminal window open while you use it."
exec python3 "$APP"
