@echo off
setlocal
title Canvas Course Manager - Starting
rem Double-click this file to open Canvas Course Manager.
rem It must sit in the same folder as canvas_course_manager.py.
cd /d "%~dp0"
set "APP=canvas_course_manager.py"
set "PKGS=pandas openpyxl tzdata sv-ttk"

if not exist "%APP%" goto :no_app

rem --- Find Python: try the "py" launcher first, then "python" ---
set "PY="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (set "PY=py -3" & goto :have_python)
python -c "import sys" >nul 2>&1
if not errorlevel 1 (set "PY=python" & goto :have_python)
goto :no_python

:have_python
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 goto :old_python

%PY% -c "import tkinter" >nul 2>&1
if errorlevel 1 goto :no_tk

rem --- Install helper packages the first time only ---
%PY% -c "import pandas, openpyxl, tzdata, sv_ttk" >nul 2>&1
if not errorlevel 1 goto :launch
echo.
echo  First-time setup: installing a few free helper packages.
echo  This only happens once and can take a minute or two...
echo.
%PY% -m pip install --user --disable-pip-version-check --no-warn-script-location %PKGS%
if errorlevel 1 goto :pip_failed

:launch
rem Open the window without leaving a black console window behind.
%PY% -c "import os,sys,subprocess; w=os.path.join(os.path.dirname(sys.executable),'pythonw.exe'); exe=w if os.path.exists(w) else sys.executable; subprocess.Popen([exe, r'%APP%'], creationflags=0x00000008)"
if errorlevel 1 goto :launch_failed
exit /b 0

:no_app
echo.
echo  Can't find %APP% in this folder:
echo    %~dp0
echo  Keep this launcher in the same folder as the program and try again.
goto :done

:no_python
echo.
echo  Python isn't installed on this computer yet (or Windows can't find it).
echo.
echo  1. A download page will open now. Click the big yellow "Download Python" button.
echo  2. Run the installer. On the FIRST screen, tick the box that says
echo     "Add python.exe to PATH", then click "Install Now".
echo  3. When it finishes, double-click this launcher again.
start "" "https://www.python.org/downloads/"
goto :done

:old_python
echo.
echo  This computer has an older Python. Canvas Course Manager needs 3.10 or newer.
echo  A download page will open now - install the newest version, then try again.
start "" "https://www.python.org/downloads/"
goto :done

:no_tk
echo.
echo  Python is installed, but without the part that draws windows (tkinter).
echo  Run the Python installer again, choose "Modify", make sure
echo  "tcl/tk and IDLE" is ticked, then try this launcher again.
goto :done

:pip_failed
echo.
echo  The helper packages didn't install. Common causes: no internet,
echo  or a school network that blocks downloads.
echo  Try again on another network, or send the messages above to Lisa.
goto :done

:launch_failed
echo.
echo  Python started but couldn't open the program. Send the messages above to Lisa.
goto :done

:done
echo.
pause
exit /b 1
