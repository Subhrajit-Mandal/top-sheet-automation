@echo off
setlocal
cd /d "%~dp0"

rem Double-click launcher for Windows. Finds Python, installs the required
rem packages on the first run, starts the program and keeps the window open.

set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo Python was not found. Install Python 3.10 or later from python.org, then run this again.
  echo.
  pause
  exit /b 1
)

%PY% -c "import docx, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing the required packages. This is only needed the first time...
  %PY% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo The packages could not be installed. See the message above.
    pause
    exit /b 1
  )
)

%PY% generate_top_sheets.py %*
echo.
pause
