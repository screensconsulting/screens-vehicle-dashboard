@echo off
REM ════════════════════════════════════════════════════════════════
REM  Screens Studio — Dashboard Weekly Updater (Windows)
REM  Double-click to run manually, or add to Task Scheduler.
REM ════════════════════════════════════════════════════════════════
REM
REM  TASK SCHEDULER SETUP (one-time):
REM  1. Open Task Scheduler (search "Task Scheduler" in Start)
REM  2. Click "Create Basic Task…"
REM  3. Name: "Screens Dashboard Update"
REM  4. Trigger: Weekly, Monday, 08:00
REM  5. Action: Start a Program
REM     Program: C:\Windows\System32\cmd.exe
REM     Arguments: /c "C:\path\to\this\folder\run_update.bat"
REM  6. Finish — it will run every Monday at 08:00 automatically.
REM
REM  REQUIREMENTS:
REM  - Python 3.8+  (download from python.org)
REM  - pip install openpyxl requests beautifulsoup4
REM ════════════════════════════════════════════════════════════════

setlocal
cd /d "%~dp0"

echo.
echo  ┌─────────────────────────────────────────────────────────┐
echo  │  Screens Studio — Dashboard Updater                     │
echo  │  %DATE% %TIME%
echo  └─────────────────────────────────────────────────────────┘
echo.

REM ── Check Python is available ───────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install from https://python.org
    pause
    exit /b 1
)

REM ── Check required packages ─────────────────────────────────────
python -c "import openpyxl, requests, bs4" >nul 2>&1
if errorlevel 1 (
    echo Installing required Python packages…
    pip install openpyxl requests beautifulsoup4 --quiet
)

REM ── Run the updater ─────────────────────────────────────────────
echo Running update_dashboard.py…
python update_dashboard.py

if errorlevel 1 (
    echo.
    echo ERROR: Update failed. Check update_log.txt for details.
    pause
    exit /b 1
)

echo.
echo  ✓  Dashboard updated successfully!
echo  ✓  Open dashboard.html in your browser to view.
echo.

REM ── Optional: open the dashboard automatically ──────────────────
REM Uncomment the next line to auto-open after update:
REM start "" "%~dp0dashboard.html"

exit /b 0
