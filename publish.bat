@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   Plan Approval Dashboard - Update and Publish
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on PATH. Install Python and try again.
    pause
    exit /b 1
)
where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git was not found on PATH. Install Git for Windows and try again.
    pause
    exit /b 1
)

echo [1/3] Calculating latest dashboard data from the Excel files in this folder...
echo.
python "tools\build_dashboard.py"
if errorlevel 1 (
    echo.
    echo Build failed - see the error above. Nothing was published.
    pause
    exit /b 1
)

git remote get-url origin >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: no GitHub remote is configured yet in this folder.
    echo This is a one-time setup step - ask for help wiring it up.
    pause
    exit /b 1
)

echo.
echo [2/3] Committing changes...
git add index.html tools .gitignore publish.bat open-dashboard.bat
git diff --cached --quiet
if not errorlevel 1 (
    echo No changes since the last update - the published dashboard is already current.
    pause
    exit /b 0
)
git commit -m "Update dashboard data - %date% %time%"
if errorlevel 1 (
    echo Commit failed - see the error above.
    pause
    exit /b 1
)

echo.
echo [3/3] Pushing to GitHub...
echo (If this is the first time on this PC, a browser window may open asking you to sign in to GitHub.)
git push
if errorlevel 1 (
    echo.
    echo Push failed - see the error above.
    pause
    exit /b 1
)

echo.
echo Done. The published dashboard is now up to date.
pause
