@echo off
cd /d "%~dp0"
if not exist "plan-approval-lead-status.html" (
    echo No local dashboard file yet - run publish.bat first to calculate it.
    pause
    exit /b 1
)
start "" "plan-approval-lead-status.html"
