@echo off
cd /d "%~dp0"
if not exist ".venv-local\Scripts\python.exe" (
    echo Local Python environment is missing. Run: python -m venv .venv-local
    pause
    exit /b 1
)
echo Open http://127.0.0.1:5000 after the server starts.
echo Keep this window open. Press Ctrl+C to stop the app.
".venv-local\Scripts\python.exe" app.py
if errorlevel 1 pause
