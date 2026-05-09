@echo off
setlocal
cd /d "%~dp0"

echo Starting MOAS Result Generator...

if exist ".venv314\Scripts\python.exe" (
    echo Using Python 3.14 virtual environment...
    ".venv314\Scripts\python.exe" main.py
) else if exist ".venv\Scripts\python.exe" (
    echo Using default virtual environment...
    ".venv\Scripts\python.exe" main.py
) else (
    echo Using system python...
    python main.py
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo Application exited with error code %ERRORLEVEL%
    pause
)
endlocal
