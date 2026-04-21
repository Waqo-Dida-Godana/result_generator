@echo off
setlocal
cd /d "%~dp0result_generator"

echo Starting MOAS Result Generator...

if exist ".venv\Scripts\python.exe" (
    echo Using virtual environment...
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
