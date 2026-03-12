@echo off
echo ========================================
echo   MOAS MIS - EXE Builder
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
)

REM Activate virtual environment
call .venv\Scripts\activate.bat

echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building EXE...
echo.

REM Build with PyInstaller - include all data files and run in windowed mode
pyinstaller --onefile --windowed ^
    --name "MOAS_MIS" ^
    --add-data "moas.ico;." ^
    --add-data "cbc_school.db;." ^
    --add-data "school_report.db;." ^
    --icon moas.ico ^
    --clean ^
    main.py

echo.
echo ========================================
echo   Build Complete!
echo ========================================
echo The EXE file is in the 'dist' folder
echo.

REM Open the dist folder
explorer dist

pause
