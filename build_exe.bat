@echo off
echo ========================================
echo   MOAS MIS - EXE Builder
echo ========================================
echo.

set "PYTHON_EXE=C:\Users\WDG\AppData\Local\Programs\Python\Python312\python.exe"
set "PYTHON_HOME=C:\Users\WDG\AppData\Local\Programs\Python\Python312"

if not exist "%PYTHON_EXE%" (
    echo Python 3.12 was not found at:
    echo %PYTHON_EXE%
    echo.
    echo Install Python 3.12 with Tcl/Tk enabled, then run this script again.
    pause
    exit /b 1
)

echo Installing dependencies...
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install -r requirements.txt
"%PYTHON_EXE%" -m pip install pyinstaller

echo.
echo Building EXE...
echo.

REM Build with PyInstaller - include all data files and run in windowed mode
"%PYTHON_EXE%" -m PyInstaller --onefile --windowed ^
    --name "MOAS_MIS" ^
    --hidden-import tkinter ^
    --hidden-import _tkinter ^
    --add-binary "%PYTHON_HOME%\DLLs\_tkinter.pyd;." ^
    --add-binary "%PYTHON_HOME%\DLLs\tcl86t.dll;." ^
    --add-binary "%PYTHON_HOME%\DLLs\tk86t.dll;." ^
    --add-data "%PYTHON_HOME%\tcl\tcl8.6;tcl\tcl8.6" ^
    --add-data "%PYTHON_HOME%\tcl\tk8.6;tcl\tk8.6" ^
    --add-data "moas.ico;." ^
    --add-data "cbc_school.db;." ^
    --add-data "school_report.db;." ^
    --runtime-hook pyinstaller_tk_runtime_hook.py ^
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
