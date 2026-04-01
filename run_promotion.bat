@echo off
REM Student Promotion Task Runner for Windows Task Scheduler
REM This batch file can be scheduled to run automatically using Windows Task Scheduler

REM Change to the script directory
cd /d "%~dp0"

REM Set Python path (adjust if Python is not in PATH)
set PYTHON_PATH=python

REM Log file location
set LOG_FILE=promotion_task.log

REM Run the promotion task
echo ======================================== >> %LOG_FILE%
echo Promotion Task Started: %date% %time% >> %LOG_FILE%
echo ======================================== >> %LOG_FILE%

%PYTHON_PATH% run_promotion_task.py --verbose >> %LOG_FILE% 2>&1

echo ======================================== >> %LOG_FILE%
echo Promotion Task Completed: %date% %time% >> %LOG_FILE%
echo ======================================== >> %LOG_FILE%

REM Exit with the Python script's exit code
exit /b %ERRORLEVEL%
