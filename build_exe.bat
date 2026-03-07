@echo off
echo Installing PyInstaller...
pip install pyinstaller
echo Converting CBC Student App to EXE...
pyinstaller --onefile --windowed cbc_student_app.py
echo Done! You can find the EXE in the 'dist' folder.
pause
