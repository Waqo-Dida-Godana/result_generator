# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[('C:\\Users\\WDG\\AppData\\Local\\Programs\\Python\\Python312\\DLLs\\_tkinter.pyd', '.'), ('C:\\Users\\WDG\\AppData\\Local\\Programs\\Python\\Python312\\DLLs\\tcl86t.dll', '.'), ('C:\\Users\\WDG\\AppData\\Local\\Programs\\Python\\Python312\\DLLs\\tk86t.dll', '.')],
    datas=[('C:\\Users\\WDG\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tcl8.6', 'tcl\\tcl8.6'), ('C:\\Users\\WDG\\AppData\\Local\\Programs\\Python\\Python312\\tcl\\tk8.6', 'tcl\\tk8.6'), ('moas.ico', '.'), ('cbc_school.db', '.'), ('school_report.db', '.')],
    hiddenimports=['tkinter', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyinstaller_tk_runtime_hook.py'],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='MOAS_MIS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['moas.ico'],
)
