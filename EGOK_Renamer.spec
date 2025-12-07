# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['EGOK_Renamer.py'],
    pathex=[],
    binaries=[],
    datas=[('background.png', '.'), ('settings.json', '.'), ('icon.ico', '.'), ('plugins', 'plugins')],
    hiddenimports=['watchdog.observers', 'watchdog.events', 'PIL', 'PIL._tkinter_finder', 'PIL.Image', 'threading', 'queue', 'pathlib', 're', 'importlib', 'inspect', 'json', 'tksheet', 'sqlite3', 'serial', 'serial.tools.list_ports', 'serial.serialutil', 'serial.win32', 'PyPDF2', 'simplekml', 'binascii', 'math'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name='EGOK_Renamer',
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
    icon=['icon.ico'],
)
