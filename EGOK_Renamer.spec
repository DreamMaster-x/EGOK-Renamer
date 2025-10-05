# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('background.png', '.'), ('settings.json', '.'), ('icon.ico', '.'), ('plugins\\*', 'plugins')]
binaries = []
hiddenimports = ['watchdog.observers', 'watchdog.events', 'PIL', 'PIL._tkinter_finder', 'threading', 'queue', 'pathlib', 're', 'importlib', 'inspect', 'importlib.util', 'importlib.machinery', 'requests', 'json', 'tksheet', 'tksheet._tksheet', 'tksheet._tksheet_formatters', 'tksheet._tksheet_other', 'tksheet._tksheet_main_table', 'tksheet._tksheet_top_left_rectangle', 'tksheet._tksheet_row_index', 'tksheet._tksheet_header', 'tksheet._tksheet_column_drag_and_drop']
tmp_ret = collect_all('plugins')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tksheet')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
