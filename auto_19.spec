# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files


a = Analysis(
    ['auto_19.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('logs', 'logs'),
        ('credenciais.env', '.'),
    ],
    hiddenimports=[
        'selenium',
        'tkinter',
        'ttkthemes',
        'dotenv',
        'pathlib',
        'requests',
        'logging',
        'threading',
        'datetime',
        'PIL',
        'PIL._tkinter_finder',
        'PIL._imaging',
        'PIL._imagingft'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'PyQt5',
        'PySide2',
        'wx'
    ],
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExtratorSSW',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # Mudado para False para evitar o erro do strip
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)
