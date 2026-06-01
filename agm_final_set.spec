# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for AGM_Final_Set.

Bundles app_main.py + scripts/ + sample/Master_Input_DEMO.xlsx into a single
console executable. The Real master/ folder is NOT bundled — only the demo
template ships in the public binary.
"""

block_cipher = None


a = Analysis(
    ['app_main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('scripts', 'scripts'),
        ('sample/Master_Input_DEMO.xlsx', 'sample'),
        ('sample/Master_Input_EMPTY.xlsx', 'sample'),
    ],
    hiddenimports=[
        # python-docx / openpyxl pull these dynamically; declare so PyInstaller
        # doesn't miss them.
        'docx',
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.styles',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AGM_Final_Set',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
