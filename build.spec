# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Lender Feedback Tool - produces a single-file executable."""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all pdfplumber and pdfminer data files (fonts, etc.)
datas = []
datas += collect_data_files('pdfplumber')
datas += collect_data_files('pdfminer')
datas += collect_data_files('docx')

# Hidden imports - modules that PyInstaller might miss
hiddenimports = []
hiddenimports += collect_submodules('pdfplumber')
hiddenimports += collect_submodules('pdfminer')
hiddenimports += collect_submodules('docx')
hiddenimports += [
    'flask',
    'werkzeug',
    'jinja2',
    'itsdangerous',
    'click',
    'PIL',
    'lxml',
    'lxml._elementpath',
    'pypdfium2',
    'pypdfium2_raw',
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'numpy.random._examples',
        'tkinter', 'test', 'unittest', 'pydoc_data',
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
    name='LenderFeedbackTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # Show console so user sees "Starting server..." messages
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
