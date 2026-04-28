# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [
    ('templates', 'templates'),
    ('BOL INPUT.docx', '.'),
]
datas += collect_data_files('pdfplumber')
datas += collect_data_files('docx')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=collect_submodules('pdfplumber') + collect_submodules('docx'),
    hookspath=[],
    runtime_hooks=[],
    excludes=['venv'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BOL-Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='BOL-Generator',
)
