# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
import os

datas = []
datas += collect_data_files('customtkinter')
datas += collect_data_files('tkinterdnd2')
datas += collect_data_files('fitz')          # pymupdf fonts/resources
datas += collect_data_files('pyzbar')        # zbar-LICENSE etc.

# pyzbar DLLs (libzbar-64.dll, libiconv.dll) — Windows only
binaries = []
binaries += collect_dynamic_libs('pyzbar')

a = Analysis(
    ['invoice_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'customtkinter',
        'tkinterdnd2',
        'fitz',
        'pyzbar',
        'pyzbar.pyzbar',
        'PIL._tkinter_finder',
    ],
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
    name='InvoiceApp',
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
    icon=None,
)
