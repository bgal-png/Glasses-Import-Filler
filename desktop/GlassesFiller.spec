# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — single-file, windowed .exe.

Build from the SHORT-PATH venv (the Store Python's path is long enough that
installing PySide6 fails with "enable-long-paths"):

    "C:\\gv\\Scripts\\pyinstaller.exe" desktop\\GlassesFiller.spec

Output: dist\\GlassesFiller.exe — needs nothing installed on the target machine.
It is large and unsigned, so SmartScreen warns once ("More info" → "Run anyway").
"""
import os

# Run from the repo root so the shared logic modules are importable.
ROOT = os.path.abspath(os.getcwd())
DESKTOP = os.path.join(ROOT, "desktop")

a = Analysis(
    [os.path.join(DESKTOP, "main.py")],
    pathex=[DESKTOP, ROOT],   # desktop modules + repo-root shared logic
    binaries=[],
    datas=[(os.path.join(DESKTOP, "app_icon.ico"), ".")],
    # No template workbook is bundled: the filler writes onto a copy of the
    # user's own workbook, so its styling is preserved.
    hiddenimports=[
        # Imported lazily inside functions, so PyInstaller can't see them.
        "sqlalchemy.dialects.postgresql",
        "psycopg2",
        "anthropic",
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "sklearn", "scipy", "streamlit", "pyarrow"],
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
    name="GlassesFiller",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,            # windowed app, no console flash
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(DESKTOP, "app_icon.ico"),
)
