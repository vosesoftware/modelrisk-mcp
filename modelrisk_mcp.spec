# PyInstaller spec for ModelRisk MCP Server.
# Phase 0 smoke build: produces a single-file Windows .exe that boots the
# (empty) FastMCP server over stdio. Validates that the pywin32 + xlwings +
# mcp + pydantic combination bundles cleanly before we add a real tool
# surface in later phases.

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(
    ["src/modelrisk_mcp/__main__.py"],
    pathex=["src"],
    binaries=[],
    datas=[],
    hiddenimports=[
        "win32com",
        "win32com.client",
        "pywintypes",
        "xlwings",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="modelrisk-mcp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
