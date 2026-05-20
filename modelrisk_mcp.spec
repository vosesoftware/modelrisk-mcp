# PyInstaller spec for ModelRisk MCP Server.
#
# Produces a single-file Windows .exe that boots the FastMCP server
# over stdio. The .exe ships:
# - every tool/resource/prompt module (declared as hidden imports so
#   PyInstaller doesn't strip them — they're only used via dynamic
#   side-effect imports in `server.py`)
# - the four data files (catalogue + override + audit + distribution
#   guide) bundled into the executable
#
# Build locally with:
#     uv run pyinstaller modelrisk_mcp.spec --clean

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Every module Claude Code's dynamic registrations touch — listed
# explicitly because PyInstaller's static analyser doesn't always
# follow the side-effect-import pattern in server.py.
_TOOL_MODULES = [
    "modelrisk_mcp.tools",
    "modelrisk_mcp.tools.reading",
    "modelrisk_mcp.tools.building",
    "modelrisk_mcp.tools.simulation",
    "modelrisk_mcp.tools.workflows",
    "modelrisk_mcp.tools.restore",
]
_RESOURCE_MODULES = [
    "modelrisk_mcp.resources",
    "modelrisk_mcp.resources.function_reference",
    "modelrisk_mcp.resources.distribution_guide",
    "modelrisk_mcp.resources.methodology",
    "modelrisk_mcp.resources.workbook_state",
    "modelrisk_mcp.resources.audit_rules",
]
_PROMPT_MODULES = [
    "modelrisk_mcp.prompts",
    "modelrisk_mcp.prompts.build_model",
    "modelrisk_mcp.prompts.audit_model",
    "modelrisk_mcp.prompts.interpret_results",
    "modelrisk_mcp.prompts.add_uncertainty",
    "modelrisk_mcp.prompts.import_legacy_model",
]
_BRIDGE_AND_AUDIT_MODULES = [
    "modelrisk_mcp.audit",
    "modelrisk_mcp.audit.engine",
    "modelrisk_mcp.audit.rules",
    "modelrisk_mcp.bridge",
    "modelrisk_mcp.bridge.catalogue",
    "modelrisk_mcp.bridge.excel",
    "modelrisk_mcp.bridge.formulas",
    "modelrisk_mcp.bridge.modelrisk",
    "modelrisk_mcp.bridge.progids",
    "modelrisk_mcp.bridge.results",
    "modelrisk_mcp.bridge.simulation",
]
_COM_MODULES = [
    "win32com",
    "win32com.client",
    "pywintypes",
    "win32api",
    "win32event",
    "xlwings",
]


a = Analysis(
    ["src/modelrisk_mcp/__main__.py"],
    pathex=["src"],
    binaries=[],
    datas=[
        ("src/modelrisk_mcp/data/functions.json", "modelrisk_mcp/data"),
        ("src/modelrisk_mcp/data/optional_overrides.yaml", "modelrisk_mcp/data"),
        ("src/modelrisk_mcp/data/audit_rules.yaml", "modelrisk_mcp/data"),
        ("src/modelrisk_mcp/data/distributions.yaml", "modelrisk_mcp/data"),
    ],
    hiddenimports=(
        _TOOL_MODULES
        + _RESOURCE_MODULES
        + _PROMPT_MODULES
        + _BRIDGE_AND_AUDIT_MODULES
        + _COM_MODULES
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "matplotlib",
        "PySide6",
        "PyQt6",
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
