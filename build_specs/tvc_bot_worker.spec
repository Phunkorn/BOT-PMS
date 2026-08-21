from pathlib import Path

PROJECT_ROOT = Path(SPECPATH).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
ASSETS_DIR = PROJECT_ROOT / "assets"
APP_ICON = ASSETS_DIR / "app_icon.ico"

datas = [
    (str(PROJECT_ROOT / "config.ini"), "."),
    (str(PROJECT_ROOT / "field_map.json"), "."),
    (str(ASSETS_DIR), "assets"),
]

# bot.py imports openpyxl/pywinauto through its normal module graph. psutil is
# included explicitly because process-tree handling is part of the distribution
# contract even though it is driven primarily by the GUI.
hiddenimports = ["psutil", "tvc_probe", "tvc_window_locator"]

a = Analysis(
    [str(SRC_DIR / "bot_worker.py")],
    pathex=[str(SRC_DIR), str(PROJECT_ROOT)],
    binaries=[],
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
    [],
    exclude_binaries=True,
    name="TVC Bot Worker",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(APP_ICON) if APP_ICON.is_file() else None,
    contents_directory="_internal",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="TVC Bot Worker",
)
