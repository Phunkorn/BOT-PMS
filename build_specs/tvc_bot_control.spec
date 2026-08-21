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

# tvc_control intentionally lazy-loads these modules after the Tk root exists.
# They must therefore be explicit for PyInstaller's static analysis.
hiddenimports = ["excel_io", "openpyxl", "psutil", "pywinauto", "tvc_probe"]

a = Analysis(
    [str(SRC_DIR / "tvc_control.py")],
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
    name="TVC Bot Control",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
    name="TVC Bot Control",
)
