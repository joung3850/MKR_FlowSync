# PyInstaller one-file, windowed build for the MKR SYNC HTML desktop app.
from pathlib import Path

project_root = Path(SPECPATH)

analysis = Analysis(
    [str(project_root / "mkr_sync_gui.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / "mkr_sync" / "web" / "index.html"), "mkr_sync/web"),
    ],
    hiddenimports=[
        "pythoncom",
        "pywintypes",
        "win32timezone",
        "win32com",
        "win32com.client",
        "win32com.server",
        "webview",
        "webview.platforms.edgechromium",
        "pypdf",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "cefpython3",
        "webview.platforms.qt",
        "webview.platforms.cef",
        "webview.platforms.gtk",
    ],
    noarchive=False,
)

pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="MKR_SYNC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
