# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

root = Path(SPECPATH)

a = Analysis(
    [str(root / 'app.py')],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / 'assets' / 'app.ico'), 'assets')],
    hiddenimports=['gspread', 'google.auth', 'google.oauth2.service_account', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(root / 'packaging' / 'pyi_rth_skip_numpy.py')],
    excludes=['numpy', 'pandas', 'matplotlib', 'scipy', 'tkinter'],
    noarchive=False,
    optimize=0,
)
# Qt 6.11 uses the Windows system ICU shim. A third-party ICU on the build PATH
# can be auto-collected and shadow the system DLL, causing WinError 127.
a.binaries = [
    item for item in a.binaries
    if Path(item[0]).name.casefold() not in {'icuuc.dll', 'icudt78.dll'}
]
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    name='SheetDataHub',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(root / 'assets' / 'app.ico')],
    exclude_binaries=True,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='SheetDataHub',
)
if sys.platform == 'darwin':
    icon_png = root / 'assets' / 'logo.png'
    app = BUNDLE(
        coll,
        name='SheetDataHub.app',
        icon=str(icon_png) if icon_png.exists() else str(root / 'assets' / 'app.ico'),
        bundle_identifier='app.sheetdatahub.desktop',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleDisplayName': '表数通',
            'CFBundleName': 'SheetDataHub',
            'CFBundleShortVersionString': '1.4.1',
            'LSMinimumSystemVersion': '12.0',
        },
    )
