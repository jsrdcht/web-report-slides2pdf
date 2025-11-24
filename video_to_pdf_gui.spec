# PyInstaller spec for GUI frontend

block_cipher = None

from PyInstaller.utils.hooks import collect_submodules, collect_all

hiddenimports = ['tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'tkinter.ttk']
# Be conservative; cv2/PIL hooks generally cover these, but allow extension when needed
hiddenimports += collect_submodules('cv2')

# Collect yt_dlp resources to ensure it works in frozen app
yd_datas, yd_binaries, yd_hiddenimports = collect_all('yt_dlp')
hiddenimports += yd_hiddenimports

a = Analysis(
    ['video_to_pdf_gui.py'],
    pathex=['.'],
    binaries=yd_binaries,
    datas=yd_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='video2pdf-gui',
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
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='video2pdf-gui'
)

import sys
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='video2pdf-gui.app',
        icon=None,
        bundle_identifier='com.github.video2pdf.gui'
    )


