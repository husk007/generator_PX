# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

a = Analysis(
    ['generator_PX.py'],
    pathex=[],
    binaries=[],
    # czcionka używana do podpisów kodów kreskowych jest dołączona do pakietu python-barcode
    datas=collect_data_files('barcode'),
    hiddenimports=[
        'pyperclip',
        'barcode',
        'barcode.writer',
        'PIL',
        'PIL._imaging',
        'reportlab',
        'reportlab.graphics.barcode',
        'reportlab.graphics.barcode.code128',
        'pygetwindow',
        'pyautogui',
        'win32api',
        'win32con',
        'win32gui',
        'pymsgbox',
        'pytweening',
        'PyTweening',
        'pywintypes',
        'uuid',
        'pyrect',
        'pyscreeze',
        'pyscreeze._screenshot',
    ],
    hookspath=['hooks'],
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
    [],
    exclude_binaries=True,
    name='generator_PX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='barcode.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='generator_PX',
)
