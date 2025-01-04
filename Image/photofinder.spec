# -*- mode: python ; coding: utf-8 -*-

import cv2  # Import cv2 to get the file path dynamically
block_cipher = None

a = Analysis(
    ['photofinder.py'],  # Your main script
    pathex=['.'],  # Path to the script
    binaries=[(cv2.__file__, '.')],  # Embed OpenCV dynamically
    datas=[],
    hiddenimports=['cv2'],  # Ensure cv2 is included
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='photofinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='photofinder',
)