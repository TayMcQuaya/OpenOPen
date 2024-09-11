# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],  # Ensure current directory is set
    binaries=[],
    datas=[
        ('icons', 'icons'),  # Copy the icons folder
        ('styles', 'styles'),  # Copy the styles folder
        ('resources', 'resources'),  # Copy the resources folder
        ('app_icon_multi.ico', '.'),  # Copy the icon to the root of the output folder
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts + [('main.py', 'C:/Users/tayfu/Desktop/BACKUP 2024-08/OpenOPen - Python/main.py', 'PYSOURCE')],
    exclude_binaries=True,
    name='OpenOPen',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True if you want a console window for debugging
    icon='app_icon_multi.ico'  # Specify the icon file for the executable
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,  # Ensure the `datas` list is included in COLLECT
    strip=False,
    upx=True,
    upx_exclude=[],
    name='OpenOPen'
)
