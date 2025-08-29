# PyInstaller spec file for 32-bit WPS Add-in
# Usage: pyinstaller build_32bit.spec

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['run_32bit.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('ribbon.xml', '.'),
        ('*.png', '.'),  # Include any icon files
    ],
    hiddenimports=[
        'win32com.client',
        'win32com.server.register',
        'win32com.server.localserver',
        'pythoncom',
        'win32api',
        'winreg',
        'requests',
        'tkinter',
        'wps_addin_base',
        'wps_addin_32bit',
        'wps_registry_utils'
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
    name='WPSAddin_32bit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='x86',  # Force 32-bit
    codesign_identity=None,
    entitlements_file=None,
)