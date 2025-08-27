# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['wps_addin\\addin_client.py'],
    pathex=[],
    binaries=[],
    datas=[('.\\wps_addin\\ribbon.xml', 'wps_addin'), ('.\\wps_addin\\*.png', 'wps_addin')],
    hiddenimports=['pythoncom', 'win32com.client', 'win32com.server.register'],
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
    name='AI_Addin_Client_64',
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
    uac_admin=True,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AI_Addin_Client_64',
)
