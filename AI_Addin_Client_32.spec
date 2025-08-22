# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['wps_addin\\addin_client.py'],
    pathex=[],
    binaries=[],
    datas=[('.\\wps_addin\\ribbon.xml', '.'), ('.\\wps_addin\\ai_icon.png', '.'), ('.\\wps_addin\\memo_icon.png', '.'), ('.\\wps_addin\\minutes_icon.png', '.'), ('.\\wps_addin\\cover_letter_icon.png', '.'), ('.\\wps_addin\\analyze_icon.png', '.'), ('.\\wps_addin\\summarize_icon.png', '.')],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='AI_Addin_Client_32',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
