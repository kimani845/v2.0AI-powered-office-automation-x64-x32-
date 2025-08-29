# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['wps_addin\\backend_server.py'],
    pathex=[],
    binaries=[],
    datas=[('.\\.env', '.'), ('.\\app', 'app')],
    hiddenimports=['uvicorn.logging', 'uvicorn.loops', 'uvicorn.protocols', 'uvicorn.lifespan'],
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
    name='AI_Backend_Server',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AI_Backend_Server',
)
