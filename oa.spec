# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pyhub_office_automation/cli/main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32com.gen_py',
        'win32com.client.gencache',
        'win32com.shell.shell',
        'pywintypes',
        'pythoncom',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['runtime_hook_win32com.py'],
    excludes=['matplotlib', 'scipy', 'sklearn', 'tkinter', 'IPython', 'jupyter'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='oa',
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
    icon='pyhub_office_automation/assets/icons/logo.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='oa',
)
