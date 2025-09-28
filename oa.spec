# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import copy_metadata

datas = [('pyhub_office_automation\\resources', 'pyhub_office_automation\\resources'), ('README.md', '.')]
datas += collect_data_files('fastmcp')
datas += collect_data_files('uvicorn')
datas += copy_metadata('fastmcp')
datas += copy_metadata('uvicorn')


a = Analysis(
    ['pyhub_office_automation\\cli\\main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'keyring',
        'keyring.backends',
        'keyring.backends.Windows',
        'keyring.backends._OS_X_API',
        'keyring.backends.SecretService',
        'win32cred',
        'win32con',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'sklearn', 'tkinter', 'IPython', 'jupyter', 'PIL.ImageQt'],
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
    name='oa',
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
    icon=['pyhub_office_automation\\assets\\icons\\logo.ico'],
)
