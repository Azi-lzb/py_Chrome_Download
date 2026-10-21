# -*- mode: python ; coding: utf-8 -*-
import importlib

from PyInstaller.utils.hooks import collect_submodules


required_modules = ['selenium', 'openpyxl']
missing_modules = [name for name in required_modules if importlib.util.find_spec(name) is None]
if missing_modules:
    raise SystemExit(
        'Missing build dependency: '
        + ', '.join(missing_modules)
        + '. Build with the HZPBC environment\'s Python or install the missing packages.'
    )

# 自动收集 selenium 全部子模块，避免动态 import 漏掉。
# If selenium is absent from the build Python, collect_submodules() returns an
# empty list and the exe later fails with "No module named selenium", so fail
# the build here instead of producing a broken artifact.
selenium_imports = collect_submodules('selenium')
if not selenium_imports:
    raise SystemExit('No selenium submodules were collected; aborting broken build.')

base_hidden_imports = [
    'openpyxl',
    'openpyxl.comments',
    'urllib3',
    'certifi',
]
optional_hidden_imports = [
    name for name in ['charset_normalizer']
    if importlib.util.find_spec(name) is not None
]
hidden_imports = sorted(set(base_hidden_imports + optional_hidden_imports + selenium_imports))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'PIL',
        'tkinter',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
        'curses',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'setuptools',
        'unittest',
        'pydoc',
        'doctest',
        'argparse',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ChromeAutoRun',
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
