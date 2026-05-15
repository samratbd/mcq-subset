# -*- mode: python ; coding: utf-8 -*-
# Onurion OMR Studio — PyInstaller spec
# Run from the PROJECT ROOT:
#   pyinstaller --noconfirm desktop/mcq_studio.spec

from PyInstaller.utils.hooks import collect_submodules
from pathlib import Path

block_cipher = None
# SPECPATH is the folder containing this .spec file (desktop/)
# Project root is one level up
DESKTOP = Path(SPECPATH).resolve()
ROOT = DESKTOP.parent

hidden = []
for pkg in ['app', 'app.omr', 'app.parsers', 'app.writers']:
    hidden += collect_submodules(pkg)
hidden += [
    'cv2', 'numpy', 'PIL._imaging', 'PIL._tkinter_finder',
    'docx', 'docx.oxml', 'docx.oxml.parser', 'docx.oxml.ns',
    'lxml', 'lxml.etree', 'lxml._elementpath',
    'openpyxl', 'openpyxl.styles',
]

# Bundle everything from the project root that the desktop app needs
datas = []
def bundle_dir(src, dest):
    for p in src.rglob('*'):
        if p.is_file() and '__pycache__' not in str(p) and '.pyc' not in str(p):
            rel = str(p.parent.relative_to(src))
            target = dest if rel == '.' else f'{dest}/{rel}'
            datas.append((str(p), target))

bundle_dir(ROOT / 'app',     'app')       # shared engine
bundle_dir(ROOT / 'desktop' / 'static', 'static')   # header image etc

a = Analysis(
    [str(DESKTOP / 'mcq_studio.py')],
    pathex=[str(ROOT), str(DESKTOP)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'pytest', 'IPython', 'jupyter'],
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
    name='Onurion_OMR_Studio',
    debug=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    bootloader_ignore_signals=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(DESKTOP / 'installer' / 'mcq_studio.ico')
         if (DESKTOP / 'installer' / 'mcq_studio.ico').exists() else None,
)
