# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for markitdown.exe (one-file, console + built-in GUI).
#
#   pyinstaller --noconfirm --clean packaging/markitdown.spec
#
# Output: dist/markitdown.exe (dist/markitdown on Linux/macOS).
import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
SKILL_SCRIPTS = os.path.join(ROOT, "skills", "markitdown", "scripts")

datas = []
# Magika (file-type detection used by MarkItDown) loads its ONNX model and
# config from files next to its package.
datas += collect_data_files("magika")
# Distribution metadata some packages read through importlib.metadata.
for dist in ("markitdown", "magika"):
    datas += copy_metadata(dist)
# GUI: Sun Valley (Windows 11) theme files and the app icon.
datas += collect_data_files("sv_ttk")
datas += [(os.path.join(ROOT, "app", "assets"), "assets")]

hiddenimports = []
# Converters are imported dynamically inside try/except blocks.
hiddenimports += collect_submodules("markitdown")
hiddenimports += ["convert", "gui", "sv_ttk"]

a = Analysis(
    [os.path.join(ROOT, "app", "markitdown_app.py")],
    pathex=[SKILL_SCRIPTS],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "IPython", "matplotlib", "notebook", "jupyter"],
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
    name="markitdown",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ROOT, "app", "assets", "icon.ico"),
)
