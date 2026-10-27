# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import PyQt5
from PyInstaller.utils.hooks.qt import pyqt5_library_info

# Qt5 wheels can report a lossy PrefixPath when installed under a Chinese path.
# Correct the build hook's source paths using Python's Unicode package path.
qt_root = Path(PyQt5.__file__).resolve().parent / 'Qt5'
if ((qt_root / 'plugins').is_dir()
        and not Path(pyqt5_library_info.location['PluginsPath']).is_dir()):
    old_prefix = pyqt5_library_info.location['PrefixPath'].rstrip('/\\')
    for name, location in list(pyqt5_library_info.location.items()):
        if location == old_prefix or location.startswith(old_prefix + '/'):
            pyqt5_library_info.location[name] = qt_root.as_posix() + location[len(old_prefix):]
    pyqt5_library_info.qt_inside_package = True
    pyqt5_library_info.qt_lib_dir = qt_root / 'bin'

a = Analysis(
    ['thermal_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/底层模板.xlsx', 'assets'),
           ('assets/附表格式参考.xlsx', 'assets')],
    hiddenimports=['pythoncom', 'pywintypes', 'win32com.client'],
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
    name='温升报告自动生成器_v4.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
