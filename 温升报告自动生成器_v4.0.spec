# -*- mode: python ; coding: utf-8 -*-

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
