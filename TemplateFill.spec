# -*- mode: python ; coding: utf-8 -*-
# TemplateFill 打包配置（onedir 模式，复刻 ChemCal 打包方式）
# - onedir：不解压到 %TEMP%，启动快、杀软误报少；分发用 Inno Setup 或便携 zip
# - 打包环境用 .venv-build（PySide6-Essentials），产物只含用得到的 Qt DLL

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('version.py', '.'), ('TemplateFill.ico', '.')],
    hiddenimports=['PySide6.QtCore', 'PySide6.QtWidgets', 'PySide6.QtGui',
                   'lxml._elementpath',
                   'doc_filler', 'xlsx_filler', 'html_to_docx',
                   'section_editor', 'batch_manager', 'labels'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'PySide2',
              'scipy', 'pandas', 'numpy', 'matplotlib',
              'tkinter', 'unittest', 'pytest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,   # onedir：二进制交给 COLLECT，不塞进 exe
    name='TemplateFill',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['TemplateFill.ico'],   # exe 文件图标（资源管理器/任务栏）
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TemplateFill',
)
