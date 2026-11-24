# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置 —— 梁溪区工单分析工具
# 打包目录: app/   (在 app 目录下执行 pyinstaller)
import os

# 用 SPECPATH 取 spec 所在目录，保证 datas 路径无论从哪执行都能解析
_APP_DIR = os.path.join(SPECPATH, '')

a = Analysis(
    ['main.py'],
    pathex=[_APP_DIR],
    binaries=[],
    datas=[
        (os.path.join(_APP_DIR, 'ui/index.html'), 'ui'),
        (os.path.join(_APP_DIR, 'ui/tokens.css'), 'ui'),
        (os.path.join(_APP_DIR, 'ui/echarts.min.js'), 'ui'),
    ],
    hiddenimports=[
        # Windows GUI 后端（pywebview）
        'webview.platforms.edgechromium',
        'webview.platforms.winforms',
        'webview.platforms.mshtml',
        # pythonnet / .NET 桥（edgechromium 依赖，动态加载需显式声明）
        'pythonnet',
        'clr_loader',
        'clr_loader.netfx',
        'clr_loader.coreclr',
        'clr_loader.mono',
        'cffi',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='梁溪区工单分析工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,        # 纯 GUI，不弹黑窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
