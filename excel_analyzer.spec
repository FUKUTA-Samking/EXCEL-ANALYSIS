# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包設定檔
# 用法（在 Windows 上，安裝好 requirements.txt + pyinstaller 之後）：
#     pyinstaller excel_analyzer.spec
# 打包完成後執行檔在 dist\ExcelAnalyzer\ExcelAnalyzer.exe（或 dist\ExcelAnalyzer.exe，視 ONEFILE 設定而定）

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # 以下是曾經在 Windows 環境下用到 matplotlib/openpyxl/PIL 時，
    # PyInstaller 靜態分析可能漏抓的模組，明確列出以策安全
    hiddenimports=[
        'PIL._tkinter_finder',
        'openpyxl.cell._writer',
        'matplotlib.backends.backend_tkagg',
        'scipy.special._cdflib',
    ],
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
    name='ExcelAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # False = 不跳出黑色命令視窗（GUI程式）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app_icon.ico',  # 如果有自己的圖示檔，取消註解並填入路徑
)
