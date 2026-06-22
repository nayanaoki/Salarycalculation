# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller ビルド定義（onefile / GUI）。

ビルド:  pyinstaller --noconfirm payroll.spec
出力  :  dist\給与自動計算.exe
"""
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# reportlab(CIDフォントデータ)・openpyxl・certifi(CA証明書) を確実に同梱する
for pkg in ("reportlab", "openpyxl", "certifi"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ['payroll_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='給与自動計算',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # GUI アプリ(コンソール非表示)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app.ico',       # アイコンを使う場合は app.ico を置いて有効化
)
