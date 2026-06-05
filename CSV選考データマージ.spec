# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['openpyxl']
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# exe へ同梱する参照ファイル（配布した単一exeに内蔵し「最初から読み込まれた状態」にする）。
# 第2要素 '.' は展開先(sys._MEIPASS)直下に置く指定で、resource_path() がここを参照する。
datas += [
    ('templates/社員一覧.csv', '.'),
    ('templates/所属一覧.xlsx', '.'),
    ('templates/本部一覧.xlsx', '.'),
    ('templates/個別送信テンプレート_AOL.xlsx', '.'),
]


a = Analysis(
    ['csv_merger.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'torchaudio', 'matplotlib', 'scipy', 'sklearn', 'IPython', 'jupyter', 'notebook', 'cv2', 'PIL'],
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
    name='CSV選考データマージ',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
