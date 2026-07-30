# -*- mode: python ; coding: utf-8 -*-
# BOHO_IMPORT_BOM_THDM.spec — PyInstaller 6.x, onedir mode
# Chạy: pyinstaller BOHO_IMPORT_BOM_THDM.spec --clean
# hoặc: python build.py
from PyInstaller.utils.hooks import collect_all

# ── Tài nguyên bundle (CHỈ ĐỌC, truy cập qua resource_path / sys._MEIPASS) ──
datas = [
    ('icon.ico', '.'),      # → _internal/icon.ico
]
binaries = []
hiddenimports = [
    # Database
    'pyodbc',
    # Excel / data
    'pandas',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.nattype',
    'openpyxl',
    'openpyxl.cell._writer',
    'msoffcrypto',
    # Fuzzy matching
    'rapidfuzz',
    'rapidfuzz.fuzz',
    'rapidfuzz.utils',
    # UI
    'tksheet',
    'PIL', 'PIL.Image', 'PIL.ImageTk',
    'tkinter',
    'tkinter.ttk',
    # Stdlib
    'json',
    'concurrent.futures',
    'threading',
    'math',
    'unicodedata',
    'datetime',
]

# customtkinter: gồm theme JSON + CTkImage + nội bộ
tmp_ret = collect_all('customtkinter')
datas         += tmp_ret[0]
binaries      += tmp_ret[1]
hiddenimports += tmp_ret[2]

# tksheet: liệt kê tường minh từng submodule để PyInstaller compile vào PYZ
hiddenimports += [
    'tksheet',
    'tksheet.sheet',
    'tksheet.colors',
    'tksheet.column_headers',
    'tksheet.constants',
    'tksheet.find_window',
    'tksheet.formatters',
    'tksheet.functions',
    'tksheet.main_table',
    'tksheet.menus',
    'tksheet.other_classes',
    'tksheet.row_index',
    'tksheet.sheet_options',
    'tksheet.sorting',
    'tksheet.text_editor',
    'tksheet.themes',
    'tksheet.tksheet_types',
    'tksheet.tooltip',
    'tksheet.top_left_rectangle',
]

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'numpy.distutils',
        'unittest', 'doctest',
        'email', 'http', 'xmlrpc', 'ftplib', 'imaplib', 'poplib',
        'pydoc', 'nntplib', 'telnetlib',
        'IPython', 'notebook', 'jupyter',
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

# ── EXE (onedir: không nhúng binaries/datas vào .exe) ────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    name='BOHO_IMPORT_BOM_THDM',
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
    icon='icon.ico',
)

# ── COLLECT: gom tất cả vào dist/BOHO_IMPORT_BOM_THDM/ ───────────────────────
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BOHO_IMPORT_BOM_THDM',
)
