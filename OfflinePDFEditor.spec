# -*- mode: python ; coding: utf-8 -*-
# One PDF Editor – single-file .exe

block_cipher = None

a = Analysis(
    ['app/main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/OnePDFEditor.ico', 'assets'),
        ('assets/icon_256.png', 'assets'),
        ('assets/icon_64.png', 'assets'),
        ('assets/icon_32.png', 'assets'),
        ('assets/dashboard', 'assets/dashboard'),
        ('assets/fonts', 'assets/fonts'),
    ],
    hiddenimports=[
        'pymupdf', 'fitz',
        'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageTk',
        'PIL.ImageEnhance', 'PIL.ImageFilter',
        'tkinter', 'tkinter.ttk', 'tkinter.filedialog',
        'tkinter.messagebox', 'tkinter.simpledialog',
        'docx',
        'cv2', 'numpy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'pandas',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'IPython', 'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    name='OnePDFEditor',
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
    icon='assets/OnePDFEditor.ico',
)
