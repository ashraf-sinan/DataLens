# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'numpy',
        'tkinter',
        'openpyxl.chart',
        'openpyxl.styles',
        'openpyxl.chart.bar_chart',
        'openpyxl.chart.pie_chart',
        'openpyxl.chart.reference',
        'openpyxl.utils',
        'openpyxl.worksheet',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'pytest',
        'IPython',
        'ipython',
        'IPython.core',
        'jupyter',
        'notebook',
        'sphinx',
        'pip',
        'setuptools',
        'wheel',
        'tornado',
        'zmq',
        'nbconvert',
        'nbformat',
        'jedi',
        'pygments',
        'prompt_toolkit',
        'traitlets',
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
    name='DataLens',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Disable UPX compression to avoid decompression errors
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # You can add an icon file here if you have one
)
