block_cipher = None

a = Analysis(
    ['pagecapture.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PIL._tkinter_finder',
        'comtypes.gen',
        'comtypes.gen.UIAutomationClient',
        'win32api',
        'win32con',
        'win32gui',
        'win32process',
        'pyperclip',
        'docx',
        'docx.shared',
        'docx.oxml',
        'docx.oxml.ns',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib','numpy','scipy','pandas','PyQt5','wx','keyboard','pystray'],
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True,
    name='PageCapture', debug=False, strip=False, upx=True,
    console=False, icon='assets/icon.ico')
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, name='PageCapture')
