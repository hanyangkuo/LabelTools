# -*- mode: python -*-

block_cipher = None


a = Analysis(['D:\\project\\Project\\LabelTools\\LabelTools_v1.3\\src\\main\\python\\main.py'],
             pathex=['D:\\project\\Project\\LabelTools\\LabelTools_v1.3\\target\\PyInstaller'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=['D:\\Program Files\\Python311\\Lib\\site-packages\\fbs\\freeze\\hooks'],
             runtime_hooks=['D:\\project\\Project\\LabelTools\\LabelTools_v1.3\\target\\PyInstaller\\fbs_pyinstaller_hook.py'],
             excludes=['_bootlocale'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='LabelTools_v1.3',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          console=False , version='D:\\project\\Project\\LabelTools\\LabelTools_v1.3\\target\\PyInstaller\\version_info.py', icon='D:\\project\\Project\\LabelTools\\LabelTools_v1.3\\src\\main\\icons\\Icon.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               name='LabelTools_v1.3')
