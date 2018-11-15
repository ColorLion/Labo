# -*- mode: python -*-

block_cipher = None


a = Analysis(['Get_Plist.py'],
             pathex=['E:\\4.git\\Get_Plist'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Get_Plist',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
