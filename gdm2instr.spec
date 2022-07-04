# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['gdm2instr.py'],
             pathex=['C:\\Users\\grokh\\Desktop\\gdm2instr'],
             binaries=[],
             datas=[],
             hiddenimports=['pandas'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='gdm2instr',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          icon='radioactive.ico') # if you need it, provide it in the dir