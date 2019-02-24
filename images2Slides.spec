# -*- mode: python -*-

block_cipher = None

a = Analysis(['images2Slides.py'],
             pathex=['/home/gussan/Documents/code/slidesFromImages'],
             binaries=[],
             datas=[("/home/gussan/anaconda3/lib/python3.6/site-packages/pptx/templates/default.pptx", "pptx/templates")],
             hiddenimports=[],
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
          [],
          exclude_binaries=True,
          name='images2Slides',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='images2Slides')
