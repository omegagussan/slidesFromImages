# -*- mode: python -*-

block_cipher = None

a = Analysis(['images2Slides.py'],
             pathex=['/home/gussan/Documents/code/slidesFromImages'],
             binaries=[],
             datas=[("./default.pptx", "pptx/templates")],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='images2slides',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
