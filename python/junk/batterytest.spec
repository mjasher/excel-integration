# -*- mode: python -*-
a = Analysis(['battery.py'],
             pathex=['C:\\Users\\Mike\\Desktop\\excel\\python'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries + [('msvcp100.dll', 'C:\\Windows\\System32\\msvcp100.dll', 'BINARY'),
                        ('msvcr100.dll', 'C:\\Windows\\System32\\msvcr100.dll', 'BINARY')],
          a.zipfiles,
          a.datas,
          name='battery.exe',
          debug=False,
          strip=None,
          upx=True,
          console=True )
