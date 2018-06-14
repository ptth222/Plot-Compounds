# -*- mode: python -*-

block_cipher = None


a = Analysis(['Make_Compound_Plots.py'],
             pathex=['/Users/higashi/Plot_Compounds'],
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
          name='Make_Compound_Plots',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
app = BUNDLE(exe,
             name='Make_Compound_Plots.app',
             icon=None,
             bundle_identifier=None,
             info_plist={
		'NSHighResolutionCapable': 'True'
		},
	     )
