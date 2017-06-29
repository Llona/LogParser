# -*- mode: python -*-

# Need copy all need file to exe folder:
#1. ..\Settings.ini
#2. ..\SubList.sdb
#3. ..\icons\*
#4. ..\NOTICE.txt
#5. ..\LICENSE
#6. ..\change_list.txt


block_cipher = None


a = Analysis(['..\\main.py'],
             pathex=['D:\\ScripFile\\python\\CSV_log_parser'],
             binaries=None,
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
          name='LogAnalysis',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='AJSub.ico')

		  