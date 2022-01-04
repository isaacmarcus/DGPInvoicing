# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['maingui.py'],
             pathex=['C:\\Users\\DGP-Yoga1\\PycharmProjects\\DGPInvoicing'],
             binaries=[],
             datas=[('C:\\Users\\DGP-Yoga1\\AppData\\Local\\Programs\\Python\\Python38-32\\tcl\\tcl8.6', 'lib\\tcl8.6'),
             ('C:\\Users\\DGP-Yoga1\\AppData\\Local\\Programs\\Python\\Python38-32\\tcl\\tk8.6', 'lib\\tk8.6')],
             hiddenimports=['pkg_resources.py2_warn'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
a.datas += [('dg_merge_png.png','C:\\Users\\DGP-Yoga1\\PycharmProjects\\DGPInvoicing\\dg_merge_png.png', 'Data')]
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='InvoiceMergeTool 1.9.3',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False, icon='dg_merge.ico')
