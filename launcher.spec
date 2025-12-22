# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 获取当前目录
import os
current_dir = os.getcwd()

a = Analysis(
    ['app_launcher.py'],
    pathex=[current_dir],
    binaries=[],
    datas=[
        ('icon.ico', '.'),
    ],
    hiddenimports=[
        'psutil', 
        'win32com',
        'win32com.client',
        'pywintypes',
        'ctypes',
        'json',
        'os',
        'sys',
        'time'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'unittest', 'pytest', 'matplotlib', 'numpy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data,
          cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='程序启动管理器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
    version='version_info.txt'  # 可选，创建版本信息文件
)

# 可选：创建版本信息文件
if not os.path.exists('version_info.txt'):
    with open('version_info.txt', 'w') as f:
        f.write('VSVersionInfo(\n')
        f.write('  ffi=FixedFileInfo(\n')
        f.write('    filevers=(1, 0, 0, 0),\n')
        f.write('    prodvers=(1, 0, 0, 0),\n')
        f.write('    mask=0x3f,\n')
        f.write('    flags=0x0,\n')
        f.write('    OS=0x40004,\n')
        f.write('    fileType=0x1,\n')
        f.write('    subtype=0x0,\n')
        f.write('    date=(0, 0)\n')
        f.write('  ),\n')
        f.write('  kids=[\n')
        f.write('    StringFileInfo(\n')
        f.write('      [\n')
        f.write('        StringTable(\n')
        f.write('          u\'040904B0\',\n')
        f.write('          [StringStruct(u\'CompanyName\', u\'Personal\'),\n')
        f.write('          StringStruct(u\'FileDescription\', u\'程序启动管理器\'),\n')
        f.write('          StringStruct(u\'FileVersion\', u\'1.0.0.0\'),\n')
        f.write('          StringStruct(u\'InternalName\', u\'app_launcher\'),\n')
        f.write('          StringStruct(u\'LegalCopyright\', u\'Copyright © 2023\'),\n')
        f.write('          StringStruct(u\'OriginalFilename\', u\'app_launcher.exe\'),\n')
        f.write('          StringStruct(u\'ProductName\', u\'程序启动管理器\'),\n')
        f.write('          StringStruct(u\'ProductVersion\', u\'1.0.0.0\')])\n')
        f.write('      ]),\n')
        f.write('    VarFileInfo([VarStruct(u\'Translation\', [0x0409, 1200])])\n')
        f.write('  ]\n')
        f.write(')\n')