# -*- mode: python ; coding: utf-8 -*-

import sys
sys.setrecursionlimit(5000)  # 增加递归限制

block_cipher = None

# 定义需要添加的数据文件
added_files = [
    ('check_standby.bat', '.'),
    ('check_standby.sql', '.'),
    ('daily_report.bat', '.'),
    ('daily_report_dg.sql', '.'),
    ('daily_report_prod.sql', '.'),
    ('start_standby.sql', '.'),
    ('opera_monitor.ini', '.'),
    ('logs', 'logs'),
]

# 创建Analysis对象
a = Analysis(
    ['opera_monitor.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 创建PYZ对象（Python的ZIP归档）
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 创建EXE对象（目录模式）
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Opera数据库监控工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
    version='1.0.0',
    uac_admin=False,
)

# 创建COLLECT对象（收集所有文件到目录中）
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Opera数据库监控工具',
)

# 单文件模式的EXE对象（可选，取消注释以启用）
'''
onefile_exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Opera数据库监控工具_单文件版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
'''