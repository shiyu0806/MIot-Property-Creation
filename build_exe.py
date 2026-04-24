#!/usr/bin/env python3
"""
Windows EXE 打包脚本
"""
import subprocess
import sys
import os
import shutil

def clean():
    """清理旧文件"""
    for d in ['build', 'dist']:
        if os.path.exists(d):
            shutil.rmtree(d)
    for f in ['MIoT属性工具.spec']:
        if os.path.exists(f):
            os.remove(f)
    print("✓ 已清理旧文件")

def build():
    """打包"""
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--windowed',
        '--name', 'MIoT属性工具',
        '--collect-all', 'PyQt6',
        '--collect-all', 'PyQt6-Qt6',
        '--collect-all', 'openpyxl',
        '--hidden-import', 'PyQt6.QtWidgets',
        '--hidden-import', 'PyQt6.QtCore',
        '--hidden-import', 'PyQt6.QtGui',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'openpyxl.styles',
        '--hidden-import', 'openpyxl.worksheet.datavalidation',
        '--hidden-import', 'miot_export_template',
        '--hidden-import', 'miot_create_properties',
        '--hidden-import', 'create_template',
        'miot_gui.py'
    ]
    
    print("开始打包...")
    print(' '.join(cmd))
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        print("\n✓ 打包成功！")
        print(f"输出路径: {os.path.abspath('dist/MIoT属性工具.exe')}")
    else:
        print("\n✗ 打包失败")
        sys.exit(1)

if __name__ == '__main__':
    clean()
    build()
