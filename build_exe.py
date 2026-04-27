#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本（macOS / Windows 通用）
macOS 输出 dist/MIoT平台工具.app
Windows 输出 dist/MIoT平台工具.exe
"""
import subprocess
import sys
import os
import shutil

# Windows 编码修复
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Windows 下 PyInstaller 打印中文会编码报错，用 ASCII 名
APP_NAME = 'MIoT_Tool' if sys.platform == 'win32' else 'MIoT平台工具'

def clean():
    """清理旧文件"""
    for d in ['build', 'dist']:
        if os.path.exists(d):
            shutil.rmtree(d)
    spec = f'{APP_NAME}.spec'
    if os.path.exists(spec):
        os.remove(spec)
    print("✓ 已清理旧文件")

def build():
    """打包"""
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--windowed',
        '--name', APP_NAME,
        # PyQt6
        '--collect-all', 'PyQt6',
        '--collect-all', 'PyQt6-Qt6',
        '--collect-all', 'PyQt6WebEngine',
        '--hidden-import', 'PyQt6.QtWidgets',
        '--hidden-import', 'PyQt6.QtCore',
        '--hidden-import', 'PyQt6.QtGui',
        '--hidden-import', 'PyQt6.QtWebEngineWidgets',
        '--hidden-import', 'PyQt6.QtWebEngineCore',
        # openpyxl
        '--collect-all', 'openpyxl',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'openpyxl.styles',
        '--hidden-import', 'openpyxl.worksheet.datavalidation',
        # pandas（服务层新增）
        '--hidden-import', 'pandas',
        # 项目模块
        '--hidden-import', 'miot_export_template',
        '--hidden-import', 'miot_create_properties',
        '--hidden-import', 'miot_service_core',
        '--hidden-import', 'miot_automation_core',
        '--hidden-import', 'create_template',
        'miot_gui.py'
    ]

    # macOS 图标（如果有）
    if sys.platform == 'darwin' and os.path.exists('icon.icns'):
        cmd.extend(['--icon', 'icon.icns'])
    elif sys.platform == 'win32' and os.path.exists('icon.ico'):
        cmd.extend(['--icon', 'icon.ico'])

    print("开始打包...")
    print(' '.join(cmd))
    result = subprocess.run(cmd, capture_output=False)

    if result.returncode == 0:
        print(f"\n✓ 打包成功！")
        ext = '.app' if sys.platform == 'darwin' else '.exe'
        print(f"输出路径: {os.path.abspath(f'dist/{APP_NAME}{ext}')}")
    else:
        print("\n✗ 打包失败")
        sys.exit(1)

if __name__ == '__main__':
    clean()
    build()
