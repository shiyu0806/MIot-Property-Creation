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
import importlib.util

# Windows 编码修复
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Windows 下 PyInstaller 打印中文会编码报错，用 ASCII 名
APP_NAME = 'MIoT_Tool' if sys.platform == 'win32' else 'MIoT平台工具'

PYQT_EXCLUDES = [
    # Keep the Widgets/WebEngine stack, but avoid collecting broad Qt modules
    # that are not used by this app and inflate the bundle dramatically.
    'PyQt6.Qt3DAnimation',
    'PyQt6.Qt3DCore',
    'PyQt6.Qt3DExtras',
    'PyQt6.Qt3DInput',
    'PyQt6.Qt3DLogic',
    'PyQt6.Qt3DRender',
    'PyQt6.QtBluetooth',
    'PyQt6.QtCharts',
    'PyQt6.QtDataVisualization',
    'PyQt6.QtDesigner',
    'PyQt6.QtGraphs',
    'PyQt6.QtHelp',
    'PyQt6.QtMultimedia',
    'PyQt6.QtMultimediaWidgets',
    'PyQt6.QtNetworkAuth',
    'PyQt6.QtNfc',
    'PyQt6.QtOpenGL',
    'PyQt6.QtOpenGLWidgets',
    'PyQt6.QtPdf',
    'PyQt6.QtPdfWidgets',
    'PyQt6.QtQml',
    'PyQt6.QtQuick',
    'PyQt6.QtQuick3D',
    'PyQt6.QtQuickWidgets',
    'PyQt6.QtRemoteObjects',
    'PyQt6.QtSensors',
    'PyQt6.QtSerialPort',
    'PyQt6.QtSpatialAudio',
    'PyQt6.QtSql',
    'PyQt6.QtStateMachine',
    'PyQt6.QtSvg',
    'PyQt6.QtSvgWidgets',
    'PyQt6.QtTest',
    'PyQt6.QtTextToSpeech',
    'PyQt6.QtWebEngineQuick',
    'PyQt6.QtWebSockets',
    'PyQt6.QtXml',
    'PyQt6.uic',
]

REQUIRED_MODULES = [
    ("PyInstaller", "PyInstaller"),
    ("PyQt6", "PyQt6"),
    ("PyQt6.QtWebEngineWidgets", "PyQt6-WebEngine"),
    ("openpyxl", "openpyxl"),
    ("pandas", "pandas"),
    ("requests", "requests"),
]


def check_dependencies():
    """打包前检查关键依赖，提前给出可执行的修复提示。"""
    missing = []
    for module_name, package_name in REQUIRED_MODULES:
        if importlib.util.find_spec(module_name) is None:
            missing.append(package_name)

    if missing:
        print("✗ 缺少打包依赖:")
        for pkg in missing:
            print(f"  - {pkg}")
        print("\n请先运行: pip install -r requirements.txt")
        sys.exit(1)
    print("✓ 依赖检查通过")

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
    # 使用 --onedir 替代 --onefile。
    # onefile 会把所有文件解压到临时目录运行，而 PyQt6 WebEngine 的 Chromium 子进程
    # 和 macOS Metal/RHI 渲染引擎在临时目录下会崩溃（shader 缓存路径异常导致堆损坏 SIGABRT）。
    # onedir 模式下资源文件在稳定的目录中，不会出现此问题。
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onedir',
        '--windowed',
        '--name', APP_NAME,
        # PyQt6: rely on PyInstaller's Qt hooks for the imported modules.
        # Avoid --collect-all PyQt6; it pulls in Qt3D/QML/Multimedia/SQL/etc.
        '--hidden-import', 'PyQt6.QtWidgets',
        '--hidden-import', 'PyQt6.QtCore',
        '--hidden-import', 'PyQt6.QtGui',
        '--hidden-import', 'PyQt6.QtNetwork',
        '--hidden-import', 'PyQt6.QtWebChannel',
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
        '--hidden-import', 'miot_auth',
        '--hidden-import', 'miot_common',
        '--hidden-import', 'miot_export_template',
        '--hidden-import', 'miot_create_properties',
        '--hidden-import', 'miot_service_core',
        '--hidden-import', 'miot_automation_core',
        '--hidden-import', 'miot_gui_automation_tabs',
        '--hidden-import', 'miot_gui_auth_ui',
        '--hidden-import', 'miot_gui_common',
        '--hidden-import', 'miot_gui_property_tabs',
        '--hidden-import', 'miot_gui_service_tabs',
        '--hidden-import', 'miot_gui_styles',
        '--hidden-import', 'miot_gui_workers',
        '--hidden-import', 'miot_reports',
        '--hidden-import', 'create_template',
        'miot_gui.py'
    ]

    for module_name in PYQT_EXCLUDES:
        cmd.extend(['--exclude-module', module_name])

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
    check_dependencies()
    clean()
    build()
