#!/usr/bin/env python3
"""语法检查脚本"""
import py_compile

files = [
    'miot_common.py',
    'miot_service_core.py', 
    'miot_create_properties.py',
    'miot_automation_core.py',
    'miot_gui.py',
    'create_template.py',
    'capture_api.py',
    'miot_auth.py',
    'miot_export_template.py',
    'build_exe.py',
]

ok = 0
for f in files:
    try:
        py_compile.compile(f, doraise=True)
        print(f'OK {f}')
        ok += 1
    except py_compile.PyCompileError as e:
        print(f'FAIL {f}: {e}')

print(f'\nTotal: {ok}/{len(files)} files syntax OK')
