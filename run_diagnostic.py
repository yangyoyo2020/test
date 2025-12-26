#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
包装脚本：运行诊断程序并收集所有输出
"""
import subprocess
import sys
import os

diagnostic_file = os.path.join(os.path.dirname(__file__), 'diagnostic.log')

# 清除旧的诊断文件
if os.path.exists(diagnostic_file):
    try:
        os.remove(diagnostic_file)
    except:
        pass

print(f"运行诊断程序，日志将保存到: {diagnostic_file}")
print("=" * 70)

try:
    # 运行诊断脚本
    process = subprocess.Popen(
        [sys.executable, 'diagnostic.py'],
        cwd=os.path.dirname(__file__),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True
    )
    
    # 实时输出
    for line in process.stdout:
        print(line, end='')
    
    returncode = process.wait()
    print("=" * 70)
    print(f"进程退出，返回代码: {returncode}")
    
except Exception as e:
    print(f"运行失败: {e}")
    sys.exit(1)

# 读取并显示诊断日志
if os.path.exists(diagnostic_file):
    print("\n诊断日志内容:")
    print("=" * 70)
    with open(diagnostic_file, 'r', encoding='utf-8') as f:
        print(f.read())
