#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
运维记录簿处理工具 - 智能启动脚本

此脚本自动处理导入路径，无论从何处运行都能正常启动应用
"""

import sys
from pathlib import Path

def setup_path():
    """设置Python路径，支持从任何位置运行"""
    # 获取当前脚本所在目录
    script_dir = Path(__file__).parent.absolute()
    
    # 添加脚本目录到sys.path的最前面
    if str(script_dir) not in sys.path:
        sys.path.insert(0, str(script_dir))
    
    return script_dir


def main():
    """主入口函数"""
    try:
        # 设置Python路径
        script_dir = setup_path()
        print(f"✓ 工作目录：{script_dir}")
        
        # 导入所需模块
        from PyQt6.QtWidgets import QApplication
        from ywjlb_ui import YWJLBAnalyzer
        
        print("✓ 模块导入成功")
        
        # 创建应用
        app = QApplication(sys.argv)
        
        # 设置应用样式
        app.setStyle('Fusion')
        
        # 创建并显示主窗口
        print("✓ 启动用户界面...")
        window = YWJLBAnalyzer()
        window.show()
        
        # 启动事件循环
        exit_code = app.exec()
        sys.exit(exit_code)
        
    except ImportError as e:
        print(f"✗ 导入错误：{e}")
        print("\n请确保已安装所有必需的包：")
        print("  pip install pandas python-docx PyQt6")
        sys.exit(1)
        
    except Exception as e:
        print(f"✗ 错误：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
