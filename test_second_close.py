#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
最小化测试：打开->关闭->打开->关闭，观察程序是否仍在运行
"""
import sys
from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication

try:
    print("=" * 70)
    print("启动程序测试：第二次打开/关闭不应导致程序退出")
    print("=" * 70)
    
    from main import UnifiedLoginWindow
    from common.logger import get_logger
    
    app_logger = get_logger('Test')
    
    app = QApplication(sys.argv)
    main_window = UnifiedLoginWindow(logger=app_logger)
    main_window.show()
    
    test_state = {'step': 0}
    
    def step_1():
        """第一次打开"""
        test_state['step'] += 1
        print(f"\n[步骤 {test_state['step']}] 第一次打开会计核算工具...")
        try:
            main_window.open_kjhs_module()
            print("[OK] 窗口已打开")
        except Exception as e:
            print(f"[ERROR] {e}")
            import traceback
            traceback.print_exc()
    
    def step_2():
        """第一次关闭"""
        test_state['step'] += 1
        print(f"\n[步骤 {test_state['step']}] 第一次关闭会计核算工具...")
        try:
            if main_window.kjhs_window:
                main_window.kjhs_window.close()
                print("[OK] 窗口已关闭")
            else:
                print("[INFO] 窗口引用为空")
        except Exception as e:
            print(f"[ERROR] {e}")
            import traceback
            traceback.print_exc()
    
    def step_3():
        """第二次打开"""
        test_state['step'] += 1
        print(f"\n[步骤 {test_state['step']}] 第二次打开会计核算工具...")
        try:
            main_window.open_kjhs_module()
            print("[OK] 窗口已打开（如果程序仍在这里，说明没有自动关闭！）")
        except Exception as e:
            print(f"[ERROR] {e}")
            import traceback
            traceback.print_exc()
    
    def step_4():
        """第二次关闭"""
        test_state['step'] += 1
        print(f"\n[步骤 {test_state['step']}] 第二次关闭会计核算工具...")
        try:
            if main_window.kjhs_window:
                main_window.kjhs_window.close()
                print("[OK] 窗口已关闭")
                print("\n[SUCCESS] 测试成功！程序没有自动关闭")
                print("=" * 70)
                print("请点击主窗口的'退出系统'按钮以正常关闭程序")
                print("=" * 70)
            else:
                print("[INFO] 窗口引用为空")
        except Exception as e:
            print(f"[ERROR] {e}")
            import traceback
            traceback.print_exc()
    
    # 按顺序执行步骤
    QTimer.singleShot(1000, step_1)
    QTimer.singleShot(3000, step_2)
    QTimer.singleShot(5000, step_3)
    QTimer.singleShot(7000, step_4)
    
    print("进入事件循环...")
    sys.exit(app.exec())

except Exception as e:
    print(f"[FATAL] {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
