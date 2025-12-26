#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
诊断脚本：运行主窗口，打开/关闭会计工具，收集所有日志
"""
import sys
import os
import time
import threading
from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication

# 设置日志输出到文件
log_file = os.path.join(os.path.dirname(__file__), 'diagnostic.log')
log_handle = open(log_file, 'w', encoding='utf-8')

def log_output(message):
    """将消息输出到文件和控制台"""
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    full_message = f"[{timestamp}] {message}"
    print(full_message)
    log_handle.write(full_message + '\n')
    log_handle.flush()

# 重定向 stderr 和 stdout
class LogWriter:
    def __init__(self, log_func):
        self.log_func = log_func
        self.buffer = []
    
    def write(self, message):
        if message and message.strip():
            self.buffer.append(message)
            if '\n' in message or len(message) > 500:
                self.log_func(''.join(self.buffer).strip())
                self.buffer = []
    
    def flush(self):
        if self.buffer:
            self.log_func(''.join(self.buffer).strip())
            self.buffer = []

sys.stdout = LogWriter(log_output)
sys.stderr = LogWriter(log_output)

log_output("=" * 70)
log_output("诊断程序启动")
log_output("=" * 70)

try:
    log_output("正在导入模块...")
    from main import UnifiedLoginWindow
    from common.logger import get_logger
    
    app_logger = get_logger('Diagnostic')
    log_output("模块导入成功")
    
    app = QApplication(sys.argv)
    
    log_output("创建主窗口...")
    main_window = UnifiedLoginWindow(logger=app_logger)
    main_window.show()
    log_output("主窗口已显示")
    
    # 测试序列
    iteration = [0]
    
    def test_iteration():
        iteration[0] += 1
        log_output(f"\n{'='*70}")
        log_output(f"第 {iteration[0]} 次迭代开始")
        log_output(f"{'='*70}")
        
        if iteration[0] > 3:
            log_output("所有迭代完成。请关闭主窗口以结束诊断。")
            return
        
        # 打开会计工具
        log_output(f"[{iteration[0]}.1] 打开会计核算工具...")
        try:
            main_window.open_kjhs_module()
            log_output("[OK] 会计核算工具已打开")
        except Exception as e:
            log_output(f"[ERROR] 打开会计工具失败: {e}")
            import traceback
            traceback.print_exc()
        
        # 2秒后关闭
        QTimer.singleShot(2000, close_kjhs)
    
    def close_kjhs():
        log_output(f"[{iteration[0]}.2] 关闭会计核算工具...")
        try:
            if main_window.kjhs_window:
                main_window.kjhs_window.close()
                log_output("[OK] 会计核算工具已关闭")
            else:
                log_output("[INFO] 会计核算工具引用为空")
        except Exception as e:
            log_output(f"[ERROR] 关闭会计工具失败: {e}")
            import traceback
            traceback.print_exc()
        
        # 2秒后进行下一次迭代
        QTimer.singleShot(2000, test_iteration)
    
    # 启动测试
    QTimer.singleShot(1000, test_iteration)
    
    log_output("进入事件循环...")
    exit_code = app.exec()
    log_output(f"应用退出，退出代码: {exit_code}")
    
except Exception as e:
    log_output(f"[FATAL] 程序异常: {e}")
    import traceback
    traceback.print_exc()

finally:
    log_output("=" * 70)
    log_output(f"诊断程序结束，日志已保存到: {log_file}")
    log_output("=" * 70)
    log_handle.close()
