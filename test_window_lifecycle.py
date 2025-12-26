#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试窗口生命周期：重复打开/关闭会计核算和JSON工具以验证修复
"""
import sys
import time
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer
from main import UnifiedLoginWindow
from common.logger import get_logger

logger = get_logger('TestWindowLifecycle')


class TestWindowLifecycle:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.main_window = UnifiedLoginWindow(logger=logger)
        self.main_window.show()
        
        self.test_step = 0
        self.max_iterations = 3  # 重复打开/关闭 3 次
        self.current_iteration = 0
        
        # 启动自动化测试
        QTimer.singleShot(1000, self.start_test)
    
    def start_test(self):
        logger.info("=" * 60)
        logger.info("开始窗口生命周期测试")
        logger.info("=" * 60)
        self.run_test_iteration()
    
    def run_test_iteration(self):
        self.current_iteration += 1
        logger.info(f"\n【迭代 {self.current_iteration}/{self.max_iterations}】")
        
        if self.current_iteration > self.max_iterations:
            logger.info("\n所有测试完成！如果程序仍在运行，说明修复成功。")
            logger.info("请手动关闭主窗口以结束测试。")
            return
        
        # 步骤1: 打开会计核算工具
        logger.info(f"  [{self.current_iteration}.1] 打开会计核算工具...")
        self.main_window.open_kjhs_module()
        
        # 1秒后关闭
        QTimer.singleShot(1000, self.close_kjhs)
    
    def close_kjhs(self):
        logger.info(f"  [{self.current_iteration}.2] 关闭会计核算工具...")
        if self.main_window.kjhs_window:
            try:
                self.main_window.kjhs_window.close()
                logger.info("  会计核算工具已关闭")
            except Exception as e:
                logger.error(f"关闭失败: {e}")
        
        # 1秒后打开JSON工具
        QTimer.singleShot(1000, self.open_json)
    
    def open_json(self):
        logger.info(f"  [{self.current_iteration}.3] 打开JSON转Excel工具...")
        self.main_window.open_json_module()
        
        # 1秒后关闭
        QTimer.singleShot(1000, self.close_json)
    
    def close_json(self):
        logger.info(f"  [{self.current_iteration}.4] 关闭JSON转Excel工具...")
        if self.main_window.json_window:
            try:
                self.main_window.json_window.close()
                logger.info("  JSON转Excel工具已关闭")
            except Exception as e:
                logger.error(f"关闭失败: {e}")
        
        # 1秒后进行下一次迭代
        QTimer.singleShot(1000, self.run_test_iteration)
    
    def run(self):
        sys.exit(self.app.exec())


if __name__ == '__main__':
    test = TestWindowLifecycle()
    test.run()
