#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单测试脚本：测试打开和关闭会计核算工具
"""
import sys
import traceback
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import QTimer
from common.logger import get_logger

logger = get_logger('TestSimple')

class SimpleTestWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.kjhs_window = None
        self.init_ui()
        logger.info("SimpleTestWindow 已初始化")
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        label = QLabel("点击按钮打开会计核算工具")
        layout.addWidget(label)
        
        btn = QPushButton("打开会计核算工具")
        btn.clicked.connect(self.open_kjhs)
        layout.addWidget(btn)
        
        self.setLayout(layout)
        self.setWindowTitle("简单测试")
        self.setGeometry(100, 100, 300, 200)
    
    def open_kjhs(self):
        logger.info("用户点击了打开按钮")
        try:
            from kjhs_test.pld_pyqt6 import MainWindow
            logger.info("成功导入 MainWindow")
            
            if self.kjhs_window:
                logger.info("窗口已存在，激活窗口")
                self.kjhs_window.activateWindow()
                self.kjhs_window.raise_()
            else:
                logger.info("创建新的 MainWindow")
                self.kjhs_window = MainWindow(logger=logger)
                logger.info("MainWindow 创建完成")
                
                # 连接 destroyed 信号
                try:
                    def on_destroyed():
                        logger.info("会计核算窗口被销毁")
                        self.kjhs_window = None
                    self.kjhs_window.destroyed.connect(on_destroyed)
                    logger.info("已连接 destroyed 信号")
                except Exception as e:
                    logger.error(f"连接 destroyed 信号失败: {e}")
                
                self.kjhs_window.show()
                logger.info("MainWindow 已显示")
        except Exception as e:
            logger.error(f"打开会计核算工具失败: {e}", exc_info=True)
            traceback.print_exc()

def main():
    logger.info("=" * 60)
    logger.info("启动简单测试程序")
    logger.info("=" * 60)
    
    app = QApplication(sys.argv)
    
    # 设置全局异常处理
    def exception_handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logger.error("未处理的异常:", exc_info=(exc_type, exc_value, exc_traceback))
        print(f"FATAL ERROR: {exc_type.__name__}: {exc_value}", file=sys.stderr)
        import traceback as tb
        tb.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
    
    sys.excepthook = exception_handler
    
    try:
        window = SimpleTestWindow()
        window.show()
        logger.info("主窗口已显示，开始事件循环")
        exit_code = app.exec()
        logger.info(f"应用正常退出，退出代码: {exit_code}")
        return exit_code
    except Exception as e:
        logger.error(f"主程序异常: {e}", exc_info=True)
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    try:
        sys.exit(main())
    except SystemExit as e:
        logger.info(f"程序以退出代码 {e.code} 结束")
        sys.exit(e.code)
    except Exception as e:
        logger.error(f"最顶层异常: {e}", exc_info=True)
        sys.exit(1)
