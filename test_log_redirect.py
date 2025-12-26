#!/usr/bin/env python3
"""
测试日志重定向功能
"""

import sys
import os
import logging

# 添加ywjlb目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'ywjlb'))

from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QTextEdit, QWidget, QPushButton
from PyQt6.QtCore import pyqtSignal

# 导入我们的类
from ywjlb.ywjlb_ui import UILogHandler

class TestWindow(QMainWindow):
    log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("日志重定向测试")
        self.setGeometry(100, 100, 600, 400)

        # 创建日志文本框
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # 创建测试按钮
        self.test_button = QPushButton("测试日志")
        self.test_button.clicked.connect(self.test_logging)

        layout = QVBoxLayout()
        layout.addWidget(self.log_text)
        layout.addWidget(self.test_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # 设置日志处理器
        ui_handler = UILogHandler(self)
        ui_handler.setLevel(logging.DEBUG)

        # 获取根logger并添加处理器
        root_logger = logging.getLogger()
        root_logger.addHandler(ui_handler)
        self.original_root_level = root_logger.level
        if root_logger.level > logging.DEBUG:
            root_logger.setLevel(logging.DEBUG)

        # 连接信号
        self.log_signal.connect(self.append_log)

    def test_logging(self):
        """测试日志输出"""
        import logging
        logging.info("这是INFO级别的测试日志")
        logging.warning("这是WARNING级别的测试日志")
        logging.error("这是ERROR级别的测试日志")
        logging.debug("这是DEBUG级别的测试日志")

    def append_log(self, message):
        current_text = self.log_text.toPlainText()
        if current_text:
            self.log_text.setText(current_text + "\n" + message)
        else:
            self.log_text.setText(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TestWindow()
    window.show()
    sys.exit(app.exec())