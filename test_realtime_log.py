#!/usr/bin/env python3
"""
测试实时日志重定向功能
"""

import sys
import os
import logging
import time

from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QTextEdit, QWidget, QPushButton
from PyQt6.QtCore import pyqtSignal

class UILogHandler(logging.Handler):
    """自定义日志处理器，将日志消息发送到UI"""
    
    def __init__(self, signal_emitter):
        super().__init__()
        self.signal_emitter = signal_emitter
        self.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
    
    def emit(self, record):
        """处理日志记录"""
        try:
            msg = self.format(record)
            # 立即发射信号，不等待
            self.signal_emitter.log_signal.emit(msg)
        except Exception as e:
            # 如果信号发射失败，至少打印到控制台
            print(f"日志处理器错误: {e}")
            print(f"原始消息: {msg}")

class TestWindow(QMainWindow):
    log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("实时日志重定向测试")
        self.setGeometry(100, 100, 600, 400)

        # 创建日志文本框
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # 创建测试按钮
        self.test_button = QPushButton("模拟process_excel_file")
        self.test_button.clicked.connect(self.simulate_process)

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

        print(f"日志处理器已设置，根logger级别: {root_logger.level}")
        print(f"处理器数量: {len(root_logger.handlers)}")

    def simulate_process(self):
        """模拟process_excel_file的日志输出"""
        import logging

        # 模拟process_excel_file的日志
        logging.info("开始转换Excel数据到Word文档... (包类型: 01包-金财工程)")
        time.sleep(0.5)  # 模拟处理时间

        logging.info("输出文件夹: 运维记录文档")
        time.sleep(0.5)

        logging.info("成功读取Excel文件，共2个工作表")
        time.sleep(0.5)

        logging.info("开始处理工作表: Sheet1")
        time.sleep(0.5)

        logging.info("工作表[Sheet1]包含5行数据")
        time.sleep(0.5)

        for i in range(1, 6):
            logging.debug(f"开始处理工作表[Sheet1]的第{i}行数据")
            time.sleep(0.2)
            logging.info(f"成功生成: 运维记录_Sheet1_第{i}页.docx")
            time.sleep(0.2)

        logging.info("工作表[Sheet1]已处理5/5行数据")
        time.sleep(0.5)

        logging.info("开始处理工作表: Sheet2")
        time.sleep(0.5)

        logging.info("工作表[Sheet2]包含3行数据")
        time.sleep(0.5)

        for i in range(1, 4):
            logging.debug(f"开始处理工作表[Sheet2]的第{i}行数据")
            time.sleep(0.2)
            logging.info(f"成功生成: 运维记录_Sheet2_第{i}页.docx")
            time.sleep(0.2)

        logging.info("转换完成！成功生成8个文档，失败0个")

    def append_log(self, message):
        current_text = self.log_text.toPlainText()
        if current_text:
            self.log_text.setText(current_text + "\n" + message)
        else:
            self.log_text.setText(message)

        # 自动滚动到底部
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

        # 强制处理事件，确保UI立即更新
        from PyQt6.QtWidgets import QApplication
        QApplication.processEvents()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TestWindow()
    window.show()
    sys.exit(app.exec())