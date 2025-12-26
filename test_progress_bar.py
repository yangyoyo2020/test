#!/usr/bin/env python3
"""
测试进度条功能
"""

import sys
import os
import time

# 添加ywjlb目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'ywjlb'))

from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QProgressBar, QGroupBox, QLabel
from PyQt6.QtCore import Qt

class ProgressTestWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("进度条测试")
        self.setGeometry(100, 100, 400, 200)

        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪")
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 创建状态标签
        self.status_label = QLabel("点击开始测试进度条")

        layout = QVBoxLayout()
        layout.addWidget(QLabel("进度条测试"))
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.status_label)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # 模拟进度更新
        self.simulate_progress()

    def simulate_progress(self):
        """模拟进度更新"""
        self.status_label.setText("开始模拟进度...")

        # 模拟处理10个项目
        total_items = 10
        for i in range(total_items + 1):
            percentage = int((i / total_items) * 100)
            self.progress_bar.setValue(percentage)
            self.progress_bar.setFormat(f"处理进度: {i}/{total_items} ({percentage}%)")
            self.status_label.setText(f"正在处理第 {i} 项...")
            QApplication.processEvents()  # 强制更新UI
            time.sleep(0.5)

        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("处理完成")
        self.status_label.setText("进度条测试完成！")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ProgressTestWindow()
    window.show()
    sys.exit(app.exec())