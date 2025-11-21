import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QFrame, QGridLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

# 导入各个模块的主窗口类
from sanbao_test.app_copy import ExpenditureAnalyzer
from kjhs_test.pld_pyqt6 import MainWindow
from json_to_excel.json_to_excel_pyqt import JSONToExcelConverter


class UnifiedLoginWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 保存对各模块窗口的引用，防止被垃圾回收
        self.sanbao_window = None
        self.kjhs_window = None
        self.json_window = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("数据分析工具")
        self.setGeometry(300, 150, 600, 400)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        
        # 标题
        title_label = QLabel("请选择要使用的功能模块")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(separator)
        
        # 按钮网格布局
        grid_layout = QGridLayout()
        grid_layout.setSpacing(15)
        
        # 三保支出分析按钮
        sanbao_btn = QPushButton("三保支出进度工具")
        sanbao_btn.setMinimumHeight(80)
        sanbao_btn.clicked.connect(self.open_sanbao_module)
        sanbao_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        grid_layout.addWidget(sanbao_btn, 0, 0)
        
        # 会计核算偏离度工具按钮
        kjhs_btn = QPushButton("会计核算偏离度工具")
        kjhs_btn.setMinimumHeight(80)
        kjhs_btn.clicked.connect(self.open_kjhs_module)
        kjhs_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        grid_layout.addWidget(kjhs_btn, 0, 1)
        
        # JSON转Excel按钮
        json_btn = QPushButton("JSON转Excel工具")
        json_btn.setMinimumHeight(80)
        json_btn.clicked.connect(self.open_json_module)
        json_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:pressed {
                background-color: #EF6C00;
            }
        """)
        grid_layout.addWidget(json_btn, 1, 0)
        
        # 退出按钮
        exit_btn = QPushButton("退出系统")
        exit_btn.setMinimumHeight(80)
        exit_btn.clicked.connect(self.close)
        exit_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #b71c1c;
            }
        """)
        grid_layout.addWidget(exit_btn, 1, 1)
        
        main_layout.addLayout(grid_layout)
        
        # 状态栏信息
        status_label = QLabel("© 数据分析工具")
        status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(status_label)
        
    def open_sanbao_module(self):
        # 创建三保支出分析窗口并保持引用
        self.sanbao_window = ExpenditureAnalyzer()
        self.sanbao_window.show()
        
    def open_kjhs_module(self):
        # 创建会计核算偏离度工具窗口并保持引用
        self.kjhs_window = MainWindow()
        self.kjhs_window.show()
        
    def open_json_module(self):
        # 创建JSON转Excel工具窗口并保持引用
        self.json_window = JSONToExcelConverter()
        self.json_window.show()


def main():
    app = QApplication(sys.argv)
    window = UnifiedLoginWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()