import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QFrame, QGridLayout, QHBoxLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QAction
from common import styles
from common.animated import AnimatedButton
from common.window_utils import open_or_activate

# 导入各个模块的主窗口类
from sanbao_test.app_copy import ExpenditureAnalyzer
from kjhs_test.pld_pyqt6 import MainWindow
from json_to_excel.json_to_excel_pyqt import JSONToExcelConverter
from common.logger import get_logger
# 导入运维记录簿主窗口
from ywjlb.ywjlb_ui import YWJLBAnalyzer


class UnifiedLoginWindow(QMainWindow):
    def __init__(self, logger=None):
        super().__init__()
        # 注入的全局 logger（可为 None）
        self.logger = logger
        # 保存对各模块窗口的引用，防止被垃圾回收
        self.sanbao_window = None
        self.kjhs_window = None
        self.json_window = None
        self.ywjlb_window = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("数据分析工具")
        self.setGeometry(300, 150, 600, 400)
        
        # 创建中央部件并设置统一的外边距（容器 + 卡片）
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        outer_layout = QVBoxLayout(central_widget)
        outer_layout.setContentsMargins(18, 18, 18, 18)
        outer_layout.setSpacing(12)

        # 使用卡片承载主内容，卡片样式在 QSS 中定义为 QFrame#card
        card = QFrame()
        card.setObjectName('card')
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(16)
        
        # 标题
        title_label = QLabel("请选择要使用的功能模块")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        card_layout.addWidget(title_label)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        card_layout.addWidget(separator)
        
        # 按钮网格布局
        grid_layout = QGridLayout()
        grid_layout.setSpacing(12)
        
        # 三保支出分析按钮
        sanbao_btn = AnimatedButton("1.三保支出进度工具")
        sanbao_btn.setMinimumHeight(80)
        sanbao_btn.clicked.connect(self.open_sanbao_module)
        # 使用全局样式，设置 variant 以选择次级样式（green）
        sanbao_btn.setProperty('variant', 'secondary')
        sanbao_btn.setObjectName('main_sanbao_btn')
        grid_layout.addWidget(sanbao_btn, 0, 0)
        
        # 会计核算偏离度工具按钮
        kjhs_btn = AnimatedButton("2.会计核算偏离度工具")
        kjhs_btn.setMinimumHeight(80)
        kjhs_btn.clicked.connect(self.open_kjhs_module)
        kjhs_btn.setProperty('variant', 'primary')
        kjhs_btn.setObjectName('main_kjhs_btn')
        grid_layout.addWidget(kjhs_btn, 0, 1)
        
        # JSON转Excel按钮
        json_btn = AnimatedButton("3.JSON转Excel工具")
        json_btn.setMinimumHeight(80)
        json_btn.clicked.connect(self.open_json_module)
        json_btn.setProperty('variant', 'warning')
        json_btn.setObjectName('main_json_btn')
        grid_layout.addWidget(json_btn, 1, 0)
        
        # 运维记录簿按钮
        ywjlb_btn = AnimatedButton("4.运维记录表转换工具")
        ywjlb_btn.setMinimumHeight(80)
        ywjlb_btn.clicked.connect(self.open_ywjlb_module)
        ywjlb_btn.setProperty('variant', 'info')
        ywjlb_btn.setObjectName('main_ywjlb_btn')
        grid_layout.addWidget(ywjlb_btn, 1, 1)
        
        # （退出按钮已移到卡片底部右侧，和其他工具分离）
        
        card_layout.addLayout(grid_layout)

        # 在卡片底部添加一个右对齐的退出按钮，使其独立于工具网格
        card_layout.addStretch()
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(0, 0, 0, 0)
        footer_layout.addStretch()
        exit_btn = AnimatedButton("退出系统")
        exit_btn.setMinimumHeight(40)
        exit_btn.clicked.connect(self.close)
        exit_btn.setProperty('variant', 'danger')
        exit_btn.setObjectName('main_exit_btn')
        footer_layout.addWidget(exit_btn)
        card_layout.addLayout(footer_layout)

        # 将卡片添加到外层布局
        outer_layout.addWidget(card)

        # 状态栏信息（放在卡片下方）
        status_row = QWidget()
        status_layout = QHBoxLayout(status_row)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_label = QLabel("© 数据分析工具")
        # 垂直+水平居中显示
        status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # 使用左右伸缩将标签推到中间
        status_layout.addStretch()
        status_layout.addWidget(status_label)
        status_layout.addStretch()
        outer_layout.addWidget(status_row)

        # 在菜单栏添加“工具”菜单，并把切换主题功能放到菜单中
        try:
            tools_menu = self.menuBar().addMenu("工具")
            current = styles.get_current_theme()
            cur_text = '深色' if current == 'dark' else '浅色'
            toggle_action = QAction(f"切换主题（当前：{cur_text}）", self)

            def _on_toggle():
                new_theme = styles.toggle_theme(QApplication.instance())
                new_text = '深色' if new_theme == 'dark' else '浅色'
                try:
                    toggle_action.setText(f"切换主题（当前：{new_text}）")
                except Exception:
                    pass

            toggle_action.triggered.connect(_on_toggle)
            tools_menu.addAction(toggle_action)
        except Exception:
            # 在极少数非标准环境中 menuBar 可能不可用，忽略不阻塞
            pass
        
    def open_sanbao_module(self):
        # 统一使用通用工具：复用已打开的窗口或创建新窗口
        open_or_activate(self, 'sanbao_window', lambda: ExpenditureAnalyzer(logger=getattr(self, 'logger', None)))
        
    def open_kjhs_module(self):
        open_or_activate(self, 'kjhs_window', lambda: MainWindow(logger=getattr(self, 'logger', None)))
        
    def open_json_module(self):
        open_or_activate(self, 'json_window', lambda: JSONToExcelConverter(logger=getattr(self, 'logger', None)))

    def open_ywjlb_module(self):
        # 打开运维记录表窗口
        open_or_activate(self, 'ywjlb_window', lambda: YWJLBAnalyzer())


def main():
    app = QApplication(sys.argv)
    # 加载全局样式（支持 light/dark），默认尝试 light
    try:
        styles.apply_theme(app, theme='light')
    except Exception:
        pass

    # 在入口创建应用级 logger 并注入到各子窗口
    app_logger = get_logger('App')
    window = UnifiedLoginWindow(logger=app_logger)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()