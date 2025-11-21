"""GUI工具类"""
from PyQt6.QtWidgets import (
    QScrollArea, QWidget, QVBoxLayout, QFrame, QLabel, QPushButton, 
    QHBoxLayout, QSizePolicy, QSpacerItem
)
from PyQt6.QtCore import Qt


class ScrollableFrame(QScrollArea):
    """可滚动的Frame组件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # 创建可滚动的框架
        self.scrollable_frame = QWidget()
        self.scrollable_frame.setObjectName("scrollable_frame")
        
        # 设置垂直布局
        self.layout = QVBoxLayout(self.scrollable_frame)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.layout.setSpacing(2)
        self.layout.setContentsMargins(5, 5, 5, 5)
        
        # 配置滚动区域
        self.setWidget(self.scrollable_frame)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # 设置对象名称以便样式设置
        self.setObjectName("scroll_area")

    def get_layout(self):
        """获取内部布局"""
        return self.layout


class LabeledFrame(QFrame):
    """带标题的Frame组件，统一样式"""
    def __init__(self, parent, text):
        super().__init__(parent)
        
        # 设置为箱式布局
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setObjectName("labeled_frame")
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # 标题标签
        self.label = QLabel(text)
        self.label.setObjectName("frame_title")
        main_layout.addWidget(self.label)
        
        # 内容区域布局
        self.content_layout = QVBoxLayout()
        self.content_layout.setSpacing(5)
        main_layout.addLayout(self.content_layout)
        
        # 设置样式
        self.setStyleSheet("""
            #labeled_frame {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                background-color: white;
            }
            #frame_title {
                font-weight: bold;
                font-size: 11pt;
                padding: 0px;
                margin: 0px;
            }
        """)


class ControlButton(QPushButton):
    """统一样式的按钮组件"""
    def __init__(self, parent, text, width=15):
        super().__init__(text, parent)
        self.setMinimumWidth(width * 6)  # 减小按钮宽度
        self.setMinimumHeight(30)  # 设置按钮高度
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.setObjectName("control_button")
        
        # 设置样式，与json_to_excel工具保持一致但尺寸更小
        self.setStyleSheet("""
            #control_button {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                border-radius: 6px;
                font-weight: bold;
            }
            #control_button:hover {
                background-color: #45a049;
            }
            #control_button:pressed {
                background-color: #3d8b40;
            }
            #control_button:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)


def create_separator(parent):
    """创建分隔线"""
    separator = QFrame(parent)
    separator.setFrameShape(QFrame.Shape.HLine)
    separator.setFrameShadow(QFrame.Shadow.Sunken)
    separator.setObjectName("separator")
    
    # 设置样式
    separator.setStyleSheet("""
        #separator {
            border: none;
            border-top: 1px solid #E6E6E6;
            margin: 5px 0px;
        }
    """)
    
    return separator