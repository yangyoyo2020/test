import sys
import pandas as pd
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QLabel, QFileDialog, QMessageBox, QFrame, QProgressDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont


class ConversionWorker(QThread):
    """转换工作线程，避免UI卡顿"""
    progress_updated = pyqtSignal(int)
    conversion_finished = pyqtSignal(bool, str)

    def __init__(self, json_path, excel_path):
        super().__init__()
        self.json_path = json_path
        self.excel_path = excel_path

    def run(self):
        try:
            # 读取JSON文件 (10%)
            self.progress_updated.emit(10)
            with open(self.json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 标准化数据 (30%)
            self.progress_updated.emit(30)
            normalized_data = self.normalize_json(data)
            
            # 转换为DataFrame (50%)
            self.progress_updated.emit(50)
            if isinstance(normalized_data, list):
                df = pd.DataFrame(normalized_data)
            elif isinstance(normalized_data, dict):
                df = pd.DataFrame([normalized_data])
            else:
                self.conversion_finished.emit(False, "不支持的 JSON 数据格式！")
                return
            
            # 保存为Excel (80%)
            self.progress_updated.emit(80)
            df.to_excel(self.excel_path, index=False)
            
            # 完成 (100%)
            self.progress_updated.emit(100)
            self.conversion_finished.emit(True, self.excel_path)
            
        except json.JSONDecodeError:
            self.conversion_finished.emit(False, "JSON 文件格式不正确！")
        except Exception as e:
            self.conversion_finished.emit(False, f"转换失败: {str(e)}")

    @staticmethod
    def flatten_dict(d, parent_key='', sep='_'):
        """将嵌套字典展平为单层字典"""
        items = []
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(ConversionWorker.flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                # 处理列表类型数据
                if v and isinstance(v[0], dict):
                    # 列表元素是字典，分别处理
                    for i, item in enumerate(v):
                        if isinstance(item, dict):
                            items.extend(ConversionWorker.flatten_dict(item, f"{new_key}{sep}{i}", sep=sep).items())
                        else:
                            items.append((f"{new_key}{sep}{i}", item))
                else:
                    # 列表元素不是字典，转换为字符串
                    items.append((new_key, ', '.join(map(str, v)) if v else ''))
            else:
                items.append((new_key, v))
        return dict(items)

    @staticmethod
    def normalize_json(data):
        """标准化JSON数据以便于转换为DataFrame"""
        if isinstance(data, list):
            normalized_data = []
            for item in data:
                if isinstance(item, dict):
                    normalized_data.append(ConversionWorker.flatten_dict(item))
                else:
                    normalized_data.append(item)
            return normalized_data
        elif isinstance(data, dict):
            return ConversionWorker.flatten_dict(data)
        else:
            return data


class JSONToExcelConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.json_file_path = ""
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("JSON 转 Excel 工具")
        self.setGeometry(300, 300, 550, 320)
        # self.setMinimumSize(500, 300)  # 设置最小窗口尺寸
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
                font-family: "Microsoft YaHei", sans-serif;
            }
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 12px 24px;
                text-align: center;
                font-size: 16px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
        """)
        
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # 标题
        title_label = QLabel("JSON 转 Excel 工具")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        
        # 描述标签
        desc_label = QLabel("将JSON数据转换为Excel表格，支持嵌套结构解析")
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")
        
        # 文件选择区域
        file_frame = QFrame()
        file_frame.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        file_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        
        file_layout = QVBoxLayout()
        file_layout.setSpacing(15)
        
        # 选择文件按钮
        self.select_btn = QPushButton("📁 选择 JSON 文件")
        self.select_btn.setMinimumHeight(50)
        
        # 显示文件路径
        self.file_label = QLabel("未选择文件")
        self.file_label.setWordWrap(True)
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setStyleSheet("""
            QLabel {
                color: #95a5a6;
                background-color: #ecf0f1;
                padding: 15px;
                border-radius: 8px;
                font-size: 13px;
            }
        """)
        
        file_layout.addWidget(self.select_btn)
        file_layout.addWidget(self.file_label)
        file_frame.setLayout(file_layout)
        
        # 转换按钮
        self.convert_btn = QPushButton("🔄 转换为 Excel")
        self.convert_btn.setMinimumHeight(50)
        self.convert_btn.setEnabled(False)  # 初始禁用
        
        # 添加控件到主布局
        main_layout.addWidget(title_label)
        main_layout.addWidget(desc_label)
        main_layout.addWidget(file_frame, 1)  # 让文件区域可伸缩
        main_layout.addWidget(self.convert_btn)
        
        self.setLayout(main_layout)
        
        # 连接信号和槽
        self.select_btn.clicked.connect(self.select_json_file)
        self.convert_btn.clicked.connect(self.convert_json_to_excel)

    def select_json_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "请选择 JSON 文件",
            str(Path.home()),  # 默认打开用户主目录
            "JSON 文件 (*.json);;所有文件 (*)"
        )
        
        if file_path:
            self.json_file_path = file_path
            # 显示完整路径但自动换行
            self.file_label.setText(file_path)
            self.file_label.setStyleSheet("""
                QLabel {
                    color: #27ae60;
                    background-color: #d5f5e3;
                    padding: 15px;
                    border-radius: 8px;
                    font-size: 13px;
                }
            """)
            self.convert_btn.setEnabled(True)  # 启用转换按钮

    def convert_json_to_excel(self):
        if not self.json_file_path or not Path(self.json_file_path).exists():
            QMessageBox.critical(self, "错误", "请先选择一个有效的 JSON 文件！")
            return
        
        # 获取默认保存文件名（与JSON文件同名）
        default_filename = Path(self.json_file_path).stem + ".xlsx"
        default_dir = str(Path(self.json_file_path).parent)
        
        # 选择保存 Excel 的路径
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存为 Excel 文件",
            str(Path(default_dir) / default_filename),
            "Excel 文件 (*.xlsx);;所有文件 (*)"
        )
        
        if not save_path:
            return  # 用户取消保存
        
        # 确保文件扩展名正确
        if not save_path.endswith('.xlsx'):
            save_path += '.xlsx'
        
        # 创建进度对话框
        progress = QProgressDialog("正在转换...", "取消", 0, 100, self)
        progress.setWindowTitle("处理中")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)
        
        # 创建并启动转换线程
        self.worker = ConversionWorker(self.json_file_path, save_path)
        self.worker.progress_updated.connect(progress.setValue)
        self.worker.conversion_finished.connect(self.on_conversion_finished)
        
        # 连接取消按钮信号
        progress.canceled.connect(self.worker.terminate)
        
        self.worker.start()
        progress.exec()

    def on_conversion_finished(self, success, message):
        if success:
            QMessageBox.information(
                self, 
                "成功", 
                f"转换完成！\n文件已保存至:\n{message}",
                QMessageBox.StandardButton.Ok
            )
        else:
            QMessageBox.critical(
                self, 
                "错误", 
                message,
                QMessageBox.StandardButton.Ok
            )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # 设置应用程序字体，确保中文显示正常
    font = QFont("Microsoft YaHei")
    app.setFont(font)
    
    converter = JSONToExcelConverter()
    converter.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()