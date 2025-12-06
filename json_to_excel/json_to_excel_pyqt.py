import sys
import pandas as pd
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QLabel, QFileDialog, QMessageBox, QFrame, QProgressDialog, QTextEdit, QGroupBox, QHBoxLayout, QSizePolicy)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from common.logger import get_logger, add_qt_text_widget, add_qt_signal
from sanbao_test.gui_utils import ControlButton, create_separator, LabeledFrame


class ConversionWorker(QThread):
    """转换工作线程，避免UI卡顿"""
    progress_updated = pyqtSignal(int)
    conversion_finished = pyqtSignal(bool, str)

    def __init__(self, json_path, excel_path, logger=None):
        super().__init__()
        self.json_path = json_path
        self.excel_path = excel_path
        # 使用外部传入的 logger（如果提供）以避免重复初始化
        self.logger = logger

    def run(self):
        try:
            if getattr(self, 'logger', None):
                self.logger.info(f"开始转换: {self.json_path} -> {self.excel_path}")
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
            if getattr(self, 'logger', None):
                self.logger.info(f"转换完成: {self.excel_path}")
            self.conversion_finished.emit(True, self.excel_path)
            
        except json.JSONDecodeError:
            if getattr(self, 'logger', None):
                self.logger.error("JSONDecodeError: JSON 文件格式不正确！")
            self.conversion_finished.emit(False, "JSON 文件格式不正确！")
        except Exception as e:
            if getattr(self, 'logger', None):
                self.logger.error(f"转换失败: {str(e)}")
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
    log_signal = pyqtSignal(str)

    def __init__(self, logger=None):
        super().__init__()
        self.json_file_path = ""
        # 先创建 UI，使得 self.log_text 可用
        self.initUI()
        # 如果外部传入 logger，则使用并附加 UI 文本控件；否则保持原有行为（UI 层创建 logger）
        try:
            # 连接信号到文本框（线程安全）
            try:
                self.log_signal.connect(lambda m: self._append_log(m))
            except Exception:
                pass
            # 优先使用外部传入的 logger（并将 QTextEdit 作为目标），否则在此处创建 logger 并附加 QTextEdit
            if logger is not None:
                # 使用外部传入的 logger，通过 signal 将日志发送到 UI，避免直接附加 QTextEdit 导致线程问题或重复输出
                self.logger = logger
                try:
                    add_qt_signal(self.logger, self.log_signal)
                except Exception:
                    pass
                self.logger.info("JSONToExcel UI 启动")
            else:
                # 兼容旧行为：在 UI 层创建 logger 并通过 signal 输出到 UI
                self.logger = get_logger('JSONToExcel')
                try:
                    add_qt_signal(self.logger, self.log_signal)
                except Exception:
                    pass
                self.logger.info("JSONToExcel UI 启动")
        except Exception:
            self.logger = None

    def _append_log(self, message: str):
        try:
            if getattr(self, 'log_text', None) is not None:
                self.log_text.append(message)
                try:
                    self.log_text.verticalScrollBar().setValue(
                        self.log_text.verticalScrollBar().maximum()
                    )
                except Exception:
                    pass
        except Exception:
            pass
    
    def initUI(self):
        self.setWindowTitle("JSON 转 Excel 工具")
        self.setGeometry(120, 120, 700, 520)

        # 主布局，借鉴三保界面规范
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # 标题与描述（使用与三保一致的 objectName）
        title_label = QLabel("JSON 转 Excel 工具")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setObjectName('title')

        desc_label = QLabel("将JSON数据转换为Excel表格，支持嵌套结构解析")
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setObjectName('subtitle')

        main_layout.addWidget(title_label)
        main_layout.addWidget(desc_label)

        # 文件选择（使用 LabeledFrame 或 GroupBox风格）
        file_frame = LabeledFrame(self, "选择导入JSON文件")
        file_layout = file_frame.content_layout
        file_layout.setSpacing(8)

        # 使用 ControlButton 与统一样式
        self.select_btn = ControlButton(self, "1. 浏览..", width=14)

        # 在路径前显示 "文件选择:" 字眼以增强语义
        self.file_label = QLabel("文件选择: 未选择文件")
        self.file_label.setWordWrap(True)
        self.file_label.setObjectName('file_placeholder')

        # 将文件路径标签与选择按钮放在同一行，路径文本右对齐并占据可用空间，按钮固定在最右侧
        file_row = QHBoxLayout()
        file_row.setSpacing(8)
        # 让 label 在水平方向上可扩展，并将文本右对齐
        self.file_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        file_row.addWidget(self.file_label)
        file_row.addWidget(self.select_btn)
        file_layout.addLayout(file_row)
        main_layout.addWidget(file_frame)

        # 操作按钮行（横向，与三保风格一致）
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.convert_btn = ControlButton(self, "2. 转换为 Excel文件", width=16)
        self.convert_btn.setEnabled(False)
        self.convert_btn.setProperty('variant', 'secondary')
        btn_layout.addWidget(self.convert_btn)

        # 退出按钮，放在转换按钮右侧
        exit_btn = ControlButton(self, "退出", width=16)
        exit_btn.clicked.connect(self.close)
        btn_layout.addWidget(exit_btn)

        main_layout.addLayout(btn_layout)

        # 分隔线
        main_layout.addWidget(create_separator(self))

        # 日志区域（使用三保相同的分组形式）
        log_group = QGroupBox("日志信息")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setObjectName('log_text')
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

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
            # 显示完整路径并在前面加语义前缀
            self.file_label.setText(f"文件选择: {file_path}")
            # 切换到已选样式（由 common/style.qss 控制）
            self.file_label.setObjectName('file_selected')
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

        # 记录操作到统一日志
        try:
            if getattr(self, 'logger', None):
                self.logger.info(f"用户选择保存路径: {save_path}")
        except Exception:
            pass
        
        # 创建进度对话框
        progress = QProgressDialog("正在转换...", "取消", 0, 100, self)
        progress.setWindowTitle("处理中")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)
        
        # 创建并启动转换线程（将 UI 层的 logger 传入 worker）
        self.worker = ConversionWorker(self.json_file_path, save_path, logger=getattr(self, 'logger', None))
        self.worker.progress_updated.connect(progress.setValue)
        self.worker.conversion_finished.connect(self.on_conversion_finished)
        
        # 连接取消按钮信号
        progress.canceled.connect(self.worker.terminate)
        
        self.worker.start()
        progress.exec()

    def on_conversion_finished(self, success, message):
        # 记录并展示转换结果
        try:
            if getattr(self, 'logger', None):
                if success:
                    self.logger.info(f"转换成功: {message}")
                else:
                    self.logger.error(f"转换失败: {message}")
        except Exception:
            pass

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