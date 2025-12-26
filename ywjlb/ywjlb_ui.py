"""
运维记录表（YWJLB）PyQt6 用户界面模块
支持三种包类型的Excel转Word文档处理
"""

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QFileDialog, QMessageBox, QLabel, QTextEdit,
    QGroupBox, QPushButton, QComboBox, QProgressBar
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont
import pandas as pd
import os
import logging
from .ywjlb_unified import PackageType, process_excel_file

# 获取模块日志记录器
logger = logging.getLogger(__name__)


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


class ProcessWorker(QThread):
    """文件处理工作线程，避免UI卡顿"""
    progress = pyqtSignal(str)  # 发送进度信息到UI
    finished = pyqtSignal(tuple)  # 发送处理结果 (成功数, 失败数)
    error = pyqtSignal(str)  # 发送错误信息


class YWJLBAnalyzer(QMainWindow):
    """运维记录簿统一处理工具主类"""
    log_signal = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.df = None
        self.input_file_path = ""
        self.output_folder_path = ""  # 默认输出文件夹
        self.selected_package_type = PackageType.GJGC
        self.worker = None
        self.progress_dialog = None
        
        self._setup_ui()
        self._setup_logging()
    
    def _setup_ui(self):
        """设置用户界面"""
        self.setWindowTitle("运维记录表处理工具")
        self.setGeometry(100, 100, 900, 500)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)
        
        # 1. 文件选择区域
        main_layout.addLayout(self._create_file_frame())
        
        # 2. 包类型选择区域
        main_layout.addLayout(self._create_package_type_frame())
        
        # 3. 预览表区域（已从界面移除；仍保留内部创建以保持逻辑兼容）
        # 调用创建函数以初始化相关属性，但不要将其布局加入主布局，从而在 UI 上隐藏该区域
        try:
            self._create_preview_frame()
        except Exception:
            pass
        
        # 4. 进度条区域
        main_layout.addLayout(self._create_progress_frame())
        
        # 5. 日志显示区域
        main_layout.addLayout(self._create_log_frame())
        
        # 6. 操作按钮区域
        main_layout.addLayout(self._create_button_frame())
        
        main_layout.addStretch()
    
    def _setup_logging(self):
        """设置日志处理器，将ywjlb_unified模块的日志重定向到UI"""
        # 创建UI日志处理器
        self.ui_handler = UILogHandler(self)
        self.ui_handler.setLevel(logging.DEBUG)  # 捕获所有级别的日志
        
        # 获取根logger并添加处理器（因为ywjlb_unified.py直接使用logging模块）
        root_logger = logging.getLogger()
        root_logger.addHandler(self.ui_handler)
        # 保存原始级别
        self.original_root_level = root_logger.level
        # 确保根logger的级别允许DEBUG消息通过
        if root_logger.level > logging.DEBUG:
            root_logger.setLevel(logging.DEBUG)
        
        # 连接信号到日志显示方法
        self.log_signal.connect(self._append_log)
        
        # 调试信息
        print(f"日志处理器已设置，根logger级别: {root_logger.level}")
        print(f"处理器数量: {len(root_logger.handlers)}")
    
    def closeEvent(self, event):
        """窗口关闭事件，清理日志处理器"""
        try:
            # 移除UI日志处理器
            root_logger = logging.getLogger()
            root_logger.removeHandler(self.ui_handler)
            # 恢复原始级别
            root_logger.setLevel(self.original_root_level)
        except Exception:
            pass
        super().closeEvent(event)
    
    def _create_file_frame(self):
        """创建文件选择区域"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        frame = QGroupBox("1. 选择输入文件和输出文件夹")
        frame.setFont(self._get_bold_font())
        
        # 输入文件路径显示和选择按钮
        input_layout = QHBoxLayout()
        input_layout.setSpacing(8)
        
        self.file_path_input = QLineEdit()
        self.file_path_input.setPlaceholderText("请选择Excel文件...")
        self.file_path_input.setReadOnly(True)
        input_layout.addWidget(QLabel("输入文件:"), 0)
        input_layout.addWidget(self.file_path_input, 1)
        
        browse_btn = QPushButton("浏览..")
        browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(self._choose_input_file)
        input_layout.addWidget(browse_btn)
        
        # 输出文件夹路径显示和选择按钮
        output_layout = QHBoxLayout()
        output_layout.setSpacing(8)
        
        self.output_folder_input = QLineEdit()
        self.output_folder_input.setPlaceholderText("选择Word文档输出文件夹...")
        self.output_folder_input.setText(self.output_folder_path)
        self.output_folder_input.setReadOnly(True)
        output_layout.addWidget(QLabel("输出文件夹:"), 0)
        output_layout.addWidget(self.output_folder_input, 1)
        
        output_browse_btn = QPushButton("浏览..")
        output_browse_btn.setFixedWidth(80)
        output_browse_btn.clicked.connect(self._choose_output_folder)
        output_layout.addWidget(output_browse_btn)
        
        # 组合输入输出到frame中
        file_layout = QVBoxLayout()
        file_layout.addLayout(input_layout)
        file_layout.addLayout(output_layout)
        
        frame.setLayout(file_layout)
        layout.addWidget(frame)
        
        return layout
    
    def _create_package_type_frame(self):
        """创建包类型选择区域"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        frame = QGroupBox("2. 选择包类型")
        frame.setFont(self._get_bold_font())
        
        type_layout = QHBoxLayout()
        type_layout.setSpacing(10)
        
        label = QLabel("包类型:")
        type_layout.addWidget(label)
        
        self.package_type_combo = QComboBox()
        self.package_type_combo.addItem(PackageType.GJGC.value, PackageType.GJGC)
        self.package_type_combo.addItem(PackageType.QTXM.value, PackageType.QTXM)
        self.package_type_combo.addItem(PackageType.PKG02.value, PackageType.PKG02)
        self.package_type_combo.currentIndexChanged.connect(self._on_package_type_changed)
        type_layout.addWidget(self.package_type_combo)
        
        # 添加包类型说明
        info_label = QLabel("（包含不同的数据字段）")
        info_label.setStyleSheet("color: gray; font-style: italic;")
        type_layout.addWidget(info_label)
        
        type_layout.addStretch()

        # 在包类型行的最右侧添加操作按钮（保持原有样式与行为）
        load_btn = QPushButton("3.加载数据")
        load_btn.setFixedWidth(100)
        load_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        load_btn.clicked.connect(self._load_data)
        type_layout.addWidget(load_btn)

        process_btn = QPushButton("4.处理并导出")
        process_btn.setFixedWidth(120)
        process_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:pressed {
                background-color: #0a68c8;
            }
        """)
        process_btn.clicked.connect(self._process_file)
        type_layout.addWidget(process_btn)

        exit_btn = QPushButton("退出")
        exit_btn.setFixedWidth(80)
        exit_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:pressed {
                background-color: #c4130a;
            }
        """)
        exit_btn.clicked.connect(self.close)
        type_layout.addWidget(exit_btn)

        frame.setLayout(type_layout)
        layout.addWidget(frame)

        return layout
    
    def _create_preview_frame(self):
        """创建数据预览区域"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        frame = QGroupBox("3. 数据预览（前50行）")
        frame.setFont(self._get_bold_font())
        
        # 创建文本编辑框用于显示数据预览
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(150)
        self.preview_text.setFont(self._get_monospace_font())
        self.preview_text.setPlaceholderText("加载文件后，数据预览将显示在这里...")
        
        frame.setLayout(QVBoxLayout())
        frame.layout().addWidget(self.preview_text)
        
        layout.addWidget(frame)
        return layout
    
    def _create_progress_frame(self):
        """创建进度条区域"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        frame = QGroupBox("转换进度")
        frame.setFont(self._get_bold_font())
        
        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪")
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        frame.setLayout(QVBoxLayout())
        frame.layout().addWidget(self.progress_bar)
        
        layout.addWidget(frame)
        return layout
    
    def _create_log_frame(self):
        """创建日志显示区域"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        frame = QGroupBox("日志信息")
        frame.setFont(self._get_bold_font())
        
        # 创建日志文本框
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setFont(self._get_monospace_font())
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #cccccc;
            }
        """)
        
        frame.setLayout(QVBoxLayout())
        frame.layout().addWidget(self.log_text)
        
        # 添加日志操作按钮
        ops_layout = QHBoxLayout()
        ops_layout.setSpacing(8)
        
        clear_btn = QPushButton("清空日志")
        clear_btn.setFixedWidth(80)
        clear_btn.clicked.connect(lambda: self.log_text.clear())
        ops_layout.addWidget(clear_btn)
        
        export_btn = QPushButton("导出日志")
        export_btn.setFixedWidth(80)
        export_btn.clicked.connect(self._export_log)
        ops_layout.addWidget(export_btn)
        
        ops_layout.addStretch()
        frame.layout().addLayout(ops_layout)
        
        layout.addWidget(frame)
        return layout
    
    def _create_button_frame(self):
        """创建底部操作按钮区域"""
        layout = QHBoxLayout()
        layout.setSpacing(10)
        # 该区域已将功能按钮移至包类型行右侧，底部保留空布局以保持布局结构
        layout.addStretch()
        return layout
    
    def _get_bold_font(self):
        """获取粗体字体"""
        font = QFont()
        font.setBold(True)
        return font
    
    def _get_monospace_font(self):
        """获取等宽字体"""
        font = QFont("Consolas" if self._is_windows() else "Courier New")
        font.setPointSize(9)
        return font
    
    @staticmethod
    def _is_windows():
        """检查是否是Windows系统"""
        return os.name == 'nt'
    
    def _choose_input_file(self):
        """弹出文件选择对话框"""
        path, _ = QFileDialog.getOpenFileName(
            self,
            '选择运维记录表.xlsx',
            '',
            'Excel 文件 (*.xlsx *.xls);;All files (*.*)'
        )
        if path:
            self.input_file_path = path
            self.file_path_input.setText(path)
            self._append_log(f"✓ 已选择文件: {os.path.basename(path)}")
    
    def _choose_output_folder(self):
        """弹出文件夹选择对话框用于选择输出路径"""
        folder = QFileDialog.getExistingDirectory(
            self,
            '选择输出文件夹',
            self.output_folder_path or os.path.expanduser("~")
        )
        if folder:
            self.output_folder_path = folder
            self.output_folder_input.setText(folder)
            self._append_log(f"✓ 已选择输出文件夹: {folder}")
    
    def _on_package_type_changed(self):
        """包类型选择变化事件"""
        self.selected_package_type = self.package_type_combo.currentData()
        self._append_log(f"✓ 已选择包类型: {self.selected_package_type.value}")
    
    def _load_data(self):
        """加载Excel数据文件"""
        if not self.input_file_path or not os.path.exists(self.input_file_path):
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
        
        try:
            self._append_log("正在加载数据...")
            
            # 读取Excel文件
            self.df = pd.read_excel(self.input_file_path, sheet_name=None)
            
            # 统计信息
            total_rows = sum(len(df) for df in self.df.values())
            self._append_log("✓ 成功加载数据")
            self._append_log(f"  - 工作表数: {len(self.df)}")
            self._append_log(f"  - 数据行数: {total_rows}")
            
            # 显示预览
            self._show_preview()
            
        except Exception as e:
            self._append_log(f"✗ 加载数据失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"加载数据失败:\n{str(e)}")
    
    def _show_preview(self):
        """显示数据预览"""
        try:
            preview_text = "数据概览:\n" + "=" * 50 + "\n\n"
            
            for sheet_name, df in self.df.items():
                preview_text += f"【工作表】{sheet_name}\n"
                preview_text += f"  行数: {len(df)}, 列数: {len(df.columns)}\n"
                preview_text += f"  列名: {', '.join(df.columns[:5])}"
                if len(df.columns) > 5:
                    preview_text += ", ..."
                preview_text += "\n\n"
            
            self.preview_text.setText(preview_text)
            
        except Exception as e:
            self._append_log(f"✗ 显示预览失败: {str(e)}")
    
    def _process_file(self):
        """处理文件并导出Word文档"""
        if not self.input_file_path or not os.path.exists(self.input_file_path):
            QMessageBox.warning(self, "警告", "请先选择并加载Excel文件")
            return
        
        if self.df is None:
            QMessageBox.warning(self, "警告", "请先加载数据")
            return
        
        if not self.output_folder_path or not self.output_folder_path.strip():
            QMessageBox.warning(self, "警告", "请先选择输出文件夹\n\n点击\"输出文件夹\"旁的\"浏览\"按钮选择保存位置")
            return
        
        try:
            # 重置进度条
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("开始处理...")
            
            self._append_log("=" * 50)
            self._append_log("开始处理数据并导出Word文档...")
            self._append_log(f"包类型: {self.selected_package_type.value}")
            self._append_log(f"输出文件夹: {self.output_folder_path}")
            
            # 定义进度回调函数
            def progress_callback(current, total):
                if total > 0:
                    percentage = int((current / total) * 100)
                    self.progress_bar.setValue(percentage)
                    self.progress_bar.setFormat(f"处理进度: {current}/{total} ({percentage}%)")
            
            # 调用处理函数，传入输出文件夹路径和进度回调
            success_count, fail_count = process_excel_file(
                self.input_file_path,
                self.selected_package_type,
                self.output_folder_path,
                progress_callback
            )
            
            # 设置进度条为完成状态
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("处理完成")
            
            self._append_log("=" * 50)
            self._append_log("✓ 处理完成!")
            self._append_log(f"  - 成功: {success_count} 个文档")
            self._append_log(f"  - 失败: {fail_count} 个文档")
            
            # 显示完成对话框
            QMessageBox.information(
                self,
                "处理完成",
                f"已成功生成 {success_count} 个Word文档\n失败 {fail_count} 个\n\n"
                f"文档保存在: {self.output_folder_path}\n"
                f"详细日志保存在: 日志文件.log"
            )
            
        except Exception as e:
            # 处理失败时重置进度条
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("处理失败")
            self._append_log(f"✗ 处理失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"处理失败:\n{str(e)}")
    
    def _append_log(self, message: str):
        """向日志窗口追加消息"""
        try:
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
        except Exception:
            pass
    
    def _export_log(self):
        """导出日志到文件"""
        try:
            path, _ = QFileDialog.getSaveFileName(
                self,
                '保存日志',
                'ywjlb_log.txt',
                'Text Files (*.txt);;All Files (*.*)'
            )
            
            if path:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                
                QMessageBox.information(self, "成功", f"日志已导出到:\n{path}")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出日志失败:\n{str(e)}")


def main():
    """应用程序入口"""
    import sys
    app = QApplication(sys.argv)
    
    # 设置应用样式
    app.setStyle('Fusion')
    
    window = YWJLBAnalyzer()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
