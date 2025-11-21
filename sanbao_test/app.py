import os
from typing import Optional, Dict, List
from datetime import datetime
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLineEdit, QFileDialog, QMessageBox, QCheckBox, QPushButton,
    QLabel, QFrame, QTextEdit, QGroupBox, QProgressDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

from .constants import *
from .gui_utils import ScrollableFrame, ControlButton, create_separator

# 输出格式化列定义（用于 display_summary）
numeric_cols = [
    '调整预算数', '计划金额', '计划剩余金额', '支付申请金额',
    '在途金额', '实际支出金额', '在途+实际支出金额'
]
percent_cols = ['实际支出进度%', '在途+实际支出进度%']


class AnalysisWorker(QThread):
    """分析工作线程，避免UI卡顿"""
    progress = pyqtSignal(str)  # 发送进度信息到UI
    finished = pyqtSignal(object)  # 发送分析结果
    error = pyqtSignal(str)  # 发送错误信息
    
    def __init__(self, df, selected_units, selected_types):
        super().__init__()
        self.df = df
        self.selected_units = selected_units
        self.selected_types = selected_types
    
    def run(self):
        try:
            self.progress.emit("开始数据分析...")
            # 分析数据
            summary = analyze_expenditure(self.df, self.selected_units, self.selected_types)
            self.progress.emit(f"数据分析完成，生成 {len(summary)} 行汇总结果")
            self.finished.emit(summary)
        except Exception as e:
            self.error.emit(str(e))


class ExpenditureAnalyzer(QMainWindow):
    """三保支出进度分析工具主类"""
    
    def __init__(self):
        super().__init__()
        self.df = None  # 原始数据
        self.checkboxes_units = {}  # 预算单位复选框
        self.checkboxes_types = {}  # 三保标识复选框
        self.selected_units = []  # 已选预算单位
        self.selected_types = []  # 已选三保标识
        
        # 用于显示/保存选择的输入文件路径
        self.input_path_var = ""
        
        # 日志文本框
        self.log_text = None
        
        # 上一次选择的三保标识（用于避免重复日志）
        self._prev_selected_types = []
        
        # 控制是否抑制日志记录的标志
        self._suppress_log = False
        
        # 工作线程和进度对话框
        self.worker = None
        self.progress_dialog = None

        self._setup_ui()
        # 不在初始化时自动加载默认文件，改为由用户通过界面选择并加载
    
    def _setup_ui(self):
        """设置用户界面"""
        self.setWindowTitle(WINDOW_TITLE)
        self.setGeometry(100, 100, 800, 600)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # 创建主框架
        main_frame = QVBoxLayout()
        main_frame.setSpacing(5)
        
        # 创建上部控制区域
        control_frame = self._create_control_frame()
        main_frame.addLayout(control_frame)
        
        separator = create_separator(self)
        main_frame.addWidget(separator)
        
        # 创建选择区域（两列布局）
        selection_layout = QHBoxLayout()
        selection_layout.setSpacing(10)
        
        # 三保标识选择区域（左列）
        types_frame = self._create_types_frame()
        selection_layout.addLayout(types_frame, 1)
        
        # 预算单位选择区域（右列）
        units_frame = self._create_units_frame()
        selection_layout.addLayout(units_frame, 1)
        
        main_frame.addLayout(selection_layout)
        
        separator2 = create_separator(self)
        main_frame.addWidget(separator2)
        
        # 创建底部按钮区域
        button_layout = self._create_button_frame()
        main_frame.addLayout(button_layout)
        
        # 创建日志区域
        log_layout = self._create_log_frame()
        main_frame.addLayout(log_layout)
        
        main_layout.addLayout(main_frame)
        
        # 应用样式表
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F3F4F6;
            }
            QWidget {
                background-color: #F3F4F6;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                background-color: white;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                background-color: white;
                font-size: 11pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                font-family: Consolas, Monaco, monospace;
                font-size: 9pt;
            }
        """)
    
    def _create_control_frame(self):
        """创建顶部控制区域"""
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 标题框架
        frame = QGroupBox("导入文件")
        
        # 输入文件路径显示和选择按钮
        input_layout = QHBoxLayout()
        input_layout.setSpacing(5)
        
        self.entry = QLineEdit()
        self.entry.setPlaceholderText("请选择输入文件...")
        self.entry.setText(self.input_path_var)
        input_layout.addWidget(self.entry)
        
        select_btn = ControlButton(self, "1.选择文件", width=12)
        select_btn.clicked.connect(self._choose_input_file)
        input_layout.addWidget(select_btn)
        
        # 添加"加载数据"按钮（由用户手动触发加载）
        load_btn = ControlButton(self, "2.加载数据", width=12)
        load_btn.clicked.connect(lambda: self._load_data(self.entry.text().strip() or None))
        input_layout.addWidget(load_btn)
        
        frame.setLayout(input_layout)
        layout.addWidget(frame)
        
        return layout

    def _create_log_frame(self):
        """创建日志显示区域"""
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        frame = QGroupBox("日志信息")
        
        # 创建文本框用于显示日志
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setObjectName("log_text")
        self.log_text.setStyleSheet("""
            #log_text {
                background-color: #FFFFFF;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                font-family: Consolas, Monaco, monospace;
                font-size: 9pt;
            }
        """)
        
        frame.setLayout(QVBoxLayout())
        frame.layout().addWidget(self.log_text)
        layout.addWidget(frame)
        
        # 添加初始日志
        self._log_message("应用程序已启动")
        
        return layout
    
    def _log_message(self, message):
        """向日志窗口添加消息"""
        if self.log_text:
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}"
            self.log_text.append(formatted_message)
            # 自动滚动到底部
            self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
    
    def _choose_input_file(self):
        """弹出文件选择对话框，让用户选择输入的 Excel 文件"""
        path, _ = QFileDialog.getOpenFileName(
            self,
            '选择导出数据.xlsx',
            '',
            'Excel 文件 (*.xlsx *.xls);;All files (*.*)'
        )
        if path:
            self.entry.setText(path)
            self._log_message(f"已选择文件: {path}")
    
    def _create_units_frame(self):
        """创建预算单位选择区域"""
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        frame = QGroupBox("4.预算单位（可选）")
        
        # 添加全选/取消按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(5)
        
        select_all = ControlButton(self, "全选", width=8)
        select_all.clicked.connect(lambda: self._toggle_all_units(True))
        btn_layout.addWidget(select_all)
        
        select_none = ControlButton(self, "取消全选", width=8)
        select_none.clicked.connect(lambda: self._toggle_all_units(False))
        btn_layout.addWidget(select_none)
        
        # 添加标签显示"*不选择按三保标识汇总"
        label = QLabel("*不选择按三保标识汇总")
        label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-style: italic;
                margin-left: 10px;
            }
        """)
        btn_layout.addWidget(label)
        
        btn_layout.addStretch()
        frame.setLayout(QVBoxLayout())
        frame.layout().addLayout(btn_layout)
        
        # 创建可滚动的复选框区域
        self.scroll_units = ScrollableFrame(self)
        frame.layout().addWidget(self.scroll_units)
        
        self.units_inner_layout = self.scroll_units.get_layout()
        layout.addWidget(frame)
        
        return layout
    
    def _create_types_frame(self):
        """创建三保标识选择区域"""
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        frame = QGroupBox("3.三保标识（必选）")
        
        # 添加全选/取消按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(5)
        
        select_all = ControlButton(self, "全选", width=8)
        select_all.clicked.connect(lambda: self._toggle_all_types(True))
        btn_layout.addWidget(select_all)
        
        select_none = ControlButton(self, "取消全选", width=8)
        select_none.clicked.connect(lambda: self._toggle_all_types(False))
        btn_layout.addWidget(select_none)
        
        btn_layout.addStretch()
        frame.setLayout(QVBoxLayout())
        frame.layout().addLayout(btn_layout)
        
        # 创建可滚动的复选框区域
        self.scroll_types = ScrollableFrame(self)
        frame.layout().addWidget(self.scroll_types)
        
        self.types_inner_layout = self.scroll_types.get_layout()
        layout.addWidget(frame)
        
        return layout
    
    def _create_button_frame(self):
        """创建底部按钮区域"""
        layout = QHBoxLayout()
        layout.addStretch()
        
        analyze_btn = ControlButton(self, "5.分析并导出", width=BUTTON_WIDTH)
        analyze_btn.clicked.connect(self._run_analysis)
        layout.addWidget(analyze_btn)
        
        # 添加退出按钮
        exit_btn = ControlButton(self, "退出", width=BUTTON_WIDTH)
        exit_btn.clicked.connect(self.close)
        layout.addWidget(exit_btn)
        
        return layout
    
    def _load_data(self, path: Optional[str] = None):
        """加载数据并初始化复选框

        接受可选的 path 参数（通过 lambda 传入），若未提供则使用默认行为。
        """
        # 优先使用传入的 path，否则使用界面上的 input_path_var，如果都没有则传 None（load_exported_data 会使用默认路径）
        chosen = path if path else (self.entry.text().strip() if self.entry.text().strip() else None)
        try:
            self.df = load_exported_data(chosen)
            self._log_message(f"成功加载数据文件: {chosen or '默认路径'}，共 {len(self.df)} 行记录")

            # 更新预算单位复选框：排除预算单位为 0 的项（同时处理数值 0 和字符串 '0'/'0.0'）
            units_series = self.df['预算单位']
            # 使用 constants.nonzero_unit_mask 来统一判断哪些单位不是 0
            mask_not_zero = nonzero_unit_mask(units_series)
            units = sorted(units_series[mask_not_zero].unique())
            self._create_unit_checkboxes(units)

            # 更新三保标识复选框
            types = sorted(self.df['三保标识'].unique())
            self._create_type_checkboxes(types)
            
            self._log_message(f"已更新三保标识 ({len(types)} 个) 和预算单位 ({len(units)} 个) 复选框")

        except Exception as e:
            self._log_message(f"加载数据失败: {str(e)}")
            QMessageBox.critical(self, "错误", str(e))
    
    def _create_unit_checkboxes(self, units):
        """创建预算单位复选框（两列布局）"""
        # 清除现有复选框
        for i in reversed(range(self.units_inner_layout.count())): 
            widget = self.units_inner_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.checkboxes_units.clear()
        
        # 创建新的复选框（单列垂直显示）
        for unit in units:
            cb = QCheckBox(unit)
            cb.setStyleSheet("""
                QCheckBox {
                    spacing: 5px;
                    padding: 2px;
                    border-radius: 4px;
                }
                QCheckBox::indicator {
                    width: 14px;
                    height: 14px;
                }
                QCheckBox::indicator:unchecked {
                    border: 2px solid #666666;
                    background-color: white;
                }
                QCheckBox::indicator:checked {
                    border: 2px solid #4CAF50;
                    background-color: #4CAF50;
                }
                QCheckBox::indicator:checked::after {
                    content: "✓";
                    color: black;
                    font-weight: bold;
                    position: absolute;
                    left: 1px;
                    top: -3px;
                    font-size: 14px;
                    line-height: 14px;
                }
                QCheckBox:hover {
                    background-color: #f0f0f0;
                }
                QCheckBox:checked {
                    background-color: #e8f5e9;
                    font-weight: bold;
                }
            """)
            cb.stateChanged.connect(self._update_selected_units)
            self.checkboxes_units[unit] = cb
            # 单列垂直排列，左对齐
            self.units_inner_layout.addWidget(cb)
            
        # 记录日志
        # self._log_message(f"已创建 {len(units)} 个预算单位复选框")
    
    def _create_type_checkboxes(self, types):
        """创建三保标识复选框（单列布局）"""
        # 清除现有复选框
        for i in reversed(range(self.types_inner_layout.count())):
            widget = self.types_inner_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.checkboxes_types.clear()
        
        # 创建新的复选框
        filtered_types = [t for t in types if not str(t).startswith('[000]')]
        for type_name in filtered_types:
            cb = QCheckBox(str(type_name))
            cb.setStyleSheet("""
                QCheckBox {
                    spacing: 5px;
                    padding: 2px;
                    border-radius: 4px;
                }
                QCheckBox::indicator {
                    width: 14px;
                    height: 14px;
                }
                QCheckBox::indicator:unchecked {
                    border: 2px solid #666666;
                    background-color: white;
                }
                QCheckBox::indicator:checked {
                    border: 2px solid #4CAF50;
                    background-color: #4CAF50;
                }
                QCheckBox::indicator:checked::after {
                    content: "✓";
                    color: black;
                    font-weight: bold;
                    position: absolute;
                    left: 1px;
                    top: -3px;
                    font-size: 14px;
                    line-height: 14px;
                }
                QCheckBox:hover {
                    background-color: #f0f0f0;
                }
                QCheckBox:checked {
                    background-color: #e8f5e9;
                    font-weight: bold;
                }
            """)
            cb.stateChanged.connect(self._update_selected_types)
            self.checkboxes_types[type_name] = cb
            self.types_inner_layout.addWidget(cb)
            
        # 记录日志
        # self._log_message(f"已创建 {len(filtered_types)} 个三保标识复选框 (过滤了[000]开头的项)")
    
    def _toggle_all_units(self, state: bool):
        """切换所有预算单位的选择状态"""
        # 设置标志以抑制更新日志
        self._suppress_log = True
        try:
            for cb in self.checkboxes_units.values():
                cb.setChecked(state)
            
            # 手动更新选择列表
            self.selected_units = list(self.checkboxes_units.keys()) if state else []
            
            # 记录全选/取消全选操作
            if state:
                self._log_message(f"预算单位全选: 共 {len(self.checkboxes_units)} 个单位")
            else:
                self._log_message(f"预算单位取消全选: 共 {len(self.checkboxes_units)} 个单位")
        finally:
            # 恢复日志记录
            self._suppress_log = False
    
    def _toggle_all_types(self, state: bool):
        """切换所有三保标识的选择状态"""
        # 设置标志以抑制更新日志
        self._suppress_log = True
        try:
            for cb in self.checkboxes_types.values():
                cb.setChecked(state)
            
            # 手动更新选择列表
            self.selected_types = list(self.checkboxes_types.keys()) if state else []
            
            # 记录全选/取消全选操作
            if state:
                self._log_message(f"三保标识全选: 共 {len(self.checkboxes_types)} 个标识")
            else:
                self._log_message(f"三保标识取消全选: 共 {len(self.checkboxes_types)} 个标识")
        finally:
            # 恢复日志记录
            self._suppress_log = False
    
    def _update_selected_units(self):
        """更新已选择的预算单位列表"""
        old_count = len(self.selected_units)
        self.selected_units = [
            unit for unit, cb in self.checkboxes_units.items()
            if cb.isChecked()
        ]
        new_count = len(self.selected_units)
        
        # 只有当选择数量发生变化且不是全选/取消全选操作时才记录日志
        if old_count != new_count and not self._suppress_log:
            self._log_message(f"预算单位: 当前已选择 {new_count} 个单位")
    
    def _update_selected_types(self):
        """更新已选择的三保标识列表"""
        old_count = len(self.selected_types)
        self.selected_types = [
            type_name for type_name, cb in self.checkboxes_types.items()
            if cb.isChecked()
        ]
        new_count = len(self.selected_types)
        
        # 只有当选择数量发生变化且不是全选/取消全选操作时才记录日志
        if old_count != new_count and not self._suppress_log:
            self._log_message(f"三保标识: 当前已选择 {new_count} 个标识")
    
    def _run_analysis(self):
        """运行分析并导出结果"""
        # 三保标识为必选项，预算单位为可选项
        if not self.selected_types:
            self._log_message("分析失败: 未选择任何三保标识")
            QMessageBox.critical(self, "错误", "请至少选择一个三保标识")
            return
        
        # 创建进度对话框
        self.progress_dialog = QProgressDialog("正在分析数据...", "取消", 0, 0, self)
        self.progress_dialog.setWindowTitle("处理中")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        
        # 创建并启动分析线程
        self.worker = AnalysisWorker(self.df, self.selected_units, self.selected_types)
        self.worker.progress.connect(self._log_message)
        self.worker.finished.connect(self._analysis_completed)
        self.worker.error.connect(self._analysis_failed)
        self.progress_dialog.canceled.connect(self.worker.terminate)
        
        self.worker.start()
        self.progress_dialog.show()
    
    def _analysis_completed(self, summary):
        """分析完成回调"""
        # 关闭进度对话框
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        
        # 清理工作线程引用
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
            
        try:
            # 显示分析结果
            self._log_message("正在生成分析结果...")
            display_summary(summary)
            
            # 选择保存位置并导出
            output_path = self._choose_output()
            if output_path:
                self._log_message("正在保存分析结果...")
                save_to_excel(summary, output_path)
                self._log_message(f"分析结果已保存: {output_path}")
                QMessageBox.information(self, "成功", f"分析结果已保存到:\n{output_path}")
            else:
                self._log_message("用户取消了文件保存操作")
        except Exception as e:
            error_msg = f"分析完成后处理出错: {str(e)}"
            self._log_message(error_msg)
            QMessageBox.critical(self, "错误", f"{error_msg}\n\n详细信息: {str(e)}")
    
    def _analysis_failed(self, error_msg):
        """分析失败回调"""
        # 关闭进度对话框
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
            
        # 清理工作线程引用
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
            
        self._log_message(f"分析过程出错: {error_msg}")
        QMessageBox.critical(self, "错误", f"分析过程中发生错误:\n{error_msg}")
    
    def _choose_output(self) -> Optional[str]:
        """选择输出文件位置"""
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_name = f"三保支出进度_{ts}.xlsx"
        
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存文件",
            default_name,
            "Excel files (*.xlsx);;All files (*.*)"
        )
        if path:
            self._log_message(f"已选择输出文件路径: {path}")
            return path
        return None if path else None


def load_exported_data(path: Optional[str] = None) -> pd.DataFrame:
    """加载导出的 Excel 数据"""
    current_directory = os.path.dirname(os.path.abspath(__file__))
    if path is None:
        path = os.path.join(current_directory, '..', '导出数据.xlsx')

    if not os.path.exists(path):
        raise FileNotFoundError(ERROR_MESSAGES['NO_FILE'].format(path))

    return pd.read_excel(path)


def analyze_expenditure(df: pd.DataFrame, selected_units: List[str], selected_types: List[str]) -> pd.DataFrame:
    """分析三保支出数据"""
    # 使用多个条件筛选数据
    mask_not_000 = ~df['三保标识'].astype(str).str.match(r'^\[000\]')
    mask_type = df['指标类型'].isin(TARGET_TYPES)
    mask_sanbao = df['三保标识'].isin(selected_types)
    # 排除预算单位为 0 的行（使用 constants.nonzero_unit_mask）
    units_series = df['预算单位']
    mask_not_zero = nonzero_unit_mask(units_series)
    
    # 如果没有选择预算单位，则不添加预算单位筛选条件
    if selected_units:
        mask_units = df['预算单位'].isin(selected_units)
        filtered_df = df[mask_not_000 & mask_type & mask_units & mask_sanbao & mask_not_zero].copy()
    else:
        # 预算单位可选，未选择时不过滤预算单位
        filtered_df = df[mask_not_000 & mask_type & mask_sanbao & mask_not_zero].copy()
    
    if filtered_df.empty:
        raise ValueError(ERROR_MESSAGES['EMPTY_FILTER'])
    
    
    ''' 计算各类金额
        实际支出金额 = 国库集中实际支出金额列之和 + 实拨实际支出金额列之和
        计划金额 = 国库计划金额列之和 + 实拨计划金额列之和
        计划剩余金额 = 国际集中支付计划剩余数列之和 + （实拨计划数-实拨实支出数）列之和
        支付申请金额 = 国库集中申请金额列之和 + 实拨实际支出金额列之和
        在途金额 = （国库集中支付计划数-国库集中支付计划剩余数-国库集中支付申请支出数） + （实拨计划数-实拨实际支出数）
        未回单金额 = 国库集中支付申请数之和-国库集中支付实际支出数之和（*实拨申请数 = 实拨实际支出数，可能存在没有实际支出的情况）
    '''
    filtered_df['实际支出金额'] = sum(filtered_df[col].fillna(0) for col in ACTUAL_COLS)
    filtered_df['计划金额'] = sum(filtered_df[col].fillna(0) for col in PLAN_COLS)
    filtered_df['计划剩余金额'] = sum(filtered_df[col].fillna(0) for col in REMAINING_COLS) + (filtered_df['实拨_计划数']-filtered_df['实拨_实际支出'])
    filtered_df['支付申请金额'] = sum(filtered_df[col].fillna(0) for col in APPLY_COLS)
    filtered_df['在途金额'] = filtered_df['计划金额'] - filtered_df['计划剩余金额'] - filtered_df['支付申请金额']
    filtered_df['未回单金额'] = filtered_df['支付申请金额'] - filtered_df['实际支出金额']
    
    # 根据是否选择了预算单位来决定分组方式
    if selected_units:
        # 按三保标识和预算单位分组汇总
        group_cols = ['三保标识', '预算单位']
    else:
        # 仅按三保标识分组汇总
        group_cols = ['三保标识']
        
    summary = filtered_df.groupby(group_cols).agg({
        '调整预算数': 'sum',
        '计划金额': 'sum',
        '计划剩余金额': 'sum',
        '支付申请金额': 'sum',
        '在途金额': 'sum',
        '未回单金额': 'sum',
        '实际支出金额': 'sum',
        
    }).round(6)
    
    # 计算进度指标（保持原始比例，不乘以100，供Excel百分比格式使用）
    summary['实际支出进度%'] = (summary['实际支出金额'] / summary['调整预算数']).round(4)
    summary['在途+实际支出金额'] = (summary['在途金额'] + summary['实际支出金额']).round(6)
    summary['在途+实际支出进度%'] = (summary['在途+实际支出金额'] / summary['调整预算数']).round(4)
    
    # 重置索引并排序
    summary = summary.reset_index()
    
    # 根据分组方式决定排序方式
    if selected_units:
        summary = summary.sort_values(['预算单位', '三保标识'])
    else:
        summary = summary.sort_values(['三保标识'])
    
    return summary[RESULT_COLS if selected_units else [col for col in RESULT_COLS if col != '预算单位']]


def save_to_excel(summary: pd.DataFrame, output_path: str) -> str:
    """将汇总结果保存到 Excel 文件"""
    if summary.empty:
        raise ValueError(ERROR_MESSAGES['NO_DATA'])
    
    try:
        from openpyxl.styles import Border, Side
        
        # 定义边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='三保进度', index=False)
            worksheet = writer.sheets['三保进度']
            
            # 冻结第一行
            worksheet.freeze_panes = 'A2'
            
            # 设置百分比列的格式
            percent_cols = ['实际支出进度%', '在途+实际支出进度%']
            for col_idx, col_name in enumerate(summary.columns, 1):
                if col_name in percent_cols:
                    # 设置整列的数字格式为百分比
                    column_letter = chr(64 + col_idx)  # A=1, B=2, ...
                    for row_idx in range(2, len(summary) + 2):  # 从第2行开始（第1行是标题）
                        worksheet[f"{column_letter}{row_idx}"].number_format = '0.00%'
                
                # 调整列宽
                try:
                    max_len = max(
                        summary[col_name].astype(str).str.len().max(),
                        len(str(col_name))
                    )
                    if col_idx <= 26:  # Excel列名最大为Z
                        worksheet.column_dimensions[chr(64 + col_idx)].width = max_len + 2
                except Exception:
                    # 忽略列宽调整错误
                    pass
            
            # 为所有数据单元格添加边框（包括标题行）
            for row_idx in range(1, len(summary) + 2):  # 从第1行(标题)到最后一行数据
                for col_idx in range(1, len(summary.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
    
        print(f"\n汇总结果已保存至: {output_path}")
        return output_path
    except Exception as e:
        raise Exception(f"保存Excel文件时出错: {str(e)}")


def display_summary(summary: pd.DataFrame) -> None:
    """格式化显示汇总结果"""
    try:
        print("\n=== 三保类型汇总 ===")
        
        # 设置显示格式
        pd.set_option('display.float_format', lambda x: f'{x:,.6f}' if isinstance(x, float) else str(x))
        
        # 按三保标识汇总
        type_stats = summary.groupby('三保标识').agg({
            '调整预算数': 'sum',
            '计划金额': 'sum',
            '计划剩余金额': 'sum',
            '支付申请金额': 'sum',
            '在途金额': 'sum',
            '未回单金额': 'sum',
            '实际支出金额': 'sum',
            '在途+实际支出金额': 'sum'
        }).round(6)
        
        # 计算汇总进度（为控制台显示，乘以100）
        type_stats['实际支出进度%'] = (type_stats['实际支出金额'] / type_stats['调整预算数'] * 100).round(2)
        type_stats['在途+实际支出进度%'] = (type_stats['在途+实际支出金额'] / type_stats['调整预算数'] * 100).round(2)
        
        print(type_stats)
        
        formatted = summary.copy()
        for col in numeric_cols:
            if col in formatted.columns:
                formatted[col] = formatted[col].apply(lambda x: f'{x:,.6f}')
        for col in percent_cols:
            if col in formatted.columns:
                # 控制台显示时乘以100并添加%符号
                formatted[col] = formatted[col].apply(lambda x: f'{x*100:,.2f}%')

        # 设置显示选项
        with pd.option_context('display.max_rows', None,
                              'display.max_columns', None,
                              'display.width', None,
                              'display.colheader_justify', 'right'):
            print(formatted.to_string(index=False))

        # 显示三保类型统计
        print("\n=== 三保类型汇总 ===")
        type_stats = summary.groupby('三保标识').agg({
            '调整预算数': 'sum',
            '计划金额': 'sum',
            '计划剩余金额': 'sum',
            '在途金额': 'sum',
            '支付申请金额': 'sum',
            '未回单金额': 'sum',
            '实际支出金额': 'sum',
            '在途+实际支出金额': 'sum',
        }).round(2)
        
        # 计算各类型实际支出进度（为控制台显示，乘以100）
        type_stats['实际支出进度%'] = (type_stats['实际支出金额'] / type_stats['调整预算数'] * 100).round(2)
        # 计算在途+实际支出进度（为控制台显示，乘以100）
        type_stats['在途+实际支出进度%'] = (type_stats['在途+实际支出金额'] / type_stats['调整预算数'] * 100).round(2)
        
        # 格式化三保类型汇总数据
        formatted_stats = type_stats.copy()
        for col in numeric_cols:
            if col in formatted_stats.columns:
                formatted_stats[col] = formatted_stats[col].apply(lambda x: f'{x:,.2f}')
        for col in ['实际支出进度%', '在途+实际支出进度%']:
            if col in formatted_stats.columns:
                formatted_stats[col] = formatted_stats[col].apply(lambda x: f'{x:,.2f}%')
        
        # 为最终输出的统计数据也添加百分号
        print(formatted_stats.to_string(float_format=lambda x: f'{x:,.2f}' if isinstance(x, (int, float)) and str(x) not in [c for c in formatted_stats.columns if "%" in c] else f'{x:,.2f}%'))
    except Exception as e:
        print(f"显示汇总结果时出错: {str(e)}")


def main():
    """主程序入口"""
    import sys
    app = QApplication(sys.argv)
    window = ExpenditureAnalyzer()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()