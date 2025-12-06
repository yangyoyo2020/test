import logging
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler

try:
    from PyQt6.QtCore import QObject
except Exception:
    QObject = None

DEFAULT_LOG_FILE = Path(os.getcwd()) / "日志文件.log"


def _level_from_env(env_value: str | None) -> int:
    """Convert environment LOG_LEVEL string to logging level int."""
    if not env_value:
        return logging.INFO
    name = env_value.strip().upper()
    return getattr(logging, name, logging.INFO)


class QtWidgetHandler(logging.Handler):
    """将日志输出到 QTextEdit 文本控件（通过 append）。"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            msg = self.format(record)
            if self.text_widget is not None:
                # 直接 append（UI 线程约束视具体使用情形而定）
                self.text_widget.append(msg)
                try:
                    self.text_widget.verticalScrollBar().setValue(
                        self.text_widget.verticalScrollBar().maximum()
                    )
                except Exception:
                    pass
        except Exception:
            pass


class QtSignalHandler(logging.Handler):
    """将日志通过 Qt 信号发送到 UI 的 handler（接受任何有 emit 方法的对象）。"""
    def __init__(self, signal):
        super().__init__()
        self.signal = signal

    def emit(self, record):
        try:
            msg = self.format(record)
            if self.signal is not None:
                # 信号对象通常有 emit 方法
                try:
                    self.signal.emit(msg)
                except Exception:
                    # 有些 signal 可能要求在主线程调用 emit；调用方需注意
                    pass
        except Exception:
            pass


def get_logger(name: str = 'app', level: int = None, log_file: str | Path = None,
               qt_signal: object = None, qt_text_widget: object = None) -> logging.Logger:
    """返回配置好的 logging.Logger。

    - `qt_signal`：若提供，日志会通过 `qt_signal.emit(str)` 发送（适用于 pyqt 信号）。
    - `qt_text_widget`：若提供，日志会直接追加到该 QTextEdit（UI 控件）。
    - `log_file`：日志文件路径，默认为工作目录下 `日志文件.log`。
    """
    # 优先从环境变量读取配置
    env_level = os.environ.get('LOG_LEVEL')
    env_log_file = os.environ.get('LOG_FILE')

    if level is None:
        level = _level_from_env(env_level)

    if log_file is None and env_log_file:
        log_file = Path(env_log_file)

    logger = logging.getLogger(name)
    logger.setLevel(level)

    # 清理已有处理器，避免重复
    if logger.hasHandlers():
        logger.handlers.clear()

    if log_file is None:
        log_file = DEFAULT_LOG_FILE

    # 文件处理器，使用按大小轮转以避免日志无限增长（可通过环境变量配置）
    try:
        max_bytes = int(os.environ.get('LOG_MAX_BYTES', 5 * 1024 * 1024))
    except Exception:
        max_bytes = 5 * 1024 * 1024
    try:
        backup_count = int(os.environ.get('LOG_BACKUP_COUNT', 5))
    except Exception:
        backup_count = 5

    try:
        fh = RotatingFileHandler(log_file, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8')
    except Exception:
        # 若文件句柄不可用，回退到标准输出
        fh = logging.StreamHandler()
    fh.setLevel(level)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Qt 相关处理器（优先使用 signal，其次使用文本控件）
    if qt_signal is not None:
        qh = QtSignalHandler(qt_signal)
        qh.setLevel(level)
        qh.setFormatter(formatter)
        logger.addHandler(qh)
    elif qt_text_widget is not None:
        qh = QtWidgetHandler(qt_text_widget)
        qh.setLevel(level)
        qh.setFormatter(formatter)
        logger.addHandler(qh)

    return logger


def add_qt_text_widget(logger: logging.Logger, qt_text_widget: object):
    """为已存在的 logger 添加 QTextEdit 输出处理器。"""
    if logger is None or qt_text_widget is None:
        return
    # 避免重复添加同一个 QTextEdit 的处理器，但允许为不同的文本控件各自添加处理器
    for h in logger.handlers:
        if isinstance(h, QtWidgetHandler) and getattr(h, 'text_widget', None) is qt_text_widget:
            return
    handler = QtWidgetHandler(qt_text_widget)
    handler.setLevel(logger.level)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)


def add_qt_signal(logger: logging.Logger, qt_signal: object):
    """为已存在的 logger 添加 Qt 信号输出处理器。"""
    if logger is None or qt_signal is None:
        return
    # 允许为同一 logger 注册多个不同的信号处理器，但如果已经为相同的 signal 注册过则跳过
    for h in logger.handlers:
        if isinstance(h, QtSignalHandler) and getattr(h, 'signal', None) is qt_signal:
            return
    handler = QtSignalHandler(qt_signal)
    handler.setLevel(logger.level)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)
