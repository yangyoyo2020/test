"""
运维记录簿（YWJLB）应用程序入口
"""

import sys
import os
import logging
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from ywjlb_ui import YWJLBAnalyzer


def setup_logging(log_dir: str = None, level: int = logging.INFO):
    """统一配置应用程序日志系统
    
    Args:
        log_dir: 日志文件目录，默认为应用根目录下的 'log' 文件夹
        level: 日志级别，默认为 INFO
    """
    if log_dir is None:
        log_dir = Path(__file__).parent.parent / 'log'
    else:
        log_dir = Path(log_dir)
    
    # 创建日志目录
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / 'ywjlb.log'
    
    # 配置根 logger
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    
    # 清除已存在的处理器
    root_logger.handlers.clear()
    
    # 文件处理器（带轮转功能）
    from logging.handlers import RotatingFileHandler
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(level)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)
    root_logger.addHandler(file_handler)
    
    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_formatter = logging.Formatter(
        '%(levelname)s - %(message)s'
    )
    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)


def main():
    """应用程序主入口"""
    try:
        # 初始化日志系统
        setup_logging(level=logging.INFO)
        logger = logging.getLogger(__name__)
        
        logger.info("="*50)
        logger.info("运维记录簿处理工具启动")
        logger.info("="*50)
        
        app = QApplication(sys.argv)
        
        # 设置应用样式和元数据
        app.setStyle('Fusion')
        app.setApplicationName("运维记录簿处理工具")
        app.setApplicationVersion("1.0.0")
        
        # 创建并显示主窗口
        window = YWJLBAnalyzer()
        window.show()
        
        logger.info("应用程序启动成功，等待用户交互")
        
        # 启动事件循环
        exit_code = app.exec()
        logger.info(f"应用程序退出，退出码: {exit_code}")
        sys.exit(exit_code)
        
    except Exception as e:
        logger.exception("应用程序启动失败")
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
