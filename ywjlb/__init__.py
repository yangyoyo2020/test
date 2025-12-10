"""
运维记录簿（YWJLB）处理系统
支持三种包类型的Excel转Word文档处理
"""

from .ywjlb_unified import (
    PackageType,
    process_excel_file,
    create_word_document,
    save_word_document
)

__version__ = "1.0.0"
__all__ = [
    "PackageType",
    "process_excel_file",
    "create_word_document",
    "save_word_document",
]
