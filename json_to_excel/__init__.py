"""Package entry for json_to_excel. Expose core functions for import by GUI or other modules."""
from .core import (
    normalize_top_items,
    split_parent_children,
    coerce_numeric_columns,
    parse_date_columns,
    write_excel_flat,
    write_excel_multi,
    convert_flat,
    convert_multi,
    stream_to_csv,
)

from . import cli

__all__ = [
    'normalize_top_items',
    'split_parent_children',
    'coerce_numeric_columns',
    'parse_date_columns',
    'write_excel_flat',
    'write_excel_multi',
    'convert_flat',
    'convert_multi',
    'stream_to_csv',
    'cli',
]
