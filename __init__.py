"""
双语Excel表格同步技能

将中文Excel表格自动转换为中英双语版本，并同步生成纯英文SHEET2工作表。
"""

__version__ = "1.0.0"
__author__ = "OpenCode"
__license__ = "MIT"

from .scripts.bilingual_excel_sync import BilingualExcelSync

__all__ = ["BilingualExcelSync"]