"""
双语Excel表格同步技能主程序入口

允许通过 `python -m bilingual_excel_sync` 运行技能
"""

import sys
import os

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from scripts.bilingual_excel_sync import main

if __name__ == "__main__":
    main()