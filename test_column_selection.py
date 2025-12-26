#!/usr/bin/env python3
"""
测试列选择功能的演示脚本
"""

import sys
import pandas as pd
from pathlib import Path

# 添加路径
sys.path.insert(0, str(Path(__file__).parent))

# 创建示例Excel文件用于测试
def create_sample_excel():
    """创建示例Excel文件"""
    data = {
        '预算单位': ['单位A', '单位B', '单位C', '单位A', '单位B'],
        '部门名称': ['部门1', '部门2', '部门3', '部门1', '部门2'],
        '项目名称': ['项目甲', '项目乙', '项目丙', '项目丁', '项目戊'],
        '三保标识': ['医保', '失保', '工伤保', '医保', '失保'],
        '金额': [1000, 2000, 3000, 4000, 5000],
    }
    df = pd.DataFrame(data)
    df.to_excel('sample_data.xlsx', index=False, sheet_name='数据')
    print("已创建示例文件: sample_data.xlsx")
    return 'sample_data.xlsx'

if __name__ == '__main__':
    # 创建示例文件
    sample_file = create_sample_excel()
    
    # 启动应用
    from sanbao_test.app_copy import ExpenditureAnalyzer
    app = __import__('PyQt6.QtWidgets', fromlist=['QApplication']).QApplication(sys.argv)
    
    window = ExpenditureAnalyzer()
    window.show()
    
    sys.exit(app.exec())
