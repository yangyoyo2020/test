"""
快速开始指南 - 运维记录簿处理系统

本文件展示如何快速使用本系统的三种方式
"""

# ============================================================================
# 方式一：最快速 - 图形界面（推荐新手）
# ============================================================================

def start_gui():
    """
    启动图形用户界面
    
    使用步骤：
    1. 点击"浏览"按钮选择Excel文件
    2. 从下拉菜单选择包类型
    3. 点击"加载数据"预览文件
    4. 点击"处理并导出"生成Word文档
    5. 检查"运维记录文档"文件夹
    """
    import sys
    from PyQt6.QtWidgets import QApplication
    from ywjlb_ui import YWJLBAnalyzer
    
    app = QApplication(sys.argv)
    window = YWJLBAnalyzer()
    window.show()
    sys.exit(app.exec())


# ============================================================================
# 方式二：快速脚本 - 命令行使用（推荐开发者）
# ============================================================================

def quick_process():
    """
    使用命令行快速处理Excel文件
    
    支持三种包类型：
    - PackageType.GJGC: 01包-金财工程
    - PackageType.QTXM: 01包-其他项目  
    - PackageType.PKG02: 02包项目
    """
    from ywjlb_unified import PackageType, process_excel_file
    
    # 准备输入文件
    excel_file = "运维记录表.xlsx"
    
    # 选择包类型
    package_type = PackageType.GJGC  # 改为QTXM或PKG02试试
    
    # 执行处理
    try:
        print(f"正在处理: {package_type.value}")
        success, fail = process_excel_file(excel_file, package_type)
        print(f"✓ 完成！成功: {success}个，失败: {fail}个")
        print("✓ 文档已保存到: 运维记录文档/")
        print("✓ 日志已保存到: 转换日志.log")
    except Exception as e:
        print(f"✗ 错误: {str(e)}")


# ============================================================================
# 方式三：高级定制 - 编程方式（推荐集成）
# ============================================================================

def advanced_custom():
    """
    使用编程方式进行高级定制处理
    """
    import pandas as pd
    from ywjlb_unified import (
        PackageType, 
        PACKAGE_FIELDS,
        create_word_document, 
        save_word_document
    )
    
    # 1. 加载Excel文件
    excel_file = "运维记录表.xlsx"
    df_dict = pd.read_excel(excel_file, sheet_name=None)
    
    # 2. 选择包类型
    package_type = PackageType.GJGC
    
    # 3. 获取该包类型的字段配置
    fields = PACKAGE_FIELDS[package_type]
    print(f"包类型: {package_type.value}")
    print(f"字段数: {len(fields)}")
    
    # 4. 处理每个工作表
    total_count = 0
    for sheet_name, df in df_dict.items():
        print(f"\n处理工作表: {sheet_name}")
        
        # 5. 处理每一行数据
        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            try:
                # 创建Word文档
                doc = create_word_document(row, idx, package_type)
                
                # 保存文档
                save_word_document(doc, idx, sheet_name, row.get('日期', ''))
                
                total_count += 1
                print(f"  ✓ 第{idx}行已处理")
            except Exception as e:
                print(f"  ✗ 第{idx}行失败: {str(e)}")
    
    print(f"\n处理完成！总共生成{total_count}个文档")


# ============================================================================
# 快速参考 - 常见任务
# ============================================================================

def quick_reference():
    """快速参考 - 常见任务代码片段"""
    
    print("""
    ╔════════════════════════════════════════════════════════════════════╗
    ║           运维记录簿处理系统 - 快速参考                            ║
    ╚════════════════════════════════════════════════════════════════════╝
    
    【任务1】使用默认设置处理01包-金财工程
    ─────────────────────────────────────────────────────────────────────
    from ywjlb_unified import PackageType, process_excel_file
    
    success, fail = process_excel_file(
        "运维记录表.xlsx", 
        PackageType.GJGC
    )
    print(f"成功: {success}, 失败: {fail}")
    
    
    【任务2】处理所有包类型
    ─────────────────────────────────────────────────────────────────────
    from ywjlb_unified import PackageType, process_excel_file
    
    for pkg_type in PackageType:
        print(f"处理: {pkg_type.value}")
        success, fail = process_excel_file("运维记录表.xlsx", pkg_type)
        print(f"  成功: {success}, 失败: {fail}")
    
    
    【任务3】启动图形界面
    ─────────────────────────────────────────────────────────────────────
    python ywjlb_app.py
    
    # 或
    python -m ywjlb.ywjlb_app
    
    
    【任务4】获取包类型字段信息
    ─────────────────────────────────────────────────────────────────────
    from ywjlb_unified import PackageType, PACKAGE_FIELDS
    
    for pkg_type in PackageType:
        fields = PACKAGE_FIELDS[pkg_type]
        print(f"{pkg_type.value}:")
        for label, key in fields:
            print(f"  - {label}")
    
    
    【任务5】自定义错误处理
    ─────────────────────────────────────────────────────────────────────
    from ywjlb_unified import PackageType, process_excel_file
    import logging
    
    logging.basicConfig(level=logging.DEBUG)
    
    try:
        success, fail = process_excel_file(
            "运维记录表.xlsx",
            PackageType.QTXM
        )
    except FileNotFoundError:
        print("找不到文件")
    except Exception as e:
        logging.error(f"处理失败: {e}")
    
    ════════════════════════════════════════════════════════════════════════
    
    更多信息请查看 README.md 或 examples.py
    """)


# ============================================================================
# 主程序
# ============================================================================

if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        
        if command == 'gui':
            print("启动图形用户界面...")
            start_gui()
        elif command == 'quick':
            print("执行快速处理...")
            quick_process()
        elif command == 'advanced':
            print("执行高级定制处理...")
            advanced_custom()
        elif command == 'help':
            quick_reference()
        else:
            print("未知命令。使用: python quickstart.py [gui|quick|advanced|help]")
    else:
        print("""
╔════════════════════════════════════════════════════════════════════╗
║         运维记录簿处理系统 - 快速开始指南                          ║
╚════════════════════════════════════════════════════════════════════╝

使用方法:
  
  1. 启动图形界面（推荐）:
     python quickstart.py gui
  
  2. 快速命令行处理:
     python quickstart.py quick
  
  3. 高级定制处理:
     python quickstart.py advanced
  
  4. 查看快速参考:
     python quickstart.py help

════════════════════════════════════════════════════════════════════════

各方式说明：

【GUI方式】- 最友好，无需编程
  ✓ 友好的图形界面
  ✓ 可视化的处理进度
  ✓ 实时日志显示
  适合：所有用户

【快速方式】- 最简单，一行命令
  ✓ 开箱即用
  ✓ 代码最少
  ✓ 快速集成
  适合：需要快速处理的用户

【高级方式】- 最灵活，完全定制
  ✓ 完全的控制权
  ✓ 支持自定义逻辑
  ✓ 适合复杂场景
  适合：需要集成到大型系统的用户

════════════════════════════════════════════════════════════════════════

更多信息：
  - README.md     详细使用文档
  - examples.py   8个完整示例
  - COMPLETION_REPORT.md  项目完成报告

        """)
