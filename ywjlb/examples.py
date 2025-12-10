"""
运维记录簿处理系统 - 使用示例
"""

# 示例1：命令行使用 - 处理01包-金财工程
def example_1_gjgc():
    """处理01包-金财工程数据"""
    from ywjlb_unified import PackageType, process_excel_file
    
    excel_file = "运维记录表.xlsx"
    success, fail = process_excel_file(excel_file, PackageType.GJGC)
    print(f"01包-金财工程处理完成: 成功{success}个，失败{fail}个")


# 示例2：命令行使用 - 处理01包-其他项目
def example_2_qtxm():
    """处理01包-其他项目数据"""
    from ywjlb_unified import PackageType, process_excel_file
    
    excel_file = "运维记录表.xlsx"
    success, fail = process_excel_file(excel_file, PackageType.QTXM)
    print(f"01包-其他项目处理完成: 成功{success}个，失败{fail}个")


# 示例3：命令行使用 - 处理02包项目
def example_3_pkg02():
    """处理02包项目数据"""
    from ywjlb_unified import PackageType, process_excel_file
    
    excel_file = "运维记录表.xlsx"
    success, fail = process_excel_file(excel_file, PackageType.PKG02)
    print(f"02包项目处理完成: 成功{success}个，失败{fail}个")


# 示例4：批量处理所有包类型
def example_4_batch_all():
    """批量处理所有包类型"""
    from ywjlb_unified import PackageType, process_excel_file
    
    excel_file = "运维记录表.xlsx"
    
    for package_type in PackageType:
        print(f"\n正在处理: {package_type.value}")
        try:
            success, fail = process_excel_file(excel_file, package_type)
            print(f"  ✓ 成功: {success}, 失败: {fail}")
        except Exception as e:
            print(f"  ✗ 错误: {str(e)}")


# 示例5：编程方式 - 自定义处理流程
def example_5_custom_process():
    """自定义处理流程示例"""
    import pandas as pd
    from ywjlb_unified import PackageType, create_word_document, save_word_document
    
    # 读取Excel文件
    excel_file = "运维记录表.xlsx"
    df_dict = pd.read_excel(excel_file, sheet_name=None)
    
    package_type = PackageType.GJGC
    
    # 遍历每个工作表
    for sheet_name, df in df_dict.items():
        print(f"处理工作表: {sheet_name}")
        
        # 遍历每一行
        for idx, row in enumerate(df.iterrows(), start=1):
            try:
                # 创建Word文档
                doc = create_word_document(row[1], idx, package_type)
                
                # 保存文档
                save_word_document(doc, idx, sheet_name, row[1].get('日期', ''))
                
                print(f"  ✓ 第{idx}行已处理")
            except Exception as e:
                print(f"  ✗ 第{idx}行处理失败: {str(e)}")


# 示例6：启动图形界面应用
def example_6_gui_app():
    """启动图形用户界面"""
    import sys
    from PyQt6.QtWidgets import QApplication
    from ywjlb_ui import YWJLBAnalyzer
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = YWJLBAnalyzer()
    window.show()
    
    sys.exit(app.exec())


# 示例7：获取包类型信息
def example_7_package_types_info():
    """获取所有包类型的信息"""
    from ywjlb_unified import PackageType, PACKAGE_FIELDS
    
    for pkg_type in PackageType:
        print(f"\n包类型: {pkg_type.value}")
        fields = PACKAGE_FIELDS[pkg_type]
        print(f"字段数: {len(fields)}")
        print("字段列表:")
        for label, _ in fields:
            print(f"  - {label}")


# 示例8：错误处理和日志记录
def example_8_error_handling():
    """演示错误处理"""
    import os
    import logging
    from ywjlb_unified import PackageType, process_excel_file
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    excel_file = "运维记录表.xlsx"
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        logging.error(f"文件不存在: {excel_file}")
        return
    
    # 处理文件
    try:
        success, fail = process_excel_file(excel_file, PackageType.GJGC)
        logging.info(f"处理完成: 成功{success}个, 失败{fail}个")
        
        # 检查输出目录
        if os.path.exists("运维记录文档"):
            doc_count = len(os.listdir("运维记录文档"))
            logging.info(f"生成的文档数: {doc_count}")
            
    except FileNotFoundError as e:
        logging.error(f"文件错误: {str(e)}")
    except Exception as e:
        logging.error(f"处理错误: {str(e)}")


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        example_num = sys.argv[1]
        
        examples = {
            '1': example_1_gjgc,
            '2': example_2_qtxm,
            '3': example_3_pkg02,
            '4': example_4_batch_all,
            '5': example_5_custom_process,
            '6': example_6_gui_app,
            '7': example_7_package_types_info,
            '8': example_8_error_handling,
        }
        
        if example_num in examples:
            examples[example_num]()
        else:
            print("使用方法: python examples.py <示例编号>")
            print("\n可用示例:")
            print("  1 - 处理01包-金财工程")
            print("  2 - 处理01包-其他项目")
            print("  3 - 处理02包项目")
            print("  4 - 批量处理所有包类型")
            print("  5 - 自定义处理流程")
            print("  6 - 启动图形界面")
            print("  7 - 获取包类型信息")
            print("  8 - 错误处理演示")
    else:
        print("运维记录簿处理系统 - 使用示例")
        print("\n使用方法: python examples.py <示例编号>")
        print("\n可用示例:")
        print("  1 - 处理01包-金财工程")
        print("  2 - 处理01包-其他项目")
        print("  3 - 处理02包项目")
        print("  4 - 批量处理所有包类型")
        print("  5 - 自定义处理流程")
        print("  6 - 启动图形界面")
        print("  7 - 获取包类型信息")
        print("  8 - 错误处理演示")
        print("\n示例: python examples.py 1")
