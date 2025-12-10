# 运维记录簿（YWJLB）处理系统

## 项目概述

运维记录簿处理系统是一个统一的工具，用于合并和处理三种不同包类型的运维数据，并自动生成Word文档。

### 支持的包类型

1. **01包-金财工程** (`PackageType.GJGC`)
   - 字段：项目名称、需求名称、使用单位、监理单位、运维单位、提出问题单位、日期、运维人员、**系统模块**、问题类型、问题描述、处理方法、处理结果

2. **01包-其他项目** (`PackageType.QTXM`)
   - 字段：项目名称、需求名称、使用单位、监理单位、运维单位、日期、运维人员、问题类型、问题描述、处理方法、处理结果

3. **02包项目** (`PackageType.PKG02`)
   - 字段：项目名称、需求名称、使用单位、监理单位、运维单位、日期、运维人员、问题类型、问题描述、处理方法、处理结果

## 系统架构

### 核心模块

#### `ywjlb_unified.py` - 数据处理模块
包含所有Excel转Word的核心逻辑：
- `PackageType` - 枚举类，定义三种包类型
- `PACKAGE_FIELDS` - 字典，定义各包类型对应的字段列表
- 文档生成函数：`create_word_document()`, `save_word_document()`
- 主处理函数：`process_excel_file()`

**主要功能：**
- 根据包类型加载对应的字段配置
- 将Excel数据转换为格式化的Word文档
- 支持日期格式化和特殊字符处理
- 提供详细的日志记录

#### `ywjlb_ui.py` - 用户界面模块
基于PyQt6的图形用户界面：
- 文件选择和加载
- 包类型选择
- 数据预览
- 实时日志显示
- 处理进度反馈

**界面组件：**
1. **文件选择区** - 浏览并选择Excel文件
2. **包类型选择区** - 从下拉菜单选择处理类型
3. **数据预览区** - 显示Excel文件的结构信息
4. **处理日志区** - 实时显示处理过程
5. **操作按钮区** - 加载、处理、导出和退出功能

#### `ywjlb_app.py` - 应用入口
启动图形化应用程序的主入口

## 使用方法

### 命令行使用（核心模块）

```python
from ywjlb_unified import PackageType, process_excel_file

# 处理01包-金财工程
success, fail = process_excel_file("运维记录表.xlsx", PackageType.GJGC)
print(f"成功: {success}, 失败: {fail}")

# 处理01包-其他项目
success, fail = process_excel_file("运维记录表.xlsx", PackageType.QTXM)

# 处理02包项目
success, fail = process_excel_file("运维记录表.xlsx", PackageType.PKG02)
```

### 图形界面使用

```bash
python -m ywjlb.ywjlb_app
# 或
python ywjlb_app.py
```

**使用步骤：**
1. 点击"浏览"按钮选择Excel文件
2. 从下拉菜单选择对应的包类型
3. 点击"加载数据"按钮预览文件内容
4. 点击"处理并导出"按钮生成Word文档
5. 查看日志了解处理进度
6. 生成的Word文档保存在 `运维记录文档/` 文件夹中

## 文件结构

```
ywjlb/
├── __init__.py              # 包初始化文件
├── ywjlb_unified.py         # 核心处理模块
├── ywjlb_ui.py              # PyQt6用户界面
├── ywjlb_app.py             # 应用程序入口
└── README.md                # 本文档
```

## 输出文件

### Word文档
- **位置：** `运维记录文档/` 文件夹
- **命名规则：** `运维记录_{表名}_第{序号}页_{日期}.docx`
- **示例：** `运维记录_工程项目_第1页_20250101.docx`

### 日志文件
- **位置：** 项目根目录
- **文件名：** `转换日志.log`
- **内容：** 详细的处理日志和错误信息

## 依赖环境

- Python 3.8+
- pandas - 数据处理
- python-docx - Word文档生成
- PyQt6 - 图形界面（可选，只有使用UI时需要）

## 字段映射

### 共同字段
所有包类型都包含：
- 项目名称
- 需求名称
- 使用单位
- 监理单位
- 运维单位
- 日期
- 运维人员
- 问题类型
- 问题描述
- 处理方法
- 处理结果

### 特定字段
- **01包-金财工程** 额外包含：
  - 提出问题单位
  - 系统模块

## 常见问题

### Q: 如何处理不同格式的日期？
A: 系统会自动识别多种日期格式，包括：
- pandas Timestamp对象
- ISO格式（YYYY-MM-DD）
- 中文格式（YYYY年MM月DD日）
- 其他常见格式

处理后统一输出为"YYYY年MM月DD日"格式。

### Q: 文档生成失败如何排查？
A: 
1. 检查日志文件 `转换日志.log`
2. 确保Excel文件不被其他程序占用
3. 验证Excel文件是否包含所需字段
4. 检查输出文件夹的写入权限

### Q: 可以同时处理多个包类型吗？
A: 可以，但需要在UI中改变包类型选择后重新处理，或通过编程使用核心模块多次调用`process_excel_file()`。

## 自定义和扩展

### 添加新的包类型

1. 在 `ywjlb_unified.py` 中修改 `PackageType` 枚举：
```python
class PackageType(Enum):
    GJGC = "01包-金财工程"
    QTXM = "01包-其他项目"
    PKG02 = "02包项目"
    PKG03 = "03包项目"  # 新增
```

2. 在 `PACKAGE_FIELDS` 中添加字段配置：
```python
PACKAGE_FIELDS = {
    # ... 其他包类型 ...
    PackageType.PKG03: [
        ("字段1", "字段1"),
        ("字段2", "字段2"),
        # ...
    ]
}
```

3. 在UI的 `ywjlb_ui.py` 中添加选项：
```python
self.package_type_combo.addItem(PackageType.PKG03.value, PackageType.PKG03)
```

### 修改Word文档格式

在 `create_content_table()` 或 `create_word_document()` 函数中修改样式设置，例如：
- 字体大小：修改 `Pt()` 参数
- 列宽：修改 `Cm()` 参数
- 对齐方式：修改 `alignment` 参数
- 行高：修改 `height` 参数

## 版本历史

- **v1.0.0** (2025-01-10)
  - 首次发布
  - 支持三种包类型
  - 完整的PyQt6用户界面
  - 详细的日志记录

## 许可证

内部使用

## 作者

开发团队

## 联系方式

如有问题或建议，请联系开发团队。
