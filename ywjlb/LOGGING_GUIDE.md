# 日志配置使用指南

## 概述

该项目采用统一的日志配置系统，在应用启动时由 `ywjlb_app.py` 统一初始化。所有模块均通过 `logging.getLogger(__name__)` 获取模块级日志记录器。

## 日志配置特性

### 1. 双渠道输出
- **文件输出**：输出到 `log/ywjlb.log`，支持按大小轮转（10MB/文件，保留5个备份）
- **控制台输出**：输出到标准输出流，便于实时监控

### 2. 日志格式
- **文件日志**：`%(asctime)s - %(name)s - %(levelname)s - %(message)s`
- **控制台日志**：`%(levelname)s - %(message)s`（简化格式便于阅读）

### 3. 默认日志级别
- 应用启动时设置为 `INFO`，可根据需要调整
- 环境变量可覆盖：`LOG_LEVEL=DEBUG`

## 使用方法

### 在各模块中使用日志

```python
import logging

# 在文件顶部获取模块日志记录器
logger = logging.getLogger(__name__)

# 使用示例
logger.info("处理开始")
logger.debug("调试信息")
logger.warning("警告信息")
logger.error("错误信息")
logger.exception("异常信息（自动包含堆栈追踪）")
```

### 日志级别说明

| 级别 | 用途 | 示例 |
|------|------|------|
| DEBUG | 详细调试信息 | `logger.debug(f"读取文件: {file_path}")` |
| INFO | 一般信息 | `logger.info("处理完成，生成10个文档")` |
| WARNING | 警告信息 | `logger.warning("文件不存在，使用默认值")` |
| ERROR | 错误信息 | `logger.error("无法打开文件")` |
| EXCEPTION | 异常信息 | `logger.exception("处理出错")` |

## 日志输出示例

### 文件日志 (log/ywjlb.log)
```
2025-12-10 10:30:45,123 - ywjlb.ywjlb_ui - INFO - 应用程序启动成功，等待用户交互
2025-12-10 10:30:50,456 - ywjlb.ywjlb_ui - DEBUG - 用户选择文件: D:\data.xlsx
2025-12-10 10:30:55,789 - ywjlb.ywjlb_unified - INFO - 开始处理文件...
2025-12-10 10:31:00,012 - ywjlb.ywjlb_unified - ERROR - 处理第5行时出错: 缺少必需列
```

### 控制台输出
```
INFO - ==================================================
INFO - 运维记录簿处理工具启动
INFO - ==================================================
INFO - 应用程序启动成功，等待用户交互
DEBUG - 用户选择文件: D:\data.xlsx
INFO - 开始处理文件...
ERROR - 处理第5行时出错: 缺少必需列
```

## 常见场景

### 处理异常时的正确做法

```python
try:
    result = process_file(file_path)
    logger.info(f"文件处理成功，生成{len(result)}个文档")
except FileNotFoundError as e:
    logger.error(f"文件不存在: {file_path}")
except Exception as e:
    # 使用 exception() 会自动包含完整的堆栈追踪
    logger.exception(f"处理文件时发生未知错误: {file_path}")
```

### 性能敏感代码中的调试

```python
import time

start_time = time.time()
logger.debug(f"开始处理{len(data)}条记录")

for i, item in enumerate(data):
    process_item(item)
    if (i + 1) % 100 == 0:
        logger.debug(f"已处理{i + 1}条，耗时{time.time() - start_time:.2f}s")

logger.info(f"总共处理{len(data)}条记录，耗时{time.time() - start_time:.2f}s")
```

## 日志文件位置

- **日志文件**：`<项目根目录>/log/ywjlb.log`
- **轮转备份**：`ywjlb.log.1`, `ywjlb.log.2` 等
- **当前工作目录**：通过应用启动时的 `log_dir` 参数指定

## 运行时调整日志级别

在 `ywjlb_app.py` 中的 `setup_logging()` 调用处修改 `level` 参数：

```python
# 调试模式（输出详细信息）
setup_logging(level=logging.DEBUG)

# 生产模式（只输出重要信息）
setup_logging(level=logging.WARNING)
```

## 注意事项

1. **不要混用配置方式**：模块中不应再调用 `logging.basicConfig()`，统一通过应用入口配置
2. **线程安全**：Python logging 模块已实现线程安全，可安全用于多线程场景
3. **性能考虑**：DEBUG 级别日志在高频调用场景可能影响性能，生产环境建议使用 INFO 级别
4. **Qt 集成**：如需将日志输出到 UI 文本控件，可使用 `common.logger` 中的 `QtWidgetHandler`

## 相关模块

- `common/logger.py`：核心日志处理类和函数
- `ywjlb/ywjlb_app.py`：应用启动和日志初始化入口
