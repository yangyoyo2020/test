# 日志配置改进 - 快速参考

## 改进概览

✅ **统一日志入口** - 日志在应用启动时统一初始化  
✅ **文件 + 控制台输出** - 支持双渠道同时记录  
✅ **自动轮转** - 日志文件超过10MB自动备份  
✅ **模块级日志** - 每条日志都能显示来源模块  
✅ **错误捕获** - 应用启动时的异常能被正确处理  

## 核心改动

### 文件 1: ywjlb_app.py ⭐ 主要改动
```python
# 新增 setup_logging() 函数
# - 配置根logger
# - 创建文件处理器（10MB轮转）
# - 创建控制台处理器

# 改进 main() 函数
# - 初始化日志系统
# - 添加异常处理
# - 添加应用程序元数据
```

### 文件 2: ywjlb_unified.py
```python
# 移除了 logging.basicConfig()
# 改为: logger = logging.getLogger(__name__)
```

### 文件 3: ywjlb_ui.py
```python
# 添加了 import logging
# 改为: logger = logging.getLogger(__name__)
```

## 快速使用

### 在你的模块中
```python
import logging
logger = logging.getLogger(__name__)

# 记录日志
logger.info("处理开始")
logger.error("出错了")
logger.exception("有异常发生")
```

### 日志输出位置
- **实时控制台**：运行应用时看到
- **日志文件**：`log/ywjlb.log`
- **备份文件**：`log/ywjlb.log.1`, `.log.2` 等

## 日志级别速查

| 级别 | 用途 | 例子 |
|------|------|------|
| DEBUG | 详细调试 | 变量值、参数检查 |
| INFO | 一般信息 | 处理开始/完成 |
| WARNING | 警告 | 使用了默认值 |
| ERROR | 错误 | 文件读写失败 |

## 常见场景

### ✅ 正确做法
```python
try:
    result = process(file)
    logger.info(f"处理成功: {len(result)} 项")
except Exception as e:
    logger.exception("处理失败")  # 自动包含堆栈
    raise
```

### ❌ 避免
```python
# 不要在模块中调用 basicConfig()
logging.basicConfig(...)  # ❌ 错误

# 不要创建全局logger
logger = logging.getLogger('myapp')  # 考虑使用 __name__
```

## 文件改动统计

| 文件 | 行数 | 改动类型 |
|------|------|---------|
| ywjlb_app.py | 103 (+74) | 添加日志系统和异常处理 |
| ywjlb_unified.py | 450 (-8) | 移除basicConfig |
| ywjlb_ui.py | 476 (+2) | 添加logger导入 |

## 下次运行

当你运行应用时，会看到类似输出：

```
INFO - ==================================================
INFO - 运维记录簿处理工具启动
INFO - ==================================================
INFO - 应用程序启动成功，等待用户交互
```

同时日志文件会自动创建在 `log/ywjlb.log`

## 更多帮助

- 详细指南：查看 `LOGGING_GUIDE.md`
- 改进说明：查看 `LOG_IMPROVEMENTS.md`
