# 日志配置改进总结

## 改进内容

### 1. **统一日志入口** ✓
- 将所有日志配置从各个模块统一移到应用启动入口 `ywjlb_app.py`
- 移除了 `ywjlb_unified.py` 中的 `logging.basicConfig()` 配置

### 2. **完整的日志系统** ✓

#### 文件输出配置
```python
# 日志文件位置: log/ywjlb.log
# 单个日志文件大小限制: 10MB
# 备份文件数量: 5个
# 编码格式: UTF-8
```

#### 日志格式
- **文件日志**：`%(asctime)s - %(name)s - %(levelname)s - %(message)s`
- **控制台日志**：`%(levelname)s - %(message)s`（简化便于阅读）

### 3. **模块级日志记录器** ✓
所有模块现在都使用标准的模块级记录器：
```python
logger = logging.getLogger(__name__)
```

这样每条日志都会显示来源模块，便于追踪问题。

### 4. **改进的异常处理** ✓
在 `main()` 函数中添加完整的错误捕获：
```python
try:
    # 应用逻辑
except Exception as e:
    logger.exception("应用程序启动失败")
    sys.exit(1)
```

### 5. **应用元数据** ✓
添加了应用程序的基本元数据：
```python
app.setApplicationName("运维记录簿处理工具")
app.setApplicationVersion("1.0.0")
```

## 改动的文件

| 文件 | 改动说明 |
|------|---------|
| `ywjlb_app.py` | ✓ 添加 `setup_logging()` 函数和错误处理 |
| `ywjlb_unified.py` | ✓ 移除 `basicConfig()`，使用模块级 logger |
| `ywjlb_ui.py` | ✓ 添加 logging 导入和模块级 logger |

## 使用示例

### 在模块中记录日志
```python
import logging

logger = logging.getLogger(__name__)

def process_file(file_path):
    logger.info(f"开始处理文件: {file_path}")
    try:
        # 处理逻辑
        result = do_something(file_path)
        logger.info(f"文件处理成功")
        return result
    except Exception as e:
        logger.exception(f"处理文件时出错")
        raise
```

### 查看日志
- **实时查看**：运行应用时会在控制台输出日志
- **历史查看**：打开 `log/ywjlb.log` 文件
- **自动轮转**：当日志文件超过 10MB 时自动备份

## 日志级别参考

| 级别 | 用途 | 场景示例 |
|------|------|---------|
| DEBUG | 详细调试信息 | 变量值、函数调用、参数检查 |
| INFO | 一般信息流 | 处理开始/结束、生成的文件数量 |
| WARNING | 潜在问题 | 文件不存在但使用了默认值 |
| ERROR | 明确的错误 | 文件读写失败、必需列缺失 |

## 环境变量支持

可以通过环境变量控制日志行为（可选）：

```powershell
# 设置日志级别为 DEBUG
$env:LOG_LEVEL = "DEBUG"

# 设置日志文件大小限制为 5MB
$env:LOG_MAX_BYTES = "5242880"

# 设置日志备份数量
$env:LOG_BACKUP_COUNT = "10"

# 设置自定义日志文件路径
$env:LOG_FILE = "C:\logs\my_app.log"
```

## 下一步优化建议

1. **性能监控**：添加处理时间日志
   ```python
   import time
   start = time.time()
   # 处理逻辑
   logger.info(f"处理完成，耗时 {time.time() - start:.2f}s")
   ```

2. **结构化日志**：考虑使用 JSON 格式的日志便于解析
   ```python
   logger.info(f"Processed {count} files", extra={"files": count, "duration": 1.23})
   ```

3. **日志聚合**：在生产环境中考虑使用 ELK Stack 或类似的日志聚合方案

4. **错误追踪**：集成错误追踪服务如 Sentry
