# 第二次打开/关闭导致程序退出 - 修复方案

## 问题症状
- 第一次打开会计核算工具：✓ 正常
- 第一次关闭会计核算工具：✓ 正常  
- 第二次打开会计核算工具：✓ 正常
- **第二次关闭会计核算工具：✗ 程序自动退出**

## 根本原因分析

通过代码审查，发现了关键问题：

### 1. **信号处理器泄漏**
在 `MainWindow.__init__` 中：
```python
add_qt_signal(self.logger, self.log_signal)  # 每次创建窗口都添加一个新的 handler
self.log_signal.connect(self.update_log)
```

在 `closeEvent` 中虽然尝试移除 handler，但：
- 只检查 `isinstance(h, QtSignalHandler) and getattr(h, 'signal', None) is self.log_signal`
- 如果判断条件不准确，handler 没有被移除
- 第二次打开时又添加一个新的 handler，导致**双重连接问题**

### 2. **destroyed 信号的闭包问题**
之前在 `window_utils.py` 中连接 `destroyed` 信号，闭包中捕获了 `logger` 和 `attr_name` 变量：
```python
def _on_destroyed(obj=None):
    setattr(parent, attr_name, None)  # 这个闭包可能在错误的上下文执行
```

当第二次窗口销毁时，闭包可能执行异常的代码路径。

## 修复方案

### 修改 1: `window_utils.py` - 移除 destroyed 信号连接
**理由**：
- `closeEvent` 已经完全负责资源清理
- `destroyed` 信号在二次事件循环中可能导致不稳定
- 父窗口持有强引用，子窗口销毁时自动管理生命周期

```python
# 之前的代码（会导致问题）：
inst.destroyed.connect(_on_destroyed)  # 移除这行

# 新的代码：不连接 destroyed 信号
# closeEvent 已经足够了
```

### 修改 2: `closeEvent` 中激进地清理所有 handler
**改进点**：
- 不再检查 `getattr(h, 'signal', None) is self.log_signal`（这个条件可能不准确）
- 改为 **激进地移除所有 QtSignalHandler**
- 额外添加 `self.log_signal.disconnect()` 断开所有信号连接
- 清空 `self.worker` 和 `self.analyzer` 引用

```python
# 新的逻辑：
# 1. 断开 log_signal 的所有连接
if getattr(self, 'log_signal', None):
    try:
        self.log_signal.disconnect()
    except Exception:
        pass

# 2. 激进地移除所有 QtSignalHandler（不检查信号名）
for h in list(self.logger.handlers):
    if isinstance(h, QtSignalHandler):
        self.logger.removeHandler(h)

# 3. 清空所有引用
self.analyzer = None
self.worker = None
```

## 验证修复

### 方案 A：运行最小化测试（推荐）
```powershell
cd d:\学习\PythonProject\test
& .\.venv\Scripts\python.exe .\test_second_close.py
```

此脚本执行：
1. 打开会计核算工具 (第 1 秒)
2. 关闭会计核算工具 (第 3 秒)
3. 打开会计核算工具 (第 5 秒)
4. 关闭会计核算工具 (第 7 秒)

**预期结果**：
```
[步骤 1] 第一次打开会计核算工具...
[OK] 窗口已打开

[步骤 2] 第一次关闭会计核算工具...
[OK] 窗口已关闭

[步骤 3] 第二次打开会计核算工具...
[OK] 窗口已打开（如果程序仍在这里，说明没有自动关闭！）

[步骤 4] 第二次关闭会计核算工具...
[OK] 窗口已关闭

[SUCCESS] 测试成功！程序没有自动关闭
```

### 方案 B：手动验证
```powershell
& .\.venv\Scripts\python.exe .\main.py
```
1. 点击"2. 会计核算偏离度工具" → 关闭
2. 点击"2. 会计核算偏离度工具" → 关闭
3. 观察：程序是否继续运行（预期：继续运行）

## 技术细节

### 为什么激进地移除所有 QtSignalHandler？

在之前的代码中，判断条件是：
```python
if isinstance(h, QtSignalHandler) and getattr(h, 'signal', None) is self.log_signal:
```

这个条件有问题：
- **`is` 比较的是对象 ID**，不是比较内容相等
- 如果闭包或其他因素导致 `self.log_signal` 的身份发生变化，条件会失败
- 即使 handler 是为了该信号添加的，也可能无法识别

新方法：
```python
if isinstance(h, QtSignalHandler):  # 不再检查信号对象
    self.logger.removeHandler(h)
```

**更加可靠**，因为：
1. 对于同一个窗口实例，通常只有一个 QtSignalHandler
2. 即使有多个，都应该在窗口关闭时移除（避免泄露）

## 关键改进总结

| 改进 | 文件 | 效果 |
|------|------|------|
| 移除 destroyed 信号 | `common/window_utils.py` | 避免闭包陷阱和二次事件循环问题 |
| 激进移除所有 handler | `kjhs_test/pld_pyqt6.py` + `json_to_excel/json_to_excel_pyqt.py` | 避免双重连接和 handler 泄露 |
| 断开 log_signal 连接 | `closeEvent` | 确保信号不再路由到已关闭的窗口 |
| 清空引用 | `closeEvent` | 明确表示窗口已销毁，避免悬挂指针 |

## 预期结果

✅ **第一次打开/关闭** - 正常  
✅ **第二次打开/关闭** - 正常  
✅ **多次重复** - 稳定  
✅ **无信号泄露** - 每次关闭都完全清理  

---

**如果修复后仍有问题**，请运行 `test_second_close.py` 并报告在哪一步失败。
