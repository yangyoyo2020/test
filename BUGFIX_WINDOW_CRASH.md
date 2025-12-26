# 程序自动关闭问题修复说明

## 问题描述
重复打开/关闭"会计核算偏离度工具"和"JSON转Excel工具"多次后，程序会自动关闭。

## 问题根源
通过代码审查发现了多个导致崩溃或异常退出的问题：

### 1. **缺少窗口关闭时的线程清理**
   - `kjhs_test/pld_pyqt6.py` 的 `MainWindow` 类没有实现 `closeEvent` 方法
   - `json_to_excel/json_to_excel_pyqt.py` 的 `JSONToExcelConverter` 类也没有实现 `closeEvent` 方法
   - 当用户关闭窗口时，后台运行的分析/转换线程（`AnalysisWorker`、`ConversionWorker`）仍然在运行
   - 线程持续访问已关闭的窗口控件（如 `log_text`），导致悬挂指针和段错误

### 2. **悬挂引用导致的异常**
   - `common/window_utils.py` 中的 `open_or_activate` 函数可能复用已被销毁的 Qt 对象
   - 对已删除对象调用方法会抛出异常，导致不稳定行为

## 修复方案

### 修改 1: `kjhs_test/pld_pyqt6.py` - 添加 `closeEvent`
```python
def closeEvent(self, event):
    """窗口关闭时清理：终止运行中的分析线程，移除 logger 的 handler"""
    # 1. 如果工作线程正在运行，调用 terminate() 并等待其退出
    # 2. 从全局 logger 中移除与本窗口绑定的 handler，避免泄露
    # 3. 安全调用父类 closeEvent
```

**关键改进：**
- 在窗口关闭前优雅终止分析线程
- 移除所有与本窗口关联的 logger handler，防止后续日志写入到已删除的控件
- 捕捉异常以确保窗口无论如何都能正确关闭

### 修改 2: `json_to_excel/json_to_excel_pyqt.py` - 添加 `closeEvent`
与修改 1 相同的模式，确保 JSON 转换工具也能正确清理转换线程。

### 修改 3: `common/window_utils.py` - 增强健壮性
1. **检测已删除的 Qt 对象**
   - 在复用窗口前，尝试调用 `isVisible()` / `isMinimized()`
   - 捕捉异常，说明对象已被删除的底层 C++ 部分

2. **自动清理悬挂引用**
   - 检测到无效对象时，立即将 `parent.<attr>` 设置为 `None`
   - 强制创建新实例而非复用

3. **连接 destroyed 信号**
   - 在新建窗口时，连接其 `destroyed` 信号
   - 窗口销毁时自动清理父对象上的引用，避免后续误用

4. **详细日志追踪**
   - 记录每个窗口的创建、复用、销毁等关键事件
   - 便于调试相似的资源泄漏问题

## 验证修复

### 手动测试（推荐）
```powershell
cd d:\学习\PythonProject\test
& .\.venv\Scripts\python.exe .\main.py
```

1. 打开"2. 会计核算偏离度工具" → 关闭
2. 打开"3. JSON转Excel工具" → 关闭
3. 重复步骤 1-2 共 3-5 次
4. 观察程序是否仍在运行且日志输出正常

### 自动化测试（可选）
```powershell
& .\.venv\Scripts\python.exe .\test_window_lifecycle.py
```
此脚本自动化执行 3 轮打开/关闭，观察日志输出是否正常。

## 预期改进
✅ **不再自动关闭** - 线程清理完全，无悬挂引用  
✅ **更清晰的日志** - 详细记录窗口生命周期，便于诊断  
✅ **更健壮的窗口管理** - 自动检测并清理已销毁的对象  

## 相关文件更改
- `common/window_utils.py` - 增强窗口管理健壮性和日志追踪
- `kjhs_test/pld_pyqt6.py` - 添加 closeEvent，清理分析线程
- `json_to_excel/json_to_excel_pyqt.py` - 添加 closeEvent，清理转换线程
- `test_window_lifecycle.py` - 新增自动化测试脚本（可选）

## 注意事项
- 如果在关闭线程时仍然出现延迟，可能是因为后台任务（如数据处理）耗时较长
  - 这种情况下，窗口会显示"未响应"，但最终仍会优雅关闭（最多等待 2000ms）
  - 这是正常行为，不是程序崩溃
