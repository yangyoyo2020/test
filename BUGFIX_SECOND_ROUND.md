# 程序自动关闭问题 - 深入诊断和修复

## 问题症状
- 重复打开/关闭会计核算工具或JSON转Excel工具后，整个应用自动退出
- 无明显的错误信息或弹窗

## 第二轮修复（深度诊断）

基于代码审查，我进行了以下深度改进：

### 1. **主程序入口加强** (`main.py`)
- ✅ 添加全局异常处理器 (`sys.excepthook`)
- ✅ 将所有异常完整记录到日志和 stderr
- ✅ 确保程序异常退出时会打印详细的 traceback

### 2. **线程异常捕获增强** (`kjhs_test/pld_pyqt6.py` 的 `AnalysisWorker`)
```python
def run(self):
    # 现在每个步骤都有细粒度的日志和异常处理
    # - 读取数据：细粒度日志 + 异常捕获
    # - 处理数据：细粒度日志 + 异常捕获
    # - 保存结果：细粒度日志 + 异常捕获
    # - 最后会记录完整的 exc_info
```

### 3. **窗口关闭安全加强** (`kjhs_test/pld_pyqt6.py` 和 `json_to_excel/json_to_excel_pyqt.py`)
```python
def closeEvent(self, event):
    # 多层嵌套异常捕获：确保任何异常都不会导致程序崩溃
    # - 第1层：顶层 try-except
    # - 第2层：logger 异常用 stderr 输出（避免循环依赖）
    # - 第3层：worker 清理异常
    # - 第4层：signal handler 清理异常
    # - 最后：super().closeEvent() 异常也被捕获
```

### 4. **数据处理异常细化** (`kjhs_test/pld_pyqt6.py`)
- ✅ `save_results()` 现在有详细的步骤日志
- ✅ `_merge_and_calculate_deviation()` 每个数据操作都有异常捕获和日志
- ✅ Excel 格式化失败不会导致程序退出（会继续返回已保存结果）

### 5. **转换工作线程异常增强** (`json_to_excel/json_to_excel_pyqt.py` 的 `ConversionWorker`)
- ✅ 读取 JSON：异常日志 + 详细错误信息
- ✅ 数据标准化：异常日志 + 详细错误信息
- ✅ Excel 转换：异常日志 + 详细错误信息
- ✅ 线程退出时确保发送信号

## 预期改进
✅ **更详细的日志** - 快速定位问题  
✅ **不会因 logger 故障而崩溃** - logger 异常用 stderr 处理  
✅ **确保 closeEvent 安全** - 多层异常保护  
✅ **完整的异常链** - 无论发生什么都能看到 traceback  

## 如何验证修复

### 方案 A：使用诊断脚本（推荐）
```powershell
cd d:\学习\PythonProject\test
& .\.venv\Scripts\python.exe .\diagnostic.py
```
此脚本自动化执行：
1. 打开会计核算工具 → 关闭
2. 再次打开会计核算工具 → 关闭
3. 重复 3 次
4. 所有日志保存到 `diagnostic.log`

**预期结果：** 程序保持运行状态，不自动关闭

### 方案 B：手动测试
```powershell
cd d:\学习\PythonProject\test
& .\.venv\Scripts\python.exe .\main.py
```
1. 点击"2. 会计核算偏离度工具"打开
2. 等待加载，点击关闭按钮或窗口的 X
3. 重复步骤 1-2 共 3-5 次
4. 观察主程序是否仍在运行

**预期结果：** 主窗口继续响应，不自动关闭

### 方案 C：查看详细日志
```powershell
# 打开日志文件查看详细信息
Get-Content .\日志文件.log -Tail 100
```

## 如果问题仍然存在

1. **运行诊断脚本** 并把 `diagnostic.log` 的内容发给我
2. **查看 stderr 输出** - 寻找任何 `[FATAL ERROR]` 或异常信息
3. **提供**：
   - 你的操作步骤（在哪一步会崩溃）
   - 任何弹窗错误信息
   - `日志文件.log` 的最后 100 行

## 技术细节

### 为什么会自动关闭？
常见原因：
1. **未捕获异常导致线程崩溃** → QApplication 接收到信号并退出
2. **logger 异常在 closeEvent 中导致循环错误** → 级联失败
3. **资源泄漏导致 Qt 内部状态不一致** → 随机崩溃

### 本次修复覆盖
- ✅ 线程中的异常（现在有详细日志和异常链）
- ✅ logger 相关异常（现在用 stderr 备用输出）
- ✅ closeEvent 中的异常（现在有多层嵌套保护）
- ✅ 数据处理异常（现在有细粒度异常捕获）

## 下一步

如果修复后仍有问题，我可以：
1. 在 `open_or_activate` 中添加更多日志来追踪窗口生命周期
2. 使用 `atexit` 钩子记录程序退出前的状态
3. 在每个信号连接处添加日志，追踪信号流
4. 使用 Python 的 `tracemalloc` 检测内存泄漏

**建议：** 请先运行诊断脚本或手动测试，如果仍有问题，请分享日志内容。
