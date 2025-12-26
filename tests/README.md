测试说明

此测试位于 `tests/test_excel_to_json_modes.py`，用途：

- 在 `tests/` 下生成一个示例 Excel（两个 sheet：`users` 与 `metrics`）。
- 运行 `Excel->JSON` 的两种模式：`always_array`（合并为数组）和 `always_object`（保持对象按 sheet 分组）。
- 输出 JSON 文件：`test_sample_out_always_array.json` 与 `test_sample_out_always_object.json`。

运行方法（在项目根目录下执行）：

```bash
python tests/test_excel_to_json_modes.py
```

运行成功后，控制台会打印生成文件路径与 JSON 预览。
