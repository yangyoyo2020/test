**JSON -> Excel/CSV 转换工具（示例：222.json）**

简介
- 本目录提供 `convert_222.py`，用于将 `222.json` 这类结构的 JSON 扁平化并导出为 Excel（或在大文件场景下导出为 CSV）。
依赖与安装

- Python 3.8+
- 推荐依赖：`pandas`, `openpyxl`, `ijson`（流式模式需要）

在项目根目录运行（推荐在虚拟环境中执行）
```powershell
pip install -r .\json_to_excel\requirements.txt
```

快速开始（命令行）

- 扁平化输出为 Excel（默认 `flat` 模式）：
```powershell
python -m json_to_excel.cli -i .\json_to_excel\222.json -o .\json_to_excel\222.xlsx
```

- multi 模式：把记录中 `list` 且元素为对象的字段拆成单独 sheet（子表带 `__parent_id`）：
```powershell
python -m json_to_excel.cli -i .\json_to_excel\222.json -o .\json_to_excel\222_multi.xlsx --mode multi --raw-sheet
```

- 流式模式（大文件）——将数据分批写入 CSV（仅支持 `flat`）：
```powershell
python -m json_to_excel.cli -i .\json_to_excel\222.json -o .\json_to_excel\222.csv --stream
```

参数详解

- `--input/-i`：输入 JSON 文件路径。
- `--output/-o`：输出文件路径，`.xlsx` 或 `.csv`（流式时建议 `.csv`）。
- `--mode`：`flat`（默认）或 `multi`。
  - `flat`：把每条记录扁平化为一行（使用 `pandas.json_normalize`）。
  - `multi`：将字段值为列表且元素为字典的字段拆为单独 sheet（子表带 `__parent_id`，主表也含 `__parent_id`）。
- `--split-fields`：multi 模式下只拆分指定字段（未指定则拆分所有符合条件的字段）。
- `--numeric-cols`：指定要尝试转换为数值的列名（例如 `AMT`）。脚本也会自动尝试识别常见金额列名。
- `--date-cols`：指定要解析为日期的列名（使用 `pandas.to_datetime`）。
- `--dedupe-by`：按列去重（保留第一次出现）。适用于主表（flat 或 multi 的 parent 表）。
- `--raw-sheet`：在输出的 Excel 中保留一个 `raw_json` sheet，包含原始 JSON 文本，便于排查或回溯。
- `--stream`：启用流式解析（需要安装 `ijson`），当前仅支持 `flat` 模式并输出 CSV。使用 `--stream-batch` 控制每批写入大小。

GUI 使用（PyQt）

- 打开应用后：
  - 点击 “1. 浏览..” 选择 JSON 文件。
  - 点击 “2. 转换为 Excel文件” 选择保存位置（默认同名 `.xlsx`）。
  - 转换过程中会弹出进度对话框并显示日志信息。
  - 转换错误会通过对话框与日志显示明确消息。

- 实现说明：GUI 使用 `json_to_excel.core` 中的 `normalize_top_items` 与 `convert_flat` 来完成转换，保留进度与日志功能，因此在大多数场景中 CLI 与 GUI 的转换结果一致。

性能建议与大文件处理

- 小至中等文件（<100MB）：直接使用 `flat` 或 `multi` 模式并输出 `.xlsx`。内存占用与 DataFrame 列数成正比。
- 大文件（>100MB 或数百万条记录）：建议使用 `--stream` 模式导出 CSV：
  - 安装 `ijson`：`pip install ijson`。
  - CSV 可分批写入，减少内存峰值；必要时再把 CSV 按需转换为 XLSX（注意：将大型 CSV 转为 XLSX 会消耗大量内存）。
- 若必须直接写入 XLSX 且文件非常大：考虑按 sheet 分块写入或使用数据库/Parquet 中间存储，再将小批量写入 XLSX。

常见问题与排查

- JSON 解析错误（JSONDecodeError）：
  - 检查文件编码（应为 UTF-8）及是否为有效 JSON。若是 NDJSON（每行 JSON），可以使用 `--stream` 或先转换为数组形式。
- 列名乱序或缺失（字段不一致）：
  - 不同记录字段不一致会产生大量 NaN。建议在转换前评估字段集合，或使用 `--split-fields`/`multi` 将复杂数组拆表。
- 金额列变为科学计数或格式错误：
  - 使用 `--numeric-cols` 指定转换列，或在 Excel 中设置单元格格式。脚本会尝试删除千分符并转为数值。
- 日期解析失败：
  - 使用 `--date-cols` 指定列名，脚本使用 `pandas.to_datetime` 自动解析；若解析不正确，请预处理或传入 `date_format`（当前脚本未实现此选项，可按需扩展）。

扩展与定制

- 如果你希望：
  - 增加 Excel 单元格格式（金额/日期显示格式）——可以使用 `xlsxwriter` 或 `openpyxl` 样式 API，脚本可以扩展以接受格式化参数。
  - 按 `BILLID` 等键自动去重并生成报表摘要——可以在 CLI 中使用 `--dedupe-by BILLID`，或我可添加摘要生成功能。
  - 将流式 CSV 自动转回 XLSX（分片写入）——可提升用户便捷性，我可以实现但会增加复杂度与运行时间。

例子

- 基本扁平化（默认）：
```powershell
python -m json_to_excel.cli -i .\json_to_excel\222.json -o .\json_to_excel\222.xlsx
```

- multi 拆表并保留原始 JSON：
```powershell
python -m json_to_excel.cli -i .\json_to_excel\222.json -o .\json_to_excel\222_multi.xlsx --mode multi --raw-sheet
```

- 流式大文件写 CSV：
```powershell
python -m json_to_excel.cli -i large_file.json -o out.csv --stream --stream-batch 10000
```

如需更多帮助

如果你希望我为你的特定 JSON 提供更精确的字段映射、添加 Excel 格式化、或把 GUI 的转换选项暴露给用户（例如选择 `split-fields`、日期/金额列等），告诉我你的偏好，我会继续实现。
