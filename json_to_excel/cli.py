"""命令行入口，包装 `core` 中的功能为 CLI 参数。"""
from pathlib import Path
import argparse
import json

from . import core


def main(argv=None):
    parser = argparse.ArgumentParser(description='JSON -> Excel/CSV 转换（通用入口）')
    parser.add_argument('--input', '-i', required=True, help='输入 JSON 文件路径')
    parser.add_argument('--output', '-o', required=True, help='输出路径 (.xlsx 或 .csv 当使用 --stream)')
    parser.add_argument('--mode', choices=('flat', 'multi'), default='flat', help='转换模式')
    parser.add_argument('--numeric-cols', '-n', nargs='*', help='要强制转换为数值的列名')
    parser.add_argument('--date-cols', '-d', nargs='*', help='要解析为日期的列名')
    parser.add_argument('--dedupe-by', nargs='*', help='按这些列去重（保留第一个）')
    parser.add_argument('--raw-sheet', action='store_true', help='在 Excel 中保留原始 JSON')
    parser.add_argument('--split-fields', nargs='*', help='multi 模式下只拆分这些字段（默认拆所有 list-of-dict 字段）')
    parser.add_argument('--stream', action='store_true', help='启用流式解析（只支持 flat 模式并输出 CSV）')
    parser.add_argument('--stream-batch', type=int, default=5000, help='流式模式批次大小')
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise SystemExit(f'输入文件不存在: {input_path}')

    if args.stream:
        if args.mode != 'flat':
            raise SystemExit('流式模式当前仅支持 flat')
        # stream to csv
        core.stream_to_csv(input_path, output_path, numeric_cols=args.numeric_cols, date_cols=args.date_cols, batch_size=args.stream_batch)
        print(f'流式写入完成: {output_path}')
        return

    with input_path.open('r', encoding='utf-8') as f:
        data = json.load(f)
        raw_text = json.dumps(data, ensure_ascii=False) if args.raw_sheet else None

    items = core.normalize_top_items(data)
    if args.mode == 'flat':
        df, report = core.convert_flat(items, output_path, numeric_cols=args.numeric_cols, date_cols=args.date_cols, dedupe_by=args.dedupe_by, raw_text=raw_text if args.raw_sheet else None)
        print(f'已写入 Excel: {output_path} （行数: {report.get("row_count")})')
        print('报告摘要:', json.dumps(report, ensure_ascii=False))
        return

    parent_df, children, report = core.convert_multi(items, output_path, numeric_cols=args.numeric_cols, date_cols=args.date_cols, dedupe_by=args.dedupe_by, split_fields=args.split_fields, raw_text=raw_text if args.raw_sheet else None)
    print(f'已写入 Excel (multi): {output_path} （主表行数: {report.get("parent_row_count")}，子表: {list(children.keys())}）')
    print('报告摘要:', json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    main()
