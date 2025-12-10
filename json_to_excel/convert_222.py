import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd


def normalize_top_items(data: Any) -> List[Dict]:
    """把常见 top-level 结构规范为 list of dict（支持 list[list[dict]]、list[dict]、dict）。

    对于 list 中有多个 dict 的子列表，也会展开这些 dict。
    """
    items: List[Dict] = []

    if isinstance(data, list):
        for el in data:
            if isinstance(el, dict):
                items.append(el)
            elif isinstance(el, list):
                for sub in el:
                    if isinstance(sub, dict):
                        items.append(sub)
    elif isinstance(data, dict):
        items.append(data)

    return items


def split_parent_children(items: List[Dict], parent_id_cols: List[str] = None) -> Tuple[List[Dict], Dict[str, List[Dict]]]:
    """将 items 划分为 parent rows（剥离 list 子字段）和多个 child 表（list-of-dict）。

    parent_id_cols: 候选的 id 字段名，用于选取现有 id 作为 parent_id（优先）。
    返回: (parent_rows, {child_key: [records...]})
    """
    parent_rows: List[Dict] = []
    children: Dict[str, List[Dict]] = {}

    for idx, item in enumerate(items):
        # 为 parent 生成 parent_id：优先使用现有列，否则生成 __parent_idx
        parent_id = None
        if parent_id_cols:
            for c in parent_id_cols:
                if c in item and item.get(c) is not None:
                    parent_id = item.get(c)
                    break
        if parent_id is None:
            parent_id = f'__parent_{idx}'

        parent = {}
        for k, v in item.items():
            if isinstance(v, list):
                # 如果是 list of dict -> child table
                if v and isinstance(v[0], dict):
                    # deepcopy-like copy of each child and attach parent_id
                    rows = []
                    for child in v:
                        r = dict(child)
                        r['__parent_id'] = parent_id
                        rows.append(r)
                    children.setdefault(k, []).extend(rows)
                else:
                    # list of primitives -> join 为字符串保留在 parent
                    parent[k] = ','.join(map(str, v)) if v else None
            else:
                parent[k] = v

        # 确保 parent_id 字段存在
        parent['__parent_id'] = parent_id
        parent_rows.append(parent)

    return parent_rows, children


def coerce_numeric_columns(df: pd.DataFrame, numeric_cols=None) -> pd.DataFrame:
    if numeric_cols is None:
        candidates = [c for c in df.columns if c.upper() in {'AMT', 'GPAMT', 'AMOUNT', 'AMT_TOTAL', 'AMT_SUM'}]
    else:
        candidates = [c for c in numeric_cols if c in df.columns]

    for col in candidates:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')

    return df


def parse_date_columns(df: pd.DataFrame, date_cols: List[str]) -> pd.DataFrame:
    if not date_cols:
        return df
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


def write_excel_flat(df: pd.DataFrame, output_path: Path, raw_json_text: str = None) -> None:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='data', index=False)
        if raw_json_text is not None:
            pd.DataFrame({'raw_json': [raw_json_text]}).to_excel(writer, sheet_name='raw_json', index=False)


def write_excel_multi(parent_df: pd.DataFrame, children: Dict[str, pd.DataFrame], output_path: Path, raw_json_text: str = None) -> None:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        parent_df.to_excel(writer, sheet_name='data', index=False)
        for name, df in children.items():
            # sheet 名称限制长度，替换不安全字符
            sheet_name = str(name)[:31].replace('/', '_').replace('\\', '_')
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        if raw_json_text is not None:
            pd.DataFrame({'raw_json': [raw_json_text]}).to_excel(writer, sheet_name='raw_json', index=False)


def main():
    parser = argparse.ArgumentParser(description='Convert JSON (like 222.json) to Excel.')
    parser.add_argument('--input', '-i', required=True, help='输入 JSON 文件路径')
    parser.add_argument('--output', '-o', required=True, help='输出 Excel 文件路径 (.xlsx)')
    parser.add_argument('--mode', choices=('flat', 'multi'), default='flat', help='转换模式：flat=单表扁平化, multi=拆分数组为多个 sheet')
    parser.add_argument('--numeric-cols', '-n', nargs='*', help='要强制转换为数值的列名')
    parser.add_argument('--date-cols', '-d', nargs='*', help='要解析为日期的列名')
    parser.add_argument('--dedupe-by', nargs='*', help='按这些列去重（保留第一个）')
    parser.add_argument('--raw-sheet', action='store_true', help='在 Excel 中保留原始 JSON（raw_json sheet）')
    parser.add_argument('--split-fields', nargs='*', help='multi 模式下只拆分这些字段（默认拆所有 list-of-dict 字段）')
    parser.add_argument('--stream', action='store_true', help='启用流式解析（适用于大文件），当前流式只支持 flat 模式并输出 CSV')
    parser.add_argument('--stream-batch', type=int, default=5000, help='流式模式下每次写入 CSV 的批次大小（默认 5000）')
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise SystemExit(f'输入文件不存在: {input_path}')

    # 如果启用流式模式：仅支持 flat 模式（写 CSV），并优先使用 ijson 或 NDJSON 行解析
    raw_text = None
    if args.stream:
        if args.mode != 'flat':
            raise SystemExit('流式模式当前仅支持 --mode flat')
        # 要求输出为 CSV 或默认改为 .csv
        if output_path.suffix.lower() != '.csv':
            csv_output = output_path.with_suffix('.csv')
            print(f'流式模式输出建议为 CSV，已改为: {csv_output}')
            output_path = csv_output
        try:
            import ijson  # type: ignore
        except Exception:
            raise SystemExit('流式模式需要安装 ijson: pip install ijson')

        # 尝试以流式解析 top-level 数组
        def stream_iter_items(fp):
            # 先尝试解析为 top-level array 的元素
            try:
                for obj in ijson.items(fp, 'item'):
                    yield obj
                return
            except Exception:
                fp.seek(0)
            # 回退到按行解析（NDJSON）
            fp.seek(0)
            for line in fp:
                line = line.strip()
                if not line:
                    continue
                try:
                    yield json.loads(line)
                except Exception:
                    continue

        with input_path.open('r', encoding='utf-8') as f:
            # 如果 raw_sheet 需要保留原始文本（小文件时可用）
            if args.raw_sheet:
                raw_text = f.read()
                f.seek(0)

            # 流式写 CSV
            batch = []
            first_write = True
            batch_size = max(100, args.stream_batch)
            with open(output_path, 'w', encoding='utf-8', newline='') as csvf:
                for item in stream_iter_items(f):
                    # normalize_top_items 接受 dict/list，item 可能是 dict 或 list
                    for rec in normalize_top_items(item):
                        batch.append(rec)
                        if len(batch) >= batch_size:
                            dfb = pd.json_normalize(batch, sep='.')
                            dfb = coerce_numeric_columns(dfb, numeric_cols=args.numeric_cols)
                            dfb = parse_date_columns(dfb, args.date_cols or [])
                            dfb.to_csv(csvf, index=False, header=first_write, mode='a')
                            first_write = False
                            batch = []
                # 写剩余
                if batch:
                    dfb = pd.json_normalize(batch, sep='.')
                    dfb = coerce_numeric_columns(dfb, numeric_cols=args.numeric_cols)
                    dfb = parse_date_columns(dfb, args.date_cols or [])
                    dfb.to_csv(csvf, index=False, header=first_write, mode='a')

            print(f'流式写入完成: {output_path}')
            return

    # 非流式（内存）解析
    with input_path.open('r', encoding='utf-8') as f:
        data = json.load(f)
        raw_text = f.read() if args.raw_sheet else None

    items = normalize_top_items(data)
    if not items:
        print('未找到记录，退出。')
        return

    # 根据模式处理
    if args.mode == 'flat':
        df = pd.json_normalize(items, sep='.')
        # 类型转换
        df = coerce_numeric_columns(df, numeric_cols=args.numeric_cols)
        df = parse_date_columns(df, args.date_cols or [])

        if args.dedupe_by:
            exist_cols = [c for c in args.dedupe_by if c in df.columns]
            if exist_cols:
                df = df.drop_duplicates(subset=exist_cols, keep='first')

        write_excel_flat(df, output_path, raw_json_text=raw_text if args.raw_sheet else None)
        print(f'已写入 Excel: {output_path} （行数: {len(df)})')
        return

    # multi 模式：拆分子数组为子表
    parent_id_candidates = ['BILLID', 'BILLSEQ', 'id', 'ID']
    parent_rows, children = split_parent_children(items, parent_id_cols=parent_id_candidates)

    parent_df = pd.json_normalize(parent_rows, sep='.')
    parent_df = coerce_numeric_columns(parent_df, numeric_cols=args.numeric_cols)
    parent_df = parse_date_columns(parent_df, args.date_cols or [])

    if args.dedupe_by:
        exist_cols = [c for c in args.dedupe_by if c in parent_df.columns]
        if exist_cols:
            parent_df = parent_df.drop_duplicates(subset=exist_cols, keep='first')

    # children -> DataFrame dict
    children_dfs: Dict[str, pd.DataFrame] = {}
    for name, rows in children.items():
        if rows:
            dfc = pd.json_normalize(rows, sep='.')
            dfc = coerce_numeric_columns(dfc, numeric_cols=args.numeric_cols)
            dfc = parse_date_columns(dfc, args.date_cols or [])
            children_dfs[name] = dfc

    write_excel_multi(parent_df, children_dfs, output_path, raw_json_text=raw_text if args.raw_sheet else None)
    print(f'已写入 Excel (multi): {output_path} （主表行数: {len(parent_df)}，子表: {list(children_dfs.keys())}）')


from json_to_excel import cli


if __name__ == '__main__':
    cli.main()
