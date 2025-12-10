"""核心转换函数：解析 JSON、扁平化、拆表、流式写入等。可被 CLI 与 GUI 导入使用。"""
from pathlib import Path
from typing import Any, Dict, List, Tuple

import json
import pandas as pd


def normalize_top_items(data: Any) -> List[Dict]:
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


def split_parent_children(items: List[Dict], parent_id_cols: List[str] = None, split_fields: List[str] = None) -> Tuple[List[Dict], Dict[str, List[Dict]]]:
    parent_rows: List[Dict] = []
    children: Dict[str, List[Dict]] = {}

    for idx, item in enumerate(items):
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
            # 如果指定了 split_fields，则只拆这些字段
            if split_fields is not None and k not in split_fields:
                parent[k] = v
                continue

            if isinstance(v, list):
                if v and isinstance(v[0], dict):
                    rows = []
                    for child in v:
                        r = dict(child)
                        r['__parent_id'] = parent_id
                        rows.append(r)
                    children.setdefault(k, []).extend(rows)
                else:
                    parent[k] = ','.join(map(str, v)) if v else None
            else:
                parent[k] = v

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
        if hasattr(write_excel_flat, 'report') and write_excel_flat.report is not None:
            report = write_excel_flat.report
            rep_rows = [(k, json.dumps(v, ensure_ascii=False) if not isinstance(v, (str, int, float)) else v) for k, v in report.items()]
            pd.DataFrame(rep_rows, columns=['key', 'value']).to_excel(writer, sheet_name='report', index=False)


def write_excel_multi(parent_df: pd.DataFrame, children: Dict[str, pd.DataFrame], output_path: Path, raw_json_text: str = None) -> None:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        parent_df.to_excel(writer, sheet_name='data', index=False)
        for name, df in children.items():
            sheet_name = str(name)[:31].replace('/', '_').replace('\\', '_')
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        if raw_json_text is not None:
            pd.DataFrame({'raw_json': [raw_json_text]}).to_excel(writer, sheet_name='raw_json', index=False)
        if hasattr(write_excel_multi, 'report') and write_excel_multi.report is not None:
            report = write_excel_multi.report
            rep_rows = [(k, json.dumps(v, ensure_ascii=False) if not isinstance(v, (str, int, float)) else v) for k, v in report.items()]
            pd.DataFrame(rep_rows, columns=['key', 'value']).to_excel(writer, sheet_name='report', index=False)


def convert_flat(items: List[Dict], output_path: Path, numeric_cols=None, date_cols: List[str] = None, dedupe_by: List[str] = None, raw_text: str = None) -> pd.DataFrame:
    df = pd.json_normalize(items, sep='.')
    df = coerce_numeric_columns(df, numeric_cols=numeric_cols)
    df = parse_date_columns(df, date_cols or [])
    if dedupe_by:
        exist_cols = [c for c in dedupe_by if c in df.columns]
        if exist_cols:
            df = df.drop_duplicates(subset=exist_cols, keep='first')
    # 构建报告
    report = {
        'row_count': int(len(df)),
        'columns': list(map(str, df.columns.tolist())),
        'missing_counts': {str(k): int(v) for k, v in df.isna().sum().to_dict().items()},
    }
    # numeric sums
    if numeric_cols is None:
        candidates = [c for c in df.columns if c.upper() in {'AMT', 'GPAMT', 'AMOUNT', 'AMT_TOTAL', 'AMT_SUM'}]
    else:
        candidates = [c for c in numeric_cols if c in df.columns]
    numeric_sums = {}
    for col in candidates:
        try:
            numeric_sums[col] = float(df[col].sum(skipna=True))
        except Exception:
            numeric_sums[col] = None
    report['numeric_sums'] = numeric_sums

    # 暂存报告到 write 函数属性以便写入
    write_excel_flat.report = report
    try:
        write_excel_flat(df, Path(output_path), raw_json_text=raw_text)
    finally:
        # 清理属性
        try:
            delattr(write_excel_flat, 'report')
        except Exception:
            write_excel_flat.report = None

    return df, report


def convert_multi(items: List[Dict], output_path: Path, numeric_cols=None, date_cols: List[str] = None, dedupe_by: List[str] = None, split_fields: List[str] = None, raw_text: str = None) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    parent_id_candidates = ['BILLID', 'BILLSEQ', 'id', 'ID']
    parent_rows, children = split_parent_children(items, parent_id_cols=parent_id_candidates, split_fields=split_fields)
    parent_df = pd.json_normalize(parent_rows, sep='.')
    parent_df = coerce_numeric_columns(parent_df, numeric_cols=numeric_cols)
    parent_df = parse_date_columns(parent_df, date_cols or [])
    if dedupe_by:
        exist_cols = [c for c in dedupe_by if c in parent_df.columns]
        if exist_cols:
            parent_df = parent_df.drop_duplicates(subset=exist_cols, keep='first')

    children_dfs: Dict[str, pd.DataFrame] = {}
    for name, rows in children.items():
        if rows:
            dfc = pd.json_normalize(rows, sep='.')
            dfc = coerce_numeric_columns(dfc, numeric_cols=numeric_cols)
            dfc = parse_date_columns(dfc, date_cols or [])
            children_dfs[name] = dfc

    # 构建报告
    report = {
        'parent_row_count': int(len(parent_df)),
        'parent_missing_counts': {str(k): int(v) for k, v in parent_df.isna().sum().to_dict().items()},
        'children_counts': {str(k): int(len(v)) for k, v in children_dfs.items()},
    }
    # numeric sums for parent
    if numeric_cols is None:
        candidates = [c for c in parent_df.columns if c.upper() in {'AMT', 'GPAMT', 'AMOUNT', 'AMT_TOTAL', 'AMT_SUM'}]
    else:
        candidates = [c for c in numeric_cols if c in parent_df.columns]
    numeric_sums = {}
    for col in candidates:
        try:
            numeric_sums[col] = float(parent_df[col].sum(skipna=True))
        except Exception:
            numeric_sums[col] = None
    report['parent_numeric_sums'] = numeric_sums

    write_excel_multi.report = report
    try:
        write_excel_multi(parent_df, children_dfs, Path(output_path), raw_json_text=raw_text)
    finally:
        try:
            delattr(write_excel_multi, 'report')
        except Exception:
            write_excel_multi.report = None

    return parent_df, children_dfs, report


def stream_to_csv(input_path: Path, output_path: Path, numeric_cols=None, date_cols: List[str] = None, batch_size: int = 5000):
    try:
        import ijson  # type: ignore
    except Exception as e:
        raise RuntimeError('流式模式需要安装 ijson') from e

    def iter_items(fp):
        try:
            for obj in ijson.items(fp, 'item'):
                yield obj
            return
        except Exception:
            fp.seek(0)
        fp.seek(0)
        for line in fp:
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except Exception:
                continue

    with input_path.open('r', encoding='utf-8') as f, open(output_path, 'w', encoding='utf-8', newline='') as csvf:
        batch = []
        first_write = True
        for item in iter_items(f):
            for rec in normalize_top_items(item):
                batch.append(rec)
                if len(batch) >= batch_size:
                    dfb = pd.json_normalize(batch, sep='.')
                    dfb = coerce_numeric_columns(dfb, numeric_cols=numeric_cols)
                    dfb = parse_date_columns(dfb, date_cols or [])
                    dfb.to_csv(csvf, index=False, header=first_write, mode='a')
                    first_write = False
                    batch = []
        if batch:
            dfb = pd.json_normalize(batch, sep='.')
            dfb = coerce_numeric_columns(dfb, numeric_cols=numeric_cols)
            dfb = parse_date_columns(dfb, date_cols or [])
            dfb.to_csv(csvf, index=False, header=first_write, mode='a')
