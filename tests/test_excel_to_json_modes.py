#!/usr/bin/env python
import sys
from pathlib import Path
import pandas as pd
import json

# Make project root importable
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from json_to_excel.json_to_excel_pyqt import ExcelToJSONWorker

BASE = Path(__file__).resolve().parent
excel_path = BASE / 'test_sample.xlsx'

# create sample Excel
df1 = pd.DataFrame({'id':[1,2], 'name':['Alice','Bob']})
df2 = pd.DataFrame({'code':[100,200], 'value':[3.14, 2.71]})
with pd.ExcelWriter(excel_path) as w:
    df1.to_excel(w, sheet_name='users', index=False)
    df2.to_excel(w, sheet_name='metrics', index=False)
print(f'WROTE_EXCEL:{excel_path}')

modes = ['always_array', 'always_object']
for mode in modes:
    out_path = BASE / f'test_sample_out_{mode}.json'
    worker = ExcelToJSONWorker(str(excel_path), str(out_path), logger=None, output_mode=mode)
    worker.run()
    with open(out_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f'MODE={mode} -> WROTE_JSON:{out_path}')
    print(json.dumps(data, ensure_ascii=False, indent=2))

print('DONE')
