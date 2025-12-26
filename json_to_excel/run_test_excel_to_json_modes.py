import sys
from pathlib import Path
import pandas as pd
import json

sys.path.insert(0, r'd:\学习\PythonProject\test')
from json_to_excel.json_to_excel_pyqt import ExcelToJSONWorker

BASE = Path(__file__).resolve().parent
excel_path = BASE / 'test_sample.xlsx'

# create sample Excel if missing
if not excel_path.exists():
    df1 = pd.DataFrame({'id':[1,2], 'name':['Alice','Bob']})
    df2 = pd.DataFrame({'code':[100,200], 'value':[3.14, 2.71]})
    with pd.ExcelWriter(excel_path) as w:
        df1.to_excel(w, sheet_name='users', index=False)
        df2.to_excel(w, sheet_name='metrics', index=False)
    print(f'WROTE_EXCEL:{excel_path}')

modes = ['always_array', 'always_object']
results = {}
for mode in modes:
    out_path = BASE / f'test_sample_out_{mode}.json'
    worker = ExcelToJSONWorker(str(excel_path), str(out_path), logger=None, output_mode=mode)
    worker.run()
    with open(out_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    results[mode] = data
    print(f'MODE={mode} -> WROTE_JSON:{out_path}')
    print(json.dumps(data, ensure_ascii=False, indent=2))

# simple diff summary
print('\nDIFFERENCE SUMMARY:')
if isinstance(results['always_array'], list):
    print('- always_array is a list with', len(results['always_array']), 'items')
else:
    print('- always_array type:', type(results['always_array']).__name__)

if isinstance(results['always_object'], dict):
    print('- always_object is an object with keys:', list(results['always_object'].keys()))
else:
    print('- always_object type:', type(results['always_object']).__name__)
