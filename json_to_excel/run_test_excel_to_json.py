import sys
from pathlib import Path
import pandas as pd
import json

# ensure project root in path
sys.path.insert(0, r'd:\学习\PythonProject\test')

from json_to_excel.json_to_excel_pyqt import ExcelToJSONWorker

BASE = Path(__file__).resolve().parent
excel_path = BASE / 'test_sample.xlsx'
json_out = BASE / 'test_sample_out.json'

# create sample Excel with two sheets
df1 = pd.DataFrame({'id':[1,2], 'name':['Alice','Bob']})
df2 = pd.DataFrame({'code':[100,'200'], 'value':[3.14, 2.71]})
with pd.ExcelWriter(excel_path) as w:
    df1.to_excel(w, sheet_name='users', index=False)
    df2.to_excel(w, sheet_name='metrics', index=False)

print(f'WROTE_EXCEL:{excel_path}')

# run conversion (auto mode)
worker = ExcelToJSONWorker(str(excel_path), str(json_out), logger=None, output_mode='auto')
worker.run()  # synchronous
print(f'WROTE_JSON:{json_out}')

# show preview
with open(json_out, 'r', encoding='utf-8') as f:
    data = json.load(f)

print('JSON_PREVIEW:')
print(json.dumps(data, ensure_ascii=False, indent=2))
