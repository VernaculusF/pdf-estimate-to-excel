import pandas as pd
import os

files = ['file-1.xlsx','file-2.xlsx','NMCK.xlsx','OSR.xlsx']
with open('final_scanned_check.txt', 'w', encoding='utf-8') as out:
    for fname in files:
        path = os.path.join('output', fname)
        out.write(f'=== {fname} ===\n')
        xl = pd.ExcelFile(path)
        for s in xl.sheet_names:
            df = pd.read_excel(path, sheet_name=s)
            out.write(f'  {s}: {df.shape}\n')
        out.write('\n')
