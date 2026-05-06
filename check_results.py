import pandas as pd
from pathlib import Path
import json

output_dir = Path("output")
results = []

for excel_file in sorted(output_dir.glob("*.xlsx")):
    df = pd.read_excel(excel_file)
    results.append({
        "file": excel_file.name,
        "columns": list(df.columns),
        "rows": len(df),
        "sample": df.head(5).to_dict('records')
    })

with open("final_check_report.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("Final report saved to final_check_report.json")
