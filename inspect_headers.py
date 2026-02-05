import pandas as pd

files = [
    '/Users/wuzhengyang/Documents/workspace/hospital_data_scripts/excels/data_export/202501出院患者费用分析1.xls',
    '/Users/wuzhengyang/Documents/workspace/hospital_data_scripts/excels/data_export/202502出院患者费用分析2.xls'
]

for f in files:
    print(f"--- Inspecting {f} ---")
    try:
        # Read first 10 rows without header to see the structure
        df = pd.read_excel(f, header=None, nrows=10)
        print(df)
    except Exception as e:
        print(f"Error reading {f}: {e}")
    print("\n")