
import pandas as pd

# Set pandas options to display all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

file_path = 'excels/data_aggregation/全院收入_按科室202501门诊-开单科室_processed.xlsx'

try:
    df = pd.read_excel(file_path)
    print("--- 生成文件的前5行 ---")
    print(df.head())
    print("\n--- 列名 ---")
    print(df.columns.tolist())
    print("\n--- 数据信息 ---")
    df.info(verbose=True, show_counts=True)


except FileNotFoundError:
    print(f"错误: 文件未找到 {file_path}")
except Exception as e:
    print(f"发生错误: {e}")
