import pandas as pd
import os

def inspect_fee_files():
    # 注意文件名后面有 '1' 和 '2'
    file1 = 'excels/data_export/202501出院患者费用分析1.xls'
    file2 = 'excels/data_export/202502出院患者费用分析2.xls'
    
    for f in [file1, file2]:
        print(f"\n====== 正在检查文件: {os.path.basename(f)} ======")
        if not os.path.exists(f):
            print(f"错误: 文件不存在 -> {f}")
            continue
            
        try:
            # 读取前 5 行，不指定 header，查看原始数据分布
            df_raw = pd.read_excel(f, header=None, nrows=5)
            print(df_raw.to_string())
            
        except Exception as e:
            print(f"读取出错: {e}")

if __name__ == "__main__":
    inspect_fee_files()
