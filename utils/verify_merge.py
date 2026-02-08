import pandas as pd
import os
import glob
import numpy as np

def verify_merge():
    # 动态寻找最新的合并文件
    merged_files = glob.glob('excels/merged/全院收入_*.xlsx')
    if not merged_files:
        print("错误：未找到合并文件。")
        return
    
    # 按修改时间排序，取最新的
    merged_file = max(merged_files, key=os.path.getmtime)
    print(f"正在验证合并文件: {merged_file}")
    
    # 动态寻找源文件 (逻辑与 batch_merge.py 保持一致)
    input_dir = 'excels/data_aggregation'
    files_to_check = [
        os.path.join(input_dir, f) for f in os.listdir(input_dir)
        if (f.endswith('.xlsx') or f.endswith('.xls')) 
        and not f.startswith('~') and not f.startswith('.~')
    ]
    
    # 过滤不存在的文件 (虽然 listdir 出来的肯定存在，但为了代码健壮性)
    valid_files = [f for f in files_to_check if os.path.exists(f)]

    if not valid_files:
        print(f"错误：在 {input_dir} 未找到任何源文件。")
        return

    print(f"找到 {len(valid_files)} 个源文件用于验证:")
    for f in valid_files:
        print(f" - {f}")

    # 读取合并文件并填充 NaN
    try:
        df_merged = pd.read_excel(merged_file, header=[0, 1], index_col=0).fillna(0)
    except Exception as e:
        print(f"读取合并文件失败: {e}")
        return

    # 读取源文件
    source_dfs = []
    for f in valid_files:
        try:
            df = pd.read_excel(f, header=[0, 1], index_col=0)
            # 清洗制表人行
            if df.index.dtype == 'object':
                df = df[~df.index.astype(str).str.contains("制表人", na=False)]
            source_dfs.append(df.fillna(0))
        except Exception as e:
            print(f"警告: 读取源文件 {f} 失败: {e}")
            return

    # 1. 总额验证
    print(f"\n--- 总额验证 ---")
    total_source_sum = 0
    for i, df in enumerate(source_dfs):
        s_sum = df.select_dtypes(include=[np.number]).values.sum()
        total_source_sum += s_sum
        # print(f"源文件 {i+1} 总额: {s_sum:,.2f}")

    merged_sum = df_merged.select_dtypes(include=[np.number]).values.sum()
    
    print(f"所有源文件总额: {total_source_sum:,.2f}")
    print(f"合并文件总额:   {merged_sum:,.2f}")
    
    diff = abs(total_source_sum - merged_sum)
    if diff < 1.0:
        print("结果: 准确 (误差在允许范围内)")
    else:
        print(f"结果: 不匹配！差额: {diff:,.2f}")

    # 2. 抽样验证
    print(f"\n--- 抽样验证 ---")
    if source_dfs:
        df1 = source_dfs[0]
        # 找到 df1 中不为 0 的位置
        non_zero_mask = df1 > 0
        non_zero_coords = np.where(non_zero_mask)
        
        if len(non_zero_coords[0]) > 0:
            row_idx = non_zero_coords[0][0]
            col_idx = non_zero_coords[1][0]
            
            test_dept = df1.index[row_idx]
            test_col = df1.columns[col_idx]
            
            print(f"抽样坐标: 科室='{test_dept}', 项目='{test_col}'")
            
            # 计算所有源文件在该坐标的和
            calculated_sum = 0
            for df in source_dfs:
                if test_dept in df.index and test_col in df.columns:
                    calculated_sum += df.loc[test_dept, test_col]
            
            # 获取合并文件在该坐标的值
            if test_dept in df_merged.index and test_col in df_merged.columns:
                merged_val = df_merged.loc[test_dept, test_col]
            else:
                merged_val = 0
                
            print(f"源文件累加值: {calculated_sum:,.2f}")
            print(f"合并文件值:   {merged_val:,.2f}")
            
            if abs(calculated_sum - merged_val) < 0.01:
                print("结果: 准确")
            else:
                print("结果: 错误！")
        else:
            print("警告：未能在第一个源文件中找到大于 0 的数值进行抽样。")


if __name__ == "__main__":
    verify_merge()
