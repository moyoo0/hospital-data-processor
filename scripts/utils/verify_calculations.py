import pandas as pd
import numpy as np

def verify_data():
    file_path = 'excels/data_aggregation/全院收入_按科室202501门诊-开单科室_带合并.xlsx'
    print(f"正在验证文件: {file_path}")

    # 读取文件，包含前两行表头
    try:
        df_raw = pd.read_excel(file_path, header=None)
    except FileNotFoundError:
        print("错误：文件未找到。")
        return

    # 解析表头
    header_ids = df_raw.iloc[0] # 第一行：分组ID (0, 01, 1...)
    header_names = df_raw.iloc[1] # 第二行：列名
    
    # 数据部分（从第3行开始）
    df_data = df_raw.iloc[2:].copy().reset_index(drop=True)
    df_data.columns = header_names # 临时赋予列名方便调试

    # 找到关键列索引
    col_dept = -1
    col_total = -1
    group_sum_cols = {} # {group_id (int): col_index}
    group_item_cols = {} # {group_id (int): [col_indices]}
    
    # 允许的科室列名
    allowed_dept_names = ['开单科室', '执行科室', '病人所在病区']

    for idx, (hid, hname) in enumerate(zip(header_ids, header_names)):
        hid_str = str(hid).strip()
        hname_str = str(hname).strip() # 去除潜在空格
        
        if hname_str in allowed_dept_names:
            col_dept = idx
        elif hname_str == '合计':
            col_total = idx
        elif len(hid_str) == 2 and hid_str.isdigit() and hid_str != '00': # 01-07 合计列
            gid = int(hid_str)
            group_sum_cols[gid] = idx
        elif len(hid_str) == 1 and hid_str.isdigit() and hid_str != '0': # 1-7 明细列
            gid = int(hid_str)
            if gid not in group_item_cols:
                group_item_cols[gid] = []
            group_item_cols[gid].append(idx)
            
    if col_dept == -1:
        print("错误：未找到科室列（开单科室/执行科室/病人所在病区）")
        return

    # 开始验证每一行
    errors = []
    
    # 转换所有相关列为数值类型
    cols_to_convert = list(group_sum_cols.values()) + [col_total]
    for cols in group_item_cols.values():
        cols_to_convert.extend(cols)
        
    for c_idx in cols_to_convert:
        df_raw.iloc[2:, c_idx] = pd.to_numeric(df_raw.iloc[2:, c_idx], errors='coerce').fillna(0)

    # 遍历数据行
    for row_idx in range(2, len(df_raw)):
        row_data = df_raw.iloc[row_idx]
        dept_name = row_data[col_dept]
        
        # 1. 验证每个分组的明细之和是否等于分组合计
        calculated_group_sums = {}
        
        for gid in range(1, 8):
            # 获取该组所有明细列的值
            item_indices = group_item_cols.get(gid, [])
            item_sum = row_data[item_indices].sum()
            
            # 获取文件中该组的合计值
            sum_col_idx = group_sum_cols.get(gid)
            if sum_col_idx is None:
                continue
                
            file_sum = row_data[sum_col_idx]
            calculated_group_sums[gid] = file_sum
            
            # 对比 (容忍浮点误差)
            if abs(item_sum - file_sum) > 0.01:
                errors.append(f"行 {row_idx+1} [{dept_name}] 组 {gid} 错误: 明细和={item_sum:.2f}, 文件值={file_sum:.2f}")

        # 2. 验证分组合计之和是否等于总合计
        total_of_groups = sum(calculated_group_sums.values())
        file_total = row_data[col_total]
        
        if abs(total_of_groups - file_total) > 0.01:
             errors.append(f"行 {row_idx+1} [{dept_name}] 总合计错误: 分组累加={total_of_groups:.2f}, 文件值={file_total:.2f}")

    if errors:
        print(f"发现 {len(errors)} 个错误：")
        for e in errors[:10]: # 只打印前10个
            print(e)
        if len(errors) > 10:
            print("...")
    else:
        print("验证通过！所有计算均正确。")

if __name__ == "__main__":
    verify_data()
