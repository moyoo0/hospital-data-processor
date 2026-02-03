import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

from core.config_loader import get_processor_config, parse_group_config

def process_hospital_data(src_file='excels/data_export/全院收入_按科室202503门诊-开单科室.xls', 
                          output_file='excels/data_aggregation/全院收入_按科室202503门诊-开单科室_带合并.xlsx',
                          custom_config=None):
    # 1. 确定分组映射规则
    if custom_config:
        print("使用用户自定义分组配置...")
        GROUP_SUMMARIES, ITEM_TO_GROUP_ID = parse_group_config(custom_config)
    else:
        print("使用默认分组配置...")
        GROUP_SUMMARIES, ITEM_TO_GROUP_ID = get_processor_config()

    if not GROUP_SUMMARIES or not ITEM_TO_GROUP_ID:
         print("错误: 分组配置为空或无效。")
         return False

    print(f"正在读取源文件: {src_file}")
    if not os.path.exists(src_file):
        print(f"错误: 源文件不存在 {src_file}")
        return False

    # 源文件表头在第 4 行 (index 3)
    try:
        df_src = pd.read_excel(src_file, header=3)
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return False
    
    # 清洗列名：去除前后空格
    df_src.columns = df_src.columns.astype(str).str.strip()
    
    # 动态识别科室列名 (可能是 '开单科室'、'执行科室' 或 '病人所在病区')
    dept_col = None
    for col in ['开单科室', '执行科室', '病人所在病区']:
        if col in df_src.columns:
            dept_col = col
            break
            
    if not dept_col:
        print(f"无法在源文件中找到识别列（'开单科室'、'执行科室'或'病人所在病区'）。当前列名: {df_src.columns.tolist()[:5]}...")
        return False

    # 基础列
    meta_cols = [dept_col, '合计']
    # 明细列 (过滤掉 Unnamed 和 基础列)
    detail_cols = [c for c in df_src.columns if c not in meta_cols and not str(c).startswith('Unnamed')]
    
    # 2. 计算分组合计
    group_sums = {} # {gid: pd.Series}
    for gid in range(1, 8):
        group_sums[gid] = pd.Series([0.0] * len(df_src))

    for col in detail_cols:
        # 获取该列所属组，没找到则默认归入 7 (其他)
        gid = ITEM_TO_GROUP_ID.get(col.strip(), 7)
        group_sums[gid] += pd.to_numeric(df_src[col], errors='coerce').fillna(0)

    # 3. 构造输出数据框
    # 顺序：科室 | 合计 | 01合计 | ... | 07合计 | 明细...
    final_cols_data = {
        dept_col: df_src[dept_col],
        '合计': df_src['合计']
    }
    
    # 添加 7 个合计列
    for gid in range(1, 8):
        gid_str = str(gid).zfill(2)
        col_name = GROUP_SUMMARIES[gid_str]
        final_cols_data[col_name] = group_sums[gid]
        
    # 添加所有原始明细列
    for col in detail_cols:
        final_cols_data[col] = df_src[col]

    df_final = pd.DataFrame(final_cols_data)

    # 4. 构造双层表头
    # 第一行：分组ID
    header_row_0 = [0, 0] 
    for gid in range(1, 8):
        header_row_0.append(str(gid).zfill(2))
    for col in detail_cols:
        header_row_0.append(ITEM_TO_GROUP_ID.get(col.strip(), 7))

    # 第二行：列名
    header_row_1 = df_final.columns.tolist()

    df_header = pd.DataFrame([header_row_0, header_row_1], columns=df_final.columns)

    # 5. 保存结果
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 写入双层表头
            df_header.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0)
            # 写入数据
            df_final.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=2)
            
            # 格式处理
            worksheet = writer.sheets['Sheet1']
            center_alignment = Alignment(horizontal='center', vertical='center')

            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    # 1. 设置居中对齐
                    cell.alignment = center_alignment
                    
                    # 2. 计算最大宽度
                    try:
                        if cell.value:
                            cell_value_str = str(cell.value)
                            length = len(cell_value_str.encode('gbk'))
                            if length > max_length:
                                max_length = length
                    except:
                        pass
                
                # 设置宽度 (稍微紧凑一点)
                adjusted_width = max_length + 2
                if adjusted_width > 40:
                    adjusted_width = 40
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"处理完成！成功生成：{output_file}")
        return True
    except Exception as e:
        print(f"保存失败: {e}")
        return False

if __name__ == "__main__":
    process_hospital_data()
