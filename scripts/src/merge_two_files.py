import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def merge_two_files():
    # === 配置区域 ===
    output_dir = 'excels/merged'
    output_filename = '全院收入_按科室202501-02门诊-开单科室.xlsx'
    
    # 硬编码的两个文件路径
    file_path_1 = 'excels/data_aggregation/全院收入_按科室202501门诊-开单科室_processed.xlsx'
    file_path_2 = 'excels/data_aggregation/全院收入_按科室202502门诊-开单科室_processed.xlsx'
    # =============

    files_to_process = [file_path_1, file_path_2]
    
    print(f"准备合并以下两个文件:")
    for f in files_to_process:
        print(f" - {f}")
        if not os.path.exists(f):
            print(f"错误: 文件不存在 {f}")
            return

    output_path = os.path.join(output_dir, output_filename)

    # 逐个读取并合并
    df_total = None
    common_index_name = "科室" # 默认索引名

    for idx, file_path in enumerate(files_to_process):
        print(f"[{idx+1}/2] 正在处理: {file_path}")
        
        try:
            # 读取双层表头
            df_current = pd.read_excel(file_path, header=[0, 1], index_col=0)
            
            # 清洗：移除包含“制表人”的行
            if df_current.index.dtype == 'object':
                df_current = df_current[~df_current.index.astype(str).str.contains("制表人", na=False)]

            # 记录第一个文件的索引名
            if idx == 0:
                common_index_name = df_current.index.name if df_current.index.name else "科室"
            
            # 统一索引名称
            df_current.index.name = common_index_name

            if df_total is None:
                df_total = df_current
            else:
                # 累加
                df_total = df_total.add(df_current, fill_value=0)
                
        except Exception as e:
            print(f"读取文件 {file_path} 失败: {e}")
            return

    if df_total is None:
        print("没有成功合并任何数据。")
        return

    # 后处理
    # 重置索引
    df_total.reset_index(inplace=True)
    
    # 修正列名
    new_columns = []
    for col in df_total.columns:
        if col == df_total.columns[0]: 
             new_columns.append((0, common_index_name))
        else:
            new_columns.append(col)
            
    df_total.columns = pd.MultiIndex.from_tuples(new_columns)

    # 自定义列排序逻辑
    def get_sort_key(col_tuple):
        group_id_str = str(col_tuple[0])
        col_name = str(col_tuple[1])
        
        priority = 99
        sub_priority = group_id_str 
        
        if group_id_str == '0' or group_id_str == '00' or group_id_str == '0':
            priority = 0
            if col_name == '合计':
                sub_priority = 'B' 
            else:
                sub_priority = 'A'
        elif len(group_id_str) == 2 and group_id_str.isdigit(): 
            priority = 1
        elif len(group_id_str) == 1 and group_id_str.isdigit():
            priority = 2
        
        return (priority, sub_priority, col_name)

    # 获取所有列，按自定义 Key 排序
    sorted_cols = sorted(df_total.columns, key=get_sort_key)
    df_total = df_total[sorted_cols]

    print("正在保存并调整格式...")
    os.makedirs(output_dir, exist_ok=True)
    
    # 手动构造表头
    header_row_0 = [col[0] for col in df_total.columns]
    header_row_1 = [col[1] for col in df_total.columns]
    
    df_header = pd.DataFrame([header_row_0, header_row_1])

    # 临时扁平化列名
    df_total.columns = range(len(df_total.columns))

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 写入表头
            df_header.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0)
            
            # 2. 写入数据
            df_total.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=2)
            
            # 样式调整
            worksheet = writer.sheets['Sheet1']
            center_alignment = Alignment(horizontal='center', vertical='center')

            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    cell.alignment = center_alignment
                    try:
                        if cell.value:
                            cell_str = str(cell.value)
                            length = len(cell_str.encode('gbk'))
                            if length > max_length:
                                max_length = length
                    except:
                        pass
                
                adj_width = max_length + 2
                if adj_width > 40:
                    adj_width = 40
                worksheet.column_dimensions[column_letter].width = adj_width

        print(f"全部完成！合并文件已保存至: {output_path}")

    except Exception as e:
        print(f"保存失败: {e}")

if __name__ == "__main__":
    merge_two_files()
