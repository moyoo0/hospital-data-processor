import pandas as pd
import os
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def merge_excel_files(input_dir='excels/data_aggregation', output_dir='excels/merged', output_filename=None):
    # 1. 获取文件列表
    if not os.path.exists(input_dir):
        print(f"错误: 输入目录不存在 {input_dir}")
        return None

    # 修改逻辑：匹配所有 xlsx 和 xls 文件，并排除临时文件
    files_to_process = [
        f for f in os.listdir(input_dir) 
        if (f.endswith('.xlsx') or f.endswith('.xls')) 
        and not f.startswith('~') and not f.startswith('.~')
    ]
    
    if not files_to_process:
        print(f"在 {input_dir} 未找到 Excel 文件 (.xlsx 或 .xls)。")
        return None

    files_to_process.sort() # 排序，保证顺序一致
    print(f"找到 {len(files_to_process)} 个待合并文件:")
    for f in files_to_process:
        print(f" - {f}")

    # 2. 推断或使用指定的输出文件名
    if not output_filename:
        output_name_base = "全院收入_批量合并汇总"
        # 尝试从第一个文件中提取 '202501' 这样的日期模式
        date_match = re.search(r'(20\d{2}|20\d{4})', files_to_process[0])
        if date_match:
            date_str = date_match.group(1)
            output_name_base += f"_{date_str}"
        
        output_filename = f"{output_name_base}.xlsx"
    
    # 确保文件名包含扩展名
    if not output_filename.endswith(('.xlsx', '.xls')):
        output_filename += '.xlsx'
        
    output_path = os.path.join(output_dir, output_filename)

    # 3. 逐个读取并合并
    df_total = None
    common_index_name = "科室" # 默认索引名

    for idx, filename in enumerate(files_to_process):
        file_path = os.path.join(input_dir, filename)
        print(f"[{idx+1}/{len(files_to_process)}] 正在处理: {filename}")
        
        try:
            # 读取双层表头
            df_current = pd.read_excel(file_path, header=[0, 1], index_col=0)
            
            # 清洗：移除包含“制表人”的行（通常在索引中）
            # 将索引转为字符串进行判断
            if df_current.index.dtype == 'object':
                df_current = df_current[~df_current.index.astype(str).str.contains("制表人", na=False)]

            # 记录第一个文件的索引名
            if idx == 0:
                common_index_name = df_current.index.name if df_current.index.name else "科室"
            
            # 统一索引名称，避免 Warning 或合并错乱
            df_current.index.name = common_index_name

            if df_total is None:
                df_total = df_current
            else:
                # 累加：fill_value=0 保证不存在的列被视为 0
                df_total = df_total.add(df_current, fill_value=0)
                
        except Exception as e:
            print(f"读取文件 {filename} 失败: {e}")
            return None

    if df_total is None:
        print("没有成功合并任何数据。")
        return None

    # 4. 后处理
    # 4.1 重置索引，将 '科室' 变回普通列，以便统一控制表头
    df_total.reset_index(inplace=True)
    
    # 4.2 修正列名
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
        return output_path

    except Exception as e:
        print(f"保存失败: {e}")
        return None

if __name__ == "__main__":
    merge_excel_files()