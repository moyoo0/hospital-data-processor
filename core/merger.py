import pandas as pd
import os
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def find_header_row(file_path):
    """
    寻找有效的表头行索引。
    策略：扫描前 15 行，找到包含 '科室' 或 '费' 字样的行。
    如果有多个符合条件的行（例如双层表头），通常取最下面一行作为实际列名。
    返回: (header_row_index, index_col_name)
    """
    try:
        # 读取前 20 行，不带 header
        df_preview = pd.read_excel(file_path, header=None, nrows=20)
        
        keywords = ['科室', '费', '金额', '人数', '项目', '合计']
        
        # 倒序扫描，为了找到“最下面”的一行有效表头
        # 例如双层表头：Row 0 是分组ID，Row 1 是具体费用名。我们更想要 Row 1。
        # 但要注意，数据行也可能包含“费”字（如果第一列是科室名且科室名含费），
        # 所以必须结合上下文，或者正序扫描但倾向于后出现的非数据行。
        
        # 修正策略：正序扫描。
        # 找到最后一行“看起来像表头”的行。
        # 如何定义“像表头”？通常包含关键词，且该行的后续行是数据（通常是数字）。
        
        candidate_rows = []
        
        for i, row in df_preview.iterrows():
            row_str = "".join(row.astype(str).tolist())
            # 检查是否包含关键词
            if any(k in row_str for k in keywords):
                candidate_rows.append(i)
        
        if not candidate_rows:
            return 0, "科室" # 没找到，默认第一行
        
        # 这里的 candidate_rows 包含了所有疑似表头的行。
        # 对于双层表头 (Row 0: '01', Row 1: '挂号费')，Row 0 可能不含关键词，Row 1 含。所以 Row 1 被选中。
        # 对于单层表头 (Row 0: '药品费', Row 1: NaN)，Row 0 含，Row 1 不含。
        
        # 我们取候选行中的最后一行，通常这就是最底层的列名。
        best_header_row = candidate_rows[-1]
        
        # 尝试在该行中找到具体的“科室”列名
        row_values = df_preview.iloc[best_header_row].astype(str).tolist()
        index_col_name = "科室" # 默认
        for val in row_values:
            if '科室' in val:
                index_col_name = val
                break
                
        return best_header_row, index_col_name
        
    except Exception:
        return 0, "科室"

def merge_excel_files(input_dir='excels/data_aggregation', output_dir='excels/merged', output_filename=None):
    if not os.path.exists(input_dir):
        print(f"错误: 输入目录不存在 {input_dir}")
        return None

    files_to_process = [
        f for f in os.listdir(input_dir) 
        if (f.endswith('.xlsx') or f.endswith('.xls')) 
        and not f.startswith('~') and not f.startswith('.~')
    ]
    
    if not files_to_process:
        print(f"在 {input_dir} 未找到 Excel 文件。")
        return None

    files_to_process.sort()
    
    # 使用第一个文件来确定表头位置，假设同批次文件格式一致
    first_file = os.path.join(input_dir, files_to_process[0])
    header_row, common_index_name = find_header_row(first_file)
    print(f"检测到有效表头在第 {header_row + 1} 行，主键列推测为: {common_index_name}")

    df_total = None
    column_order = [] # 用于记录列的原始顺序

    for idx, filename in enumerate(files_to_process):
        file_path = os.path.join(input_dir, filename)
        print(f"[{idx+1}/{len(files_to_process)}] 正在处理: {filename}")
        
        try:
            # ... (Existing reading logic) ...
            # 直接读取指定行作为 header
            df_current = pd.read_excel(file_path, header=header_row)
            
            # 清洗列名：去除换行、空格
            df_current.columns = [str(c).replace('\r', '').replace('\n', '').strip() for c in df_current.columns]
            
            # ... (Existing index setting logic) ...
            if common_index_name in df_current.columns:
                df_current.set_index(common_index_name, inplace=True)
            else:
                # 尝试模糊匹配
                found = False
                for col in df_current.columns:
                    if '科室' in col:
                        df_current.rename(columns={col: common_index_name}, inplace=True)
                        df_current.set_index(common_index_name, inplace=True)
                        found = True
                        break
                if not found and not df_current.empty:
                    df_current.set_index(df_current.columns[0], inplace=True)
                    df_current.index.name = common_index_name

            # 记录第一个文件的列顺序
            if idx == 0:
                # 此时 index 已经被移出 columns 了，所以记录剩下的 columns 即可
                column_order = df_current.columns.tolist()

            # 移除“制表人”行
            if df_current.index.dtype == 'object':
                df_current = df_current[~df_current.index.astype(str).str.contains("制表人", na=False)]

            # 移除无效列 (Unnamed, NaN)
            df_current = df_current.loc[:, ~df_current.columns.str.startswith('Unnamed')]
            df_current = df_current.loc[:, ~df_current.columns.str.lower().isin(['nan', 'none'])]

            # 只保留数值列，且填充0
            df_current = df_current.apply(pd.to_numeric, errors='coerce').fillna(0)

            if df_total is None:
                df_total = df_current
            else:
                df_total = df_total.add(df_current, fill_value=0)
                
        except Exception as e:
            print(f"  -> 失败: {e}")
            continue

    if df_total is None:
        return None

    # 后处理与保存
    df_total.reset_index(inplace=True)
    
    # --- 恢复列顺序 ---
    # 1. 索引列必须在第一位
    final_cols = [common_index_name]
    
    # 2. 按照第一个文件的顺序添加存在的列
    current_cols_set = set(df_total.columns)
    
    for col in column_order:
        if col in current_cols_set and col != common_index_name:
            final_cols.append(col)
            
    # 3. 添加新出现的列 (追加在后面)
    for col in df_total.columns:
        if col not in final_cols:
            final_cols.append(col)
            
    # (已移除) 特殊处理：强制移动 '合计' 到最后的逻辑已删除，以保持原文件顺序

    # 应用顺序
    df_total = df_total[final_cols]

    
    # 构造输出文件名
    if not output_filename:
        output_name_base = "合并汇总"
        # 尝试提取日期
        match = re.search(r'(20\d{4}|20\d{2})', files_to_process[0])
        if match: output_name_base += f"_{match.group(1)}"
        output_filename = f"{output_name_base}.xlsx"
    if not output_filename.endswith('.xlsx'):
        output_filename += '.xlsx'
    
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)

    print("正在保存...")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 直接输出单层表头，简单纯粹
            df_total.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 美化样式
            worksheet = writer.sheets['Sheet1']
            center_alignment = Alignment(horizontal='center', vertical='center')
            for column in worksheet.columns:
                max_len = 0
                col_letter = get_column_letter(column[0].column)
                for cell in column:
                    cell.alignment = center_alignment
                    try:
                        if cell.value:
                            l = len(str(cell.value).encode('gbk'))
                            if l > max_len: max_len = l
                    except: pass
                worksheet.column_dimensions[col_letter].width = min(max_len + 2, 40)
        
        print(f"完成! 文件已保存: {output_path}")
        return output_path
    except Exception as e:
        print(f"保存失败: {e}")
        return None

if __name__ == "__main__":
    merge_excel_files()
