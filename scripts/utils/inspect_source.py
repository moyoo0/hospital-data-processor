
import pandas as pd

def find_header_row(file_path, keyword='开单科室'):
    """
    Finds the row index of the header by looking for a keyword in the first 10 rows.
    """
    try:
        df_peek = pd.read_excel(file_path, header=None, nrows=10)
        for i, row in df_peek.iterrows():
            if any(str(cell).strip() == keyword for cell in row):
                return i
    except Exception as e:
        print(f"查找表头时发生预览错误: {e}, 将回退到默认值。")
        return 1 # Fallback
    return None

def check_source_file_columns():
    """
    Reads the source file and checks for duplicate column names.
    """
    source_file = 'excels/data_export/全院收入_按科室202501门诊-开单科室.xls'
    
    header_row_index = find_header_row(source_file)
    if header_row_index is None:
        print(f"错误: 在源文件 {source_file} 中找不到表头。")
        return

    try:
        df = pd.read_excel(source_file, header=header_row_index)

        print("--- 正在检查源文件列名 ---")
        
        # Check for duplicate columns
        column_list = df.columns.tolist()
        seen_columns = set()
        duplicate_columns = set()

        for col in column_list:
            if col in seen_columns:
                duplicate_columns.add(col)
            else:
                seen_columns.add(col)

        if duplicate_columns:
            print(f"!!! 发现重复的列名: {list(duplicate_columns)}")
        else:
            print("源文件中没有发现重复的列名。")

        print("\n--- 源文件完整列名列表 ---")
        print(column_list)

    except FileNotFoundError:
        print(f"错误：源数据文件未找到：{source_file}")
    except Exception as e:
        print(f"读取源数据文件时出错：{e}")


if __name__ == '__main__':
    check_source_file_columns()
