
import pandas as pd

def get_group_definitions(aggregation_file_path):
    """
    Reads the aggregation file to get the column group definitions.
    """
    try:
        df_agg_head = pd.read_excel(aggregation_file_path, header=None, nrows=2)
        group_map = {}
        for col_idx, group_id in enumerate(df_agg_head.iloc[0]):
            if pd.notna(group_id) and group_id != 0:
                # Group IDs are read as floats, so convert to int then to zero-padded string
                group_id_str = str(int(group_id)).zfill(2)
                column_name = df_agg_head.iloc[1, col_idx]
                if group_id_str not in group_map:
                    group_map[group_id_str] = []
                group_map[group_id_str].append(column_name)
        return group_map
    except FileNotFoundError:
        print(f"错误：聚合定义文件未找到：{aggregation_file_path}")
        return None
    except Exception as e:
        print(f"读取聚合定义文件时出错：{e}")
        return None

if __name__ == '__main__':
    agg_file = 'excels/data_aggregation/全院收入_按科室2024年01-11月门诊开单科室发票项目收入汇总.xls'
    groups = get_group_definitions(agg_file)
    if groups:
        for group_id, columns in sorted(groups.items()):
            print(f"--- Group {group_id} ---")
            print(columns)
