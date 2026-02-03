import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def process_hospital_data(src_file='excels/data_export/全院收入_按科室202503门诊-开单科室.xls', 
                          output_file='excels/data_aggregation/全院收入_按科室202503门诊-开单科室_带合并.xlsx'):
    # 1. 硬编码分组映射规则
    # 合计列定义 (组ID -> 合计显示名称)
    GROUP_SUMMARIES = {
        "01": "医疗服务性收入合计",
        "02": "中医医疗服务性收入合计",
        "03": "检查收入合计",
        "04": "西成药收入合计",
        "05": "中草药合计",
        "06": "材料费合计",
        "07": "其他收入合计"
    }

    # 明细项归属映射 (明细列名 -> 组ID 1-7)
    ITEM_TO_GROUP_ID = {
        # 组 1
        '产科手术': 1, '肠镜下电切术': 1, '穿刺': 1, '妇科手术': 1, '骨科处置费': 1, '挂号费': 1, 
        '护理费': 1, '会诊费': 1, '监护病房费': 1, '接生费': 1, '介入治疗': 1, '康复治疗': 1, 
        '麻醉费': 1, '母婴同室费': 1, '其他费': 1, '手术费': 1, '手术仪器费': 1, '碎石费': 1, 
        '特殊器械': 1, '胃镜下电切术': 1, '五官处置费': 1, '五官手术费': 1, '五官治疗费': 1, 
        '镶牙': 1, '血液透析': 1, '氧气费': 1, '院前急救费': 1, '诊查费': 1, '诊疗费': 1, 
        '镇疼费': 1, '治疗费': 1, '住院费': 1, '注射费': 1,
        # 组 2
        '煎药费': 2, '理疗费': 2, '针灸费': 2, '中医治疗费': 2,
        # 组 3
        'B超费': 3, 'CT费': 3, 'DR费': 3, '病理费': 3, '查体费': 3, '磁共振费': 3, '多普勒': 3, 
        '放射费': 3, '高压氧': 3, '化验费': 3, '检查费': 3, '结肠镜': 3, '脑电图': 3, '特检费': 3, 
        '胃肠透视': 3, '胃镜费': 3, '心电图': 3, '支气管镜检查': 3,
        # 组 4
        '西药费': 4, '疫苗费': 4, '中成药': 4,
        # 组 5
        '中草药': 5,
        # 组 6
        '材料费': 6, '接种服务材料费': 6, '手术材料费': 6,
        # 组 7
        '病历复印费': 7, '工本费': 7, '血费': 7
    }

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
