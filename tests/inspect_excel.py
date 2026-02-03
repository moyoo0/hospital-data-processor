import pandas as pd

# 读取 Excel 文件
file_path = 'excels/data_aggregation/全院收入_按科室2024年01-11月门诊开单科室发票项目收入汇总.xls'
df = pd.read_excel(file_path, header=None)

# 打印前5行
print(df.head())
