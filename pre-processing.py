from pandas import read_excel


# 定义参数
src_file_path = 'data/0323.xlsx'
src_sheet_name = '(30827)240124095839_IN_REPAIR_R'
dist_file_path = 'output/processed.csv'

# 读取文件
src_sheet = read_excel(src_file_path, sheet_name=src_sheet_name)

# 1. 保留需要的 5 列
src_sheet = src_sheet[['LEVEL_6', 'NEW_BARCODE', 'RECEIVE_DATE', 'ERROR_CODE', 'MODEL']]

# 2. 筛选 MODEL ，只保留 9S7-158 开头的数据行
src_sheet = src_sheet[src_sheet['MODEL'].str.startswith('9S7-158')]

# 3. 筛选 ERROR_CODE ，只保留
#  1. 以 NXM 开头的数据行
#  2. 等于 NXSEV 或 NXDRC 或为空的数据行，并满足 LEVEL_6 为 TW
src_sheet = src_sheet[
    (src_sheet['ERROR_CODE'].str.startswith('NXM', na=False)) |
    ((src_sheet['ERROR_CODE'].isin(['NXSEV', 'NXDRC']) | src_sheet['ERROR_CODE'].isnull()) &
     (src_sheet['LEVEL_6'] == 'TW'))
]

# 4. 重设 src_sheet 索引号
src_sheet.reset_index(drop=True, inplace=True)

# 5. 保存处理后的数据到 dist_file_path
src_sheet.to_csv(dist_file_path, index=False)
