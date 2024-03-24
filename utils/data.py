from os.path import exists, abspath

from pandas import read_csv, read_excel
from pandas.core.frame import DataFrame
from pywintypes import com_error
from win32com.client import DispatchEx

data_file_path = 'output/processed.csv'
src_output_file_path = 'data/REPAIR_RPT_P1_MS-158X_0313.xlsx'
src_output_sheet_name = '158x_REPAIR_RPT_三月份流水記錄 '
src_barcode_col_name = 'NEW_BARCODE（機台條碼）'
dist_output_file_path = 'output/REPAIR_RPT_P1_MS-158X_0313.xlsx'

cols_range = [chr(i) for i in range(65, 75)]
cols_mapping = {
    'LEVEL_6': 'F',
    'NEW_BARCODE': 'B',
    'RECEIVE_DATE': 'G',
    'MODEL': 'C'
}

sample_row_index = 77


def read_data() -> DataFrame:
    return read_csv(data_file_path)


def retrieve_data(index_list: list[int]) -> DataFrame:
    data_sheet = read_data()
    return data_sheet.iloc[index_list]


def search_data_by_barcode(new_barcode: str) -> (DataFrame, list[int]):
    data_sheet = read_data()
    result = data_sheet.loc[data_sheet['NEW_BARCODE'] == new_barcode.upper().strip()]
    return result, result.index.astype(int).to_list()


def get_last_changed() -> str:
    if exists(dist_output_file_path):
        return dist_output_file_path
    else:
        return src_output_file_path


def get_exist_barcodes() -> set[str]:
    origin_sheet = read_excel(get_last_changed(), sheet_name=src_output_sheet_name)
    return set(origin_sheet.loc[:, src_barcode_col_name])


def gen_script(row_index: int, script_count: int) -> str:
    script = f'''
    Sub Script{str(script_count)}()
        Range("A{str(sample_row_index)}:J{str(sample_row_index)}").Select
        Selection.Copy
        Range("A{row_index}").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End Sub
    '''
    return script


def write_to_output(index_list: list[int]):
    data_cached = retrieve_data(index_list)
    client = DispatchEx('Excel.Application')
    client.Visible = False
    client.DisplayAlerts = False
    vba_enabled = True
    script_count = 1

    origin_workbook = client.Workbooks.Open(abspath(get_last_changed()), ReadOnly=False)
    origin_worksheet = origin_workbook.Worksheets(src_output_sheet_name)

    row_index = 1
    while origin_worksheet.Range(f'A{row_index}').Value is not None:
        row_index += 1

    for data_index in range(len(data_cached)):
        data_row = data_cached.iloc[data_index]
        if data_row['NEW_BARCODE'] in get_exist_barcodes():
            print(f'{data_row['NEW_BARCODE']} 已经存在，跳过写入。')
            continue
        origin_worksheet.Range(f'A{row_index}').Value = row_index - 1
        for col in cols_mapping:
            origin_worksheet.Range(f'{cols_mapping[col]}{row_index}').Value = data_row[col]
        if vba_enabled:
            try:
                script_module = origin_workbook.VBProject.VBComponents.Add(1)
                script_module.CodeModule.AddFromString(gen_script(row_index, script_count))
                client.Application.Run(f'Script{script_count}')
                script_count += 1
            except com_error:
                vba_enabled = False
        row_index += 1

    origin_workbook.SaveAs(abspath(dist_output_file_path))
    origin_workbook.Close()
    if vba_enabled is False:
        print('宏不可用，无法自动调整格式。')
    client.Quit()
