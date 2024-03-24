from os.path import exists, abspath

from pandas import read_csv, read_excel
from pandas.core.frame import DataFrame
from pywintypes import com_error
from win32com.client import DispatchEx

from utils.log_formatter import setup_logger

logger = setup_logger(__name__)

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
    """
    读取CSV文件数据。

    :return: 读取的 CSV 数据。
    """
    return read_csv(data_file_path)


def retrieve_data(index_list: list[int]) -> DataFrame:
    """
    根据给定的索引列表从CSV数据中检索数据。

    :param index_list: 索引列表。
    :return: 检索到的数据。
    """
    data_sheet = read_data()
    return data_sheet.iloc[index_list]


def search_data_by_barcode(new_barcode: str) -> (DataFrame, list[int]):
    """
    根据给定的新条形码搜索数据。

    :param new_barcode: 用于搜索的 NEW_BARCODE。
    :return: 匹配结果和索引列表。
    """
    data_sheet = read_data()
    result = data_sheet.loc[data_sheet['NEW_BARCODE'] == new_barcode.upper().strip()]
    return result, result.index.astype(int).to_list()


def get_last_changed() -> str:
    """
    检查分发输出文件是否存在。

    :return: 若存在则返回该路径，否则返回源输出文件路径。
    """
    if exists(dist_output_file_path):
        return dist_output_file_path
    else:
        return src_output_file_path


def get_exist_barcodes() -> set[str]:
    """
    读取Excel文件中的条形码列。

    :return: 一个包含所有已存在条形码的集合
    """
    origin_sheet = read_excel(get_last_changed(), sheet_name=src_output_sheet_name)
    return set(origin_sheet.loc[:, src_barcode_col_name])


def gen_script(row_index: int, script_count: int) -> str:
    """
    根据行索引和脚本计数生成 VBA 脚本，用于复制样本行格式到指定行。

    :param row_index: 行索引。
    :param script_count: 脚本计数。
    :return: 生成的 VBA 脚本。
    """
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


def write_to_output(index_list: list[int]) -> None:
    """
    根据索引列表将数据写入Excel输出文件，并根据需要调整格式。

    :param index_list: 索引列表。
    """
    data_cached = retrieve_data(index_list)
    logger.debug('启动 Excel 进程。')
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
            logger.warning(f'{data_row['NEW_BARCODE']} 已经存在，跳过写入。')
            continue
        origin_worksheet.Range(f'A{row_index}').Value = row_index - 1
        for col in cols_mapping:
            origin_worksheet.Range(f'{cols_mapping[col]}{row_index}').Value = data_row[col]
            logger.info(f'数据行 {data_row['NEW_BARCODE']} 成功写入。')
        if vba_enabled:
            try:
                logger.debug(f'生成并运行 VBA # {script_count}。')
                script_module = origin_workbook.VBProject.VBComponents.Add(1)
                script_module.CodeModule.AddFromString(gen_script(row_index, script_count))
                client.Application.Run(f'Script{script_count}')
                script_count += 1
            except com_error:
                vba_enabled = False
                logger.error('宏不可用，无法自动调整格式。')
        row_index += 1

    logger.debug('保存并退出 Excel 进程。')
    origin_workbook.SaveAs(abspath(dist_output_file_path))
    origin_workbook.Close()
    logger.info('Excel 处理完毕。')
    client.Quit()
