from os.path import exists

from preprocessing import preprocessing
from utils.caching import initial_cache, save_cache, clear_cache, load_cache
from utils.data import search_data_by_barcode, write_to_output, data_file_path
from utils.log_formatter import setup_logger

logger = setup_logger(__name__)


def main():
    cache_content = initial_cache()
    logger.info(f'已读取 {len(cache_content)} 条缓存。')

    if not exists(data_file_path):
        logger.info('正在预处理数据。')
        preprocessing()

    while True:
        command = input('请输入 NEW_BARCODE （或输入 "End" 终止并将已缓存的数据写入输出表）：')
        if command == 'End':
            if len(load_cache()) > 0:
                write_to_output(load_cache())
                clear_cache()
            else:
                logger.warning('缓存为空，跳过 Excel 写入。')
            break
        else:
            data_content, data_index_list = search_data_by_barcode(command)
            if len(data_index_list) == 0:
                logger.error('查询失败，未找到该 NEW_BARCODE 对应的数据行。')
                continue
            elif len(data_index_list) == 1:
                cache_content += data_index_list
                logger.info(f'{command} 成功写入缓存。')
            else:
                logger.info(f'该 NEW_BARCODE 对应 {len(data_index_list)} 个数据行：')
                print(data_content)
                while True:
                    data_index_list_as_str = ', '.join([str(index) for index in data_index_list])
                    index_input = input(f'请输入你需要的数据行的索引号 ({data_index_list_as_str})：')
                    try:
                        index_input = int(index_input)
                    except (TypeError, ValueError):
                        logger.error('输入不合法，索引号必须为整数。')
                        continue
                    if index_input not in data_index_list:
                        logger.error(f'输入不在范围 ({data_index_list_as_str}) 之内。')
                        continue
                    else:
                        cache_content.append(index_input)
                        logger.info(f'{command} 成功写入缓存。')
                        break
        save_cache(cache_content)


if __name__ == '__main__':
    main()
