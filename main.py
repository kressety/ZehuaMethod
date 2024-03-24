from utils.caching import initial_cache, save_cache, clear_cache, load_cache
from utils.data import search_data_by_barcode, write_to_output


def main():
    cache_content = initial_cache()
    print(f'已读取 {len(cache_content)} 条缓存。')

    while True:
        command = input('请输入 NEW_BARCODE （或输入 "End" 终止并将已缓存的数据写入输出表）：')
        if command == 'End':
            write_to_output(load_cache())
            clear_cache()
            break
        else:
            data_content, data_index_list = search_data_by_barcode(command)
            if len(data_index_list) == 0:
                print('查询失败，未找到该 NEW_BARCODE 对应的数据行。')
                continue
            elif len(data_index_list) == 1:
                cache_content += data_index_list
                print(f'{command} 成功写入缓存。')
            else:
                print(f'该 NEW_BARCODE 对应 {len(data_index_list)} 个数据行：')
                print(data_content)
                while True:
                    data_index_list_as_str = ', '.join([str(index) for index in data_index_list])
                    index_input = input(f'请输入你需要的数据行的索引号 ({data_index_list_as_str})：')
                    try:
                        index_input = int(index_input)
                    except (TypeError, ValueError):
                        print('输入不合法，索引号必须为整数。')
                        continue
                    if index_input not in data_index_list:
                        print(f'输入不在范围 ({data_index_list_as_str}) 之内。')
                        continue
                    else:
                        cache_content.append(index_input)
                        print(f'{command} 成功写入缓存。')
                        break
        save_cache(cache_content)


if __name__ == '__main__':
    main()
