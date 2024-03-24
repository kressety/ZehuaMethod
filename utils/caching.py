from os import remove
from os.path import exists
from pickle import load, dump

cache_path = '_cache.pkl'


def load_cache() -> list[int]:
    """
    加载缓存文件。

    :return: 缓存内容。
    """
    with open(cache_path, 'rb') as cache_file:
        return load(cache_file)


def initial_cache() -> list[int]:
    """
    初始化缓存。如果缓存文件不存在或内容不是列表类型，则创建一个空列表作为缓存。

    :return: 缓存内容。
    """
    try:
        if exists(cache_path):
            cache_content = load_cache()
            if not isinstance(cache_content, list):
                raise ValueError("Cache content is not a list")
            else:
                return cache_content
        else:
            raise FileNotFoundError("Cache file does not exist")
    except (EOFError, FileNotFoundError, ValueError):
        with open(cache_path, 'wb') as cache_file:
            dump([], cache_file)
        return []


def save_cache(obj: list[int]) -> None:
    """
    将新的对象列表添加到现有缓存中，并将更新后的缓存保存回文件。

    :param obj: 需要添加到缓存的整数列表。
    """
    cache_content = load_cache()
    cache_content += obj
    cache_content = list(set(cache_content))
    with open(cache_path, 'wb') as cache_file:
        dump(cache_content, cache_file)


def clear_cache() -> None:
    """
    清除缓存。
    """
    remove(cache_path)
