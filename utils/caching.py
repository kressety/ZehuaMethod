from os import remove
from os.path import exists
from pickle import load, dump

from utils.log_formatter import setup_logger

logger = setup_logger(__name__)

cache_path = '_cache.pkl'


def load_cache() -> list[int]:
    """
    加载缓存文件。

    :return: 缓存内容。
    """
    logger.debug('加载缓存文件。')
    with open(cache_path, 'rb') as cache_file:
        return load(cache_file)


def initial_cache() -> list[int]:
    """
    初始化缓存。如果缓存文件不存在或内容不是列表类型，则创建一个空列表作为缓存。

    :return: 缓存内容。
    """
    logger.debug('初始化缓存文件。')
    try:
        if exists(cache_path):
            cache_content = load_cache()
            if not isinstance(cache_content, list):
                logger.debug('缓存文件格式错误，重新生成。')
                raise ValueError
            else:
                return cache_content
        else:
            logger.debug('缓存文件不存在，重新生成。')
            raise FileNotFoundError
    except (EOFError, FileNotFoundError, ValueError):
        logger.debug('读取已保存的缓存文件。')
        with open(cache_path, 'wb') as cache_file:
            dump([], cache_file)
        return []


def save_cache(obj: list[int]) -> None:
    """
    将新的对象列表添加到现有缓存中，并将更新后的缓存保存回文件。

    :param obj: 需要添加到缓存的整数列表。
    """
    logger.debug('写入新的缓存。')
    cache_content = load_cache()
    cache_content += obj
    cache_content = list(set(cache_content))
    with open(cache_path, 'wb') as cache_file:
        dump(cache_content, cache_file)


def clear_cache() -> None:
    """
    清除缓存。
    """
    logger.debug('清除缓存文件。')
    remove(cache_path)
