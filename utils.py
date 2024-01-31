from itertools import groupby


supported_cn_punctuation = "，。！？【】（）；：“”‘’《》、"
en_punctuation = ",.!?[]();:\"\"''<>\\"

table = str.maketrans(supported_cn_punctuation, en_punctuation)


def char_type(char):
    """获取字符的类型"""
    if '\u4e00' <= char <= '\u9fff':
        return '汉字'
    elif char in supported_cn_punctuation:
        return '中文标点'
    elif '\u0000' <= char <= '\u007f':
        return '换行' if char == '\n' else 'ASCII'
    else:
        return '其他'


def split_string(s):
    """按照类型切分字符串"""
    return [{
        'type': type,
        'string': ''.join(group),
    } for type, group in groupby(s, char_type)]


def chinese_punctuation_translate(s):
    return s.translate(table)


if __name__ == '__main__':
    str = '，。！'
    print(str.translate(table))