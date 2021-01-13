# -*- coding=utf-8 -*-
"""文档处理

"""
__author__ = "闭锦秀"

from docx import Document
import re
import copy

# 正文索引字典
indexDir = {}
# 参考文献字典
quotesDir = {}
# 记录参考文献列表的字体
quoteFontName = None
quoteFontSize = None


# 创建文档对象
def create_document(path):
    document = Document(path)
    return document


# 给正文中的索引重新排序
def index_sort(document):
    global indexDir
    # 设定一个num, 记录是第几个参考文献引用
    num = 0
    for paragraph in document.paragraphs:
        # 获取正文中的引用
        indexs = re.findall(r"(\w)+(\[\d+\])", paragraph.text)
        # 获取有索引的段落, 进行操作, 没有索引的段落忽略
        if len(indexs) >= 1:
            # 获取本段的字体
            font_size = paragraph.runs[0].font.size
            font_name = paragraph.runs[0].font.name
            # 判断这一段中的引用个数是不是>1, 如果>1, 则逐句拆开进行修改
            if len(indexs) == 1:
                # 获取原文中引用的标号, 这里是提取中括号中的标号
                indexNum = indexs[0][1]
                # 判断字典中是否存在此标号
                if indexNum in indexDir:
                    # 修改前的标号
                    old = indexNum
                    # 修改后的标号
                    new = "[" + str(indexDir[indexNum]) + "]"
                    # 替换
                    paragraph.text = paragraph.text.replace(old, new)
                else:
                    # 序号+1
                    num += 1
                    # 先将未在字典中的索引标号加入字典
                    indexDir[indexNum] = num
                    # 修改前的标号
                    old = indexNum
                    # 修改后的标号
                    new = "[" + str(indexDir[indexNum]) + "]"
                    # 替换
                    paragraph.text = paragraph.text.replace(old, new)
                    # 替换参考文献列表
                    quotesDir[indexNum].text = quotesDir[indexNum].text.replace(old, new)

            elif len(indexs) > 1:
                sentenses = paragraph.text.split("，")  # 如果是英文，注意使用英文标点
                text = ""
                for sentense in sentenses:
                    index = re.findall(r"(\w)+(\[\d+\])", sentense)
                    if len(index) > 0:
                        # 获取原文中引用的标号, 这里是提取中括号中的标号
                        indexNum = index[0][1]
                        # 判断字典中是否存在此标号
                        if indexNum in indexDir:
                            # 修改前的标号
                            old = indexNum
                            # 修改后的标号
                            new = "[" + str(indexDir[indexNum]) + "]"
                            sentense = sentense.replace(old, new)
                        else:
                            # 序号+1
                            num += 1
                            # 先将未在字典中的索引标号加入字典
                            indexDir[indexNum] = num
                            # 修改前的标号
                            old = indexNum
                            # 修改后的标号
                            new = "[" + str(indexDir[indexNum]) + "]"
                            sentense = sentense.replace(old, new)
                            # 替换参考文献列表
                            quotesDir[indexNum].text = quotesDir[indexNum].text.replace(old, new)

                    # if re.search(r"(\w)+(\[\d+\])", sentense):
                    #     partern = re.compile("\[\d\]")
                    #     new = "[" + str(num) + "]"
                    #     sentense = partern.sub(new, sentense)
                    #     num += 1
                    text = text + "，" + sentense
                paragraph.text = text

            # 修改字体
            repair_font_size(paragraph, font_name, font_size)


# 获取文末的参考文献
def quote_get(document):
    global quotesDir
    # 字体
    global quoteFontName
    global quoteFontSize

    # 遍历段落, 查找参考文献列表
    for p in document.paragraphs:
        flag = re.match(r"^(\[\d+\])", p.text)
        # 获取原文的字体和字号
        if flag:
            quotesDir[flag.group()] = p
            quoteFontName = p.runs[0].font.name
            quoteFontSize = p.runs[0].font.size


# 给参考文献列表排序并删除多余索引
def quote_sort():
    # 给参考文献排序
    quotesDirCopy = copy.deepcopy(quotesDir)
    # 遍历正文索引, 找到对应的参考文献, 进行标号替换
    for indexKey in indexDir:
        quoteIndex = "[" + str(indexDir[indexKey]) + "]"
        quotesDir[quoteIndex].text = quotesDirCopy[indexKey].text

    # 删除多余的索引, 并且按照参考文献标号调整顺序
    # quotesDirCopy = copy.deepcopy(quotesDir)
    num = 0
    for key in quotesDir:
        num += 1
        if num > len(quoteIndex):
            quotesDir[key].text = ""
        # 改变字体
        repair_font_size(quotesDir[key], quoteFontName, quoteFontSize)


# 保存文档
def docx_save(document, save_path):
    document.save(save_path)


# 修改字体
def repair_font_size(paragraph, font_name, font_size):
    # 遍历每个run, 更新为原文的字体和字号
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = font_size
