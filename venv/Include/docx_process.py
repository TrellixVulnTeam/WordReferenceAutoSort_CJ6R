# -*- coding=utf-8 -*-
"""文档处理

"""
__author__ = "bjx"

from docx import Document
import re
import copy


# 正文索引字典
indexDir = {}
# 参考文献字典
quotesDir = {}

# 创建文档对象
def create_document(path):
    document = Document(path)
    return document


# 给正文中的索引重新排序
def index_sort(document):
    global indexDir

    num = 1
    for paragraph in document.paragraphs:
        indexs = re.findall(r"(\w)+(\[\d+\])", paragraph.text)
        if len(indexs) == 1:
            indexNum = indexs[0][1][1:-1]
            old = "[" + str(indexNum) + "]"
            new = "[" + str(num) + "]"
            paragraph.text = paragraph.text.replace(old, new)
            indexDir[indexNum] = paragraph
            num += 1
        elif len(indexs) > 1:
            sentenses = paragraph.text.split("，")  # 如果是英文，注意使用英文标点
            text = ""
            for sentense in sentenses:
                if re.search(r"(\w)+(\[\d+\])", sentense):
                    partern = re.compile("\[\d\]")
                    new = "[" + str(num) + "]"
                    sentense = partern.sub(new, sentense)
                    num += 1
                text = text + "，" + sentense
            paragraph.text = text

            for index in indexs:
                indexNum = index[1][1:-1]
                indexDir[indexNum] = paragraph


# 给文末的参考文献排序
def quote_sort(document):
    global indexDir
    global quotesDir
    for p in document.paragraphs:
        flag = re.match(r"^(\[\d+\])", p.text)
        if flag:
            # print(type(flag.group()))
            quotesDir[flag.group()[1:-1]] = p

    # 给引用排序
    quotesDirCopy = copy.deepcopy(quotesDir)
    num = 1
    for indexKey in indexDir:
        quotesDir[str(num)].text = quotesDirCopy[indexKey].text
        quotesDir[str(num)].text = re.sub("\[\d\]", "[" + str(num) + "]", quotesDir[str(num)].text)
        num += 1

    ### 删除多余的索引
    for key in quotesDir:
        if int(key) > len(quotesDir):
            quotesDir[key].text = ""


# 保存文档
def docx_save(document):
    document.save("结果.docx")
