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


# 创建文档对象
def create_document(path):
    document = Document(path)
    return document


# 给正文中的索引重新排序
def index_sort(document):
    global indexDir
    # 设定一个num, 记录是第几个参考文献引用
    num = 1
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
                indexNum = indexs[0][1][1:-1]
                # 修改前的标号
                old = "[" + str(indexNum) + "]"
                # 修改后的标号
                new = "[" + str(num) + "]"
                # 替换
                paragraph.text = paragraph.text.replace(old, new)
                # 将更新好的段落添加到字典
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
            # 修改字体
            repair_font_size(paragraph, font_name, font_size)


## 给文末的参考文献排序
def quote_sort(document):
    global indexDir
    global quotesDir

    # 字体
    font_size = None
    font_name = None

    # 遍历段落, 查找参考文献列表
    for p in document.paragraphs:
        flag = re.match(r"^(\[\d+\])", p.text)
        # 获取原文的字体和字号
        if flag:
            quotesDir[flag.group()[1:-1]] = p
            font_name = p.runs[0].font.name
            font_size = p.runs[0].font.size

    # 给参考文献排序
    quotesDirCopy = copy.deepcopy(quotesDir)
    # 设定一个num, 记录参考文献原本的序号
    num = 1
    # 遍历正文索引, 找到对应的参考文献, 进行标号替换
    for indexKey in indexDir:
        quotesDir[str(num)].text = quotesDirCopy[indexKey].text
        quotesDir[str(num)].text = re.sub("\[\d\]", "[" + str(num) + "]", quotesDir[str(num)].text)
        num += 1

    # 删除多余的索引, 并且按照参考文献标号调整顺序
    quotesDirCopy = copy.deepcopy(quotesDir)
    num = 1
    for key in quotesDir:
        if int(key) > len(quotesDir):
            quotesDir[key].text = ""
        else:
           quotesDir[key].text = quotesDirCopy[str(num)].text
           num += 1
        # 改变字体
        repair_font_size(quotesDir[key], font_name, font_size)



# 保存文档
def docx_save(document, save_path):
    document.save(save_path)


# 修改字体
def repair_font_size(paragraph, font_name, font_size):
    # 遍历每个run, 更新为原文的字体和字号
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = font_size
