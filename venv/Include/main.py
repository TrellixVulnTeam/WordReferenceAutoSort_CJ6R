# 注意: 需要安装python-docx库
from docx import Document
import re
import copy


file_path = input("请输入文档的绝对路径地址: ")
# file_path = r'D:\Development\python\docx-quite\venv\Include\test.docx'
# file_path = r'D:\Development\python\docx-quite\venv\Include\Cloud Club Platform.docx'

document = Document(file_path)


### 原文

indexDir = {}
"匹配文中的索引"
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




### 引用
quotesDir = {}
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
    quotesDir[str(num)].text = re.sub("\[\d\]","["+str(num)+"]",quotesDir[str(num)].text)
    num += 1

### 删除多余的索引
for key in quotesDir:
    if int(key) > len(quotesDir):
        quotesDir[key].text = ""


document.save("结果.docx")
