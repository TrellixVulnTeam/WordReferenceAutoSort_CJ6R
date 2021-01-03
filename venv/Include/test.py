from docx import Document
import re


quoteInText = {}
def searchIndex(document):
    "匹配文中的索引"
    num = 1
    for p in document.paragraphs:
        flag = re.findall(r"(\w|\d)+\[\d+\]", p.text)
        for f in flag:
            quoteInText[num]=p.text
            num+=1

def quotes(document):
    "匹配引用的书目录"
    quotesDir = {}
    for p in document.paragraphs:
        flag = re.match(r"^(\[\d+\])", p.text)
        if flag:
            quotesDir[flag.group()[1:-1]] = p

    for key, value in quotesDir.items():
        print(key, value.text)

# 获取style名称
# for s in document.styles:
#     print(s.name)


# for key, value in quotesDir.items():
#     print(key, value.text)
