# 注意: 需要安装python-docx库
from docx import Document
import re
import copy


# file_path = input("请输入文档的绝对路径地址: ")
file_path = r'D:\Development\python\docx-quotes-sort\venv\Include\test.docx'


document = Document(file_path)
for p in document.paragraphs:
    flag = re.match(r"^(\[\d+\])", p.text)
    if flag:
        # print("size: ",p.runs[0].font.size)
        for run in p.runs:
            print(run.font.name)
