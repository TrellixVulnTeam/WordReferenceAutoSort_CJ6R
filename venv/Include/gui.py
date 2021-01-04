# -*- coding=utf-8 -*-
"""用户界面"""
__author__ = "bjx"

import os
import tkinter
from tkinter import filedialog
from tkinter import StringVar
import docx_process


# 输入/输出文档路径
file_input_path = ""
file_output_path = ""




# click chooseFile button
def choose_input_file(file_input_entry):
    global file_input_path
    file_input_path = tkinter.filedialog.askopenfilename()
    file_input_entry.delete(0,"end")
    file_input_entry.insert(0, file_input_path)


# click choose output path button
def choose_output_path(file_output_entry):
    global file_output_path
    file_output_path = tkinter.filedialog.askdirectory()
    file_output_entry.delete(0,"end")
    file_output_entry.insert(0, file_output_path)


# 提交路径开始处理文档
def submit(file_input_path):
    document = docx_process.create_document(file_input_path)
    docx_process.index_sort(document)
    docx_process.quote_sort(document)
    docx_process.docx_save(document)
    os.startfile(file_output_path)
    result_variable.set("已完成, 已为你打开输出文件夹")


def gui_main():
    global file_input_path
    global file_output_path

    window = tkinter.Tk()
    window.title("Word文档参考文献自动排序工具")
    window.geometry("600x400")

    # 文档路径变量
    # input_path_variable = tkinter.StringVar()
    # input_path_variable.set("")
    # output_path_variable = tkinter.StringVar()
    # output_path_variable.set("")

    ## 输入
    input_label_frame = tkinter.LabelFrame(window, text="选择文档")
    input_label_frame.pack(fill=tkinter.Y, side=tkinter.TOP, padx=5, pady=15, ipadx=5, ipady=15)

    # 选择文件标签
    file_input_label = tkinter.Label(input_label_frame, text="文档路径:")
    file_input_label.pack(side=tkinter.LEFT)

    # 文件路径输入框
    file_input_entry = tkinter.Entry(input_label_frame, width=50)
    file_input_entry.pack(side=tkinter.LEFT, padx=5)

    # 选择输入文件按钮
    file_input_choose_button = tkinter.Button(input_label_frame, text="选择文件", \
                                              command=lambda: choose_input_file(file_input_entry))
    file_input_choose_button.pack(side=tkinter.LEFT, padx=5)

    ## 输出
    output_label_frame = tkinter.LabelFrame(window, text="处理结果")
    output_label_frame.pack(fill=tkinter.Y, side=tkinter.TOP, padx=5, pady=15, ipadx=5, ipady=15)

    # 选择输出文件标签
    file_output_label = tkinter.Label(output_label_frame, text="结果输出路径:")
    file_output_label.pack(side=tkinter.LEFT)

    # 文件路径输出框
    file_output_entry = tkinter.Entry(output_label_frame, width=50)
    file_output_entry.pack(side=tkinter.LEFT, padx=5)

    # 选择输出文件按钮
    file_output_choose_button = tkinter.Button(output_label_frame, text="选择输出文件路径", \
                                               command=lambda: choose_output_path(file_output_entry))
    file_output_choose_button.pack(side=tkinter.LEFT, padx=5)

    # submit Button
    file_input_submit_button = tkinter.Button(window, text="开始执行", width=10, \
                                              command=lambda: submit(file_input_path))
    file_input_submit_button.pack(side=tkinter.TOP, padx=5)

    # 结果标签
    global result_variable
    result_variable=StringVar()
    result_variable.set("准备")
    result_label = tkinter.Label(window, textvariable=result_variable)
    result_label.pack(side=tkinter.TOP, padx=5, pady=15)

    # 开始循环窗口
    window.mainloop()


if __name__ == '__main__':
    gui_main()
