# -*- coding=utf-8 -*-
"""用户界面"""
__author__ = "闭锦秀"

import os
import tkinter
from tkinter import filedialog
from tkinter import StringVar
import docx
import docx_process
import traceback

# 输入/输出文档路径
file_input_path = None
file_output_path = None


# click chooseFile button
def choose_input_file(file_input_entry):
    global file_input_path
    file_input_path = tkinter.filedialog.askopenfilename()
    input_variable.set(file_input_path)


# click choose output path button
def choose_output_path(file_output_entry):
    global file_output_path
    file_output_path = tkinter.filedialog.askdirectory()
    output_variable.set(file_output_path)


# 提交路径开始处理文档
def submit(file_input_path):
    try:
        document = docx_process.create_document(file_input_path)
        docx_process.index_sort(document)
        docx_process.quote_sort(document)
        print(file_output_path)
        docx_process.docx_save(document, file_output_path + "/结果.docx")
        os.startfile(file_output_path)
        result_variable.set("已完成, 已为你打开输出文件夹")

    except docx.opc.exceptions.PackageNotFoundError:
        result_variable.set("错误! 请检查输入文件以及格式(.docx)是否正确!")
    except TypeError:
        result_variable.set("错误! 请检查输出路径是否存在!!")
    except PermissionError:
        result_variable.set("错误! 输出目录权限不足, 或请检查输出目录是否有同名文件正在使用, 请关闭后重试!!")
    except BaseException:
        result_variable.set("未知错误!!")
    else:
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
    file_input_label = tkinter.Label(input_label_frame, text=" 原文档路径: ")
    file_input_label.pack(side=tkinter.LEFT)

    # 文件路径输入框
    global input_variable
    input_variable = StringVar()
    input_variable.set("请在此处填写需要修改的文档路径")
    file_input_entry = tkinter.Entry(input_label_frame, textvariable=input_variable, width=50)
    file_input_entry.pack(side=tkinter.LEFT, padx=5)

    # 选择输入文件按钮
    file_input_choose_button = tkinter.Button(input_label_frame, text="选择文件", width=20, \
                                              command=lambda: choose_input_file(file_input_entry))
    file_input_choose_button.pack(side=tkinter.LEFT, padx=5)

    ## 输出
    output_label_frame = tkinter.LabelFrame(window, text="处理结果")
    output_label_frame.pack(fill=tkinter.Y, side=tkinter.TOP, padx=5, pady=15, ipadx=5, ipady=15)

    # 选择输出文件标签
    file_output_label = tkinter.Label(output_label_frame, text="结果输出路径:")
    file_output_label.pack(side=tkinter.LEFT)

    # 文件路径输出框
    global output_variable
    output_variable = StringVar()
    output_variable.set("请在此处填写文件的输出目录")
    file_output_entry = tkinter.Entry(output_label_frame, textvariable=output_variable, width=50)
    file_output_entry.pack(side=tkinter.LEFT, padx=5)

    # 选择输出文件按钮
    file_output_choose_button = tkinter.Button(output_label_frame, text="选择输出文件路径", width=20, \
                                               command=lambda: choose_output_path(file_output_entry))
    file_output_choose_button.pack(side=tkinter.LEFT, padx=5)

    # submit Button
    file_input_submit_button = tkinter.Button(window, text="开始执行", font=30, width=20,height=2, \
                                              command=lambda: submit(file_input_path))
    file_input_submit_button.pack(side=tkinter.TOP, padx=5)

    # 结果标签
    global result_variable
    result_variable = StringVar()
    result_variable.set("请选择输入/输出路径, 执行成功后会在输出路径生成名为'结果.docx的文档'")
    result_label = tkinter.Label(window, textvariable=result_variable)
    result_label.pack(side=tkinter.TOP, padx=5, pady=15)

    more_func_notic_label_frame = tkinter.LabelFrame(window, text="更多功能")
    more_func_notic_label_frame.pack(side=tkinter.BOTTOM, pady=15)
    more_func_label1 = tkinter.Label(more_func_notic_label_frame, text="更多功能正在开发中, 欢迎提出宝贵意见!", width=80)
    more_func_label1.pack()
    more_func_label2 = tkinter.Label(more_func_notic_label_frame, text="QQ: 1739473807", width=80)
    more_func_label2.pack()

    # 开始循环窗口
    window.mainloop()


if __name__ == '__main__':
    gui_main()
