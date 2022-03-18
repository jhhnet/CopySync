#!/usr/bin/env python
# -*- encoding: utf-8 -*-

"""
@Project : first 
@File    : ReadPages.py
@IDE     : PyCharm-2021.1 python-3.7
@Description : 输入含office文件的路径,得到各个文件的页码及名称

@Date    : 2022/3/15 23:39
@Author  : jhhnet
@Contact : jiang_hnu@163.com
@Version : 1.0          
"""

import os
import stat
import time
from typing import Tuple
# import PyPDF2
import pdfplumber
import pythoncom
from pdfminer.pdfparser import PDFSyntaxError
from pptx import Presentation
from win32com.client import DispatchEx


class GetPages:
    def __init__(self, path: str, ends=()):
        self.__path = path.replace('\\', '\\\\')
        self.__ends = ends

    def current(self):
        return self.__path

    def get_path(self) -> str:
        return self.__path

    def set_path(self, spath: str):
        # try:{} except ReferenceError:
        if os.path.exists(spath):
            self.__path = spath
        else:
            raise ReferenceError

    def get_ends(self) -> Tuple:
        return self.__ends

    def set_ends(self, ends: Tuple):
        self.__ends = ends

    # 获得以ends类型结尾的所有文件，返回一个list
    def get_all_file_name(self):
        files_path_list = []
        # root, _, _ = os.walk(path) = path
        files = os.listdir(self.__path)
        for name in files:
            if name.endswith(self.__ends):
                files_path = os.path.join(self.__path, name)
                files_path_list.append(files_path)
        print("总共有%d个文件" % len(files_path_list))
        return files_path_list

    # 获取PDF文档页数
    # reader = PyPDF2.PdfFileReader(pdf_path)
    # page_num = reader.getNumPages()
    @staticmethod
    def get_pdf_page(file_path: str):
        if file_path.endswith(("PDF", "pdf")):
            try:
                times = time.strftime("%Y/%m/%d", time.localtime(os.path.getmtime(file_path)))
                pdf = pdfplumber.open(file_path)
                page = len(pdf.pages)
            except PDFSyntaxError:
                page = 'error0'
        else:
            page = 'NOpdf'
        return times, page

    # 获取PPT页数
    @staticmethod
    def get_ppt_page(file_path: str):
        if file_path.endswith(("pptx", "ppt")):
            try:
                times = time.strftime("%Y/%m/%d", time.localtime(os.path.getmtime(file_path)))
                # 修改文件权限为 owner group other只可读
                # os.chmod(file_path, stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)
                os.chmod(file_path, stat.S_IREAD)
                ppt = Presentation(file_path)
                page = len(ppt.slides)
            except KeyError:
                page = 'error0'
        else:
            page = 'NOppt'
        return times, page

    # 获取word文档页数
    @staticmethod
    def get_word_page(file_path: str):
        if file_path.endswith(("docx", "doc")):
            times = time.strftime("%Y/%m/%d", time.localtime(os.stat(file_path).st_mtime))
            os.chmod(file_path, stat.S_IREAD)
            pythoncom.CoInitialize()  # 是可以正常运行的 启用线程
            word = DispatchEx("Word.Application")
            word.Visible = False  # 后台
            word.DisplayAlerts = 0
            doc = word.Documents.Open(file_path)
            word.ActiveDocument.Repaginate()  # 重标文档页码 Close会更改修改日期
            page = word.ActiveDocument.ComputeStatistics(2)
            # doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()  # 是可以正常运行的
            os.chmod(file_path, stat.S_IWRITE)
        else:
            page = 'NOdoc'
        return times, page

        # 获取文件夹内的docx pdf ppt文件的页数及路径->test.txt
    def get_dir_pages(self):
        files = self.get_all_file_name()
        for file in files:
            if file.endswith(("PDF", "pdf")):
                times, page = self.get_pdf_page(file)
                with open("test.txt", "a", encoding="utf-8") as f:
                    f.write(file + " %s %d 页\n" % (times, page))
            elif file.endswith(("pptx", "ppt")):
                times, page = self.get_ppt_page(file)
                with open("test.txt", "a", encoding="utf-8") as f:
                    f.write(file + " %s %d 页\n" % (times, page))

            elif file.endswith(("docx", "doc")):
                times, page = self.get_word_page(file)
                with open("test.txt", "a", encoding="utf-8") as f:
                    f.write(file + " %s %d 页\n" % (times, page))
            else:
                with open("test.txt", "a", encoding="utf-8") as f:
                    f.write(file + " error" + " error 页\n")


def main():
    path = r"D:\Desktop\新建文件夹"
    ends = ("PDF", "pdf", "pptx", "ppt", "docx", "doc")
    gg = GetPages(path, ends=ends)
    print(gg.current())
    gg.get_dir_pages()


if __name__ == '__main__':
    main()
