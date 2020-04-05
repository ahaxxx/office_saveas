# -*- coding: UTF-8 -*-
# -*- python:3.7    -*-
# -*- author:LiuBX  -*-
import win32com.client
import os

def word_to_pdf(word_path,pdf_path):
    # 相对路径转绝对路径
    word_path = os.path.abspath(word_path)
    pdf_path = os.path.abspath(pdf_path)
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(word_path)           
    print("读取文件成功")
    doc.SaveAs(pdf_path,17)                        
    print("转换成功")
    doc.Close()
    word.Quit()
