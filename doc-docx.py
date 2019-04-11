import docx
import os
from win32com import client
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

#将doc文件转换为docx
log=input("请输入要转换的文件的位置：")
cov=input("请输入转换后文件的保存位置：")

pathfile=os.listdir(log)
for filename in pathfile:
    filedir=log+'/'+filename

    print(filename)
    filelist = filename.strip('.').split('.')

    if filelist[4] == "doc":
        print(filedir)
        ip='.'.join(filelist[:4])
        print(ip)
        word=client.Dispatch('Word.Application')
        doc = word.Documents.Open(filedir)  # 目标路径下的文件
        doc.SaveAs(cov+ '/'+ip+ ".docx", 16)  # 转化后路径下的文件
        doc.Close()
    else:
        continue

