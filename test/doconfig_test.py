# -*- coding: utf-8 -*-
import sys

#解决无法导入不同目录模块的问题
sys.path.append("..")
sys.path.append(".")

from src.buildtb import buildtb
from src.docmaker import docmaker


if __name__ == '__main__':
    print("test")

    dm = docmaker()

    # ret = dm.insertdocdata("test1","test1.docx","test11.docx","text1.json","test1","excel1.xlsx","test1","doc_rmrk1")
    ret = dm.insertdocdata("test2","test2.docx","test21.docx","text2.json","test2","excel2.xlsx","test2","doc_rmrk2")

    



    
