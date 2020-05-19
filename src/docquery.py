# -*- coding: utf-8 -*-
import sys

from docmaker import docmaker


def doquerylist():

    if len(sys.argv) > 1:
        
        print("进程 " +sys.argv[1] +" 执行。") 
        para = sys.argv[1]
    else:
        print("无参数")

    dm = docmaker()
    return dm.querydocdatacount()






if __name__ == "__main__":
    doquerylist()
        
        



