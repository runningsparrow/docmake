# -*- coding:utf-8 -*-

import sys

#解决无法导入不同目录模块的问题
sys.path.append("..")

from src.mysqlhelp import Basedb


class buildtb():

    def __init__(self):
        pass

    def build(self):
        basedb = Basedb();
        engine = basedb.crengine()
        # 建表
        basedb.createtb(engine)

    def drop(self):
        basedb = Basedb();
        engine = basedb.crengine()
        #删表
        basedb.dropdb(engine)

if __name__ == '__main__':
    bt = buildtb()
    # bt.drop()
    bt.build();
