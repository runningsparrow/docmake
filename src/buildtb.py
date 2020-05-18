# -*- coding:utf-8 -*-

from mysqlhelp import Basedb

class buildtb():

    def __init__(self):
        pass

    def build(self):
        basedb = Basedb();
        engine = basedb.crengine()
        # 建表
        basedb.createtb(engine)

if __name__ == '__main__':
    bt = buildtb()
    bt.build();
