# -*- coding: utf-8 -*-


from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Float
from sqlalchemy.orm import sessionmaker


Base = declarative_base()  # 生成orm基类


class Basedb():
    
    def __init__(self):
        #
        MSSQL_CONNECT_STR = 'mssql+pymssql://sa:Jsfc-11111@139.196.98.186/JSFCOD'


        # 初始化数据库连接:
        self.db_conn_str = MSSQL_CONNECT_STR


    def crengine(self):
        engine = create_engine(self.db_conn_str, pool_recycle=3600,echo=True)
        return engine


    def createtb(self,engine):
        Base.metadata.create_all(engine)  # 创建表结构 （这里是父类调子类）


    def dropdb(self,engine):
        Base.metadata.drop_all(engine)

    def createsession(self):
        # dbsession = sessionmaker(bind=engine)
        dbsession = sessionmaker(bind=engine, expire_on_commit=False)
        session = dbsession()
        return session

    def adddata(self,session,objects):
        for object in objects:
            session.add(object)
            session.commit()

    def datatomodel(self,data,modelname):
        pass

    