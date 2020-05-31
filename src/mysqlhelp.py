# -*- coding: utf-8 -*-


from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Float
from sqlalchemy.orm import sessionmaker


Base = declarative_base()  # 生成orm基类


class Basedb():
    
    def __init__(self):
        #
        # MSSQL_CONNECT_STR = 'mssql+pymssql://sa:Jsfc-11111@139.196.98.186/JSFCOD'
        MYSQL_CONNECT_STR = 'mysql+pymysql://root:wzkj2015@39.99.168.212/test1'


        # 初始化数据库连接:
        self.db_conn_str = MYSQL_CONNECT_STR


    def crengine(self):
        engine = create_engine(self.db_conn_str, pool_recycle=3600,echo=False)
        return engine


    def createtb(self,engine):
        Base.metadata.create_all(engine)  # 创建表结构 （这里是父类调子类）


    def dropdb(self,engine):
        Base.metadata.drop_all(engine)

    def createsession(self,engine):
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


    

class doconfig(Base):
    __tablename__ = 'doconfig'  # 生成 doc 配置表 

    
    doc_id = Column(Integer, primary_key=True)  #文档 id
    doc_name = Column(String(200)) #文档 名称  
    doc_template = Column(String(200)) #文档 模板路径
    doc_outpath = Column(String(200))  #文档 输出路径
    doc_label_text = Column(String(200)) #文档 json 文件 (包含路径)
    doc_image_dir = Column(String(200)) #文档图片目录
    doc_excel = Column(String(200)) #文档 excel 文件
    doc_attach_dir = Column(String(200)) #附件目录
    doc_rmrk = Column(String(500)) #文档备注








