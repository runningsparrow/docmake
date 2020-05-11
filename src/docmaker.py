# -*- coding: utf-8 -*-

from docx import Document
from mysqlhelp import Basedb,doconfig


class docmaker():

    def __init__(self):
        self.name = "docmaker"
        self.tablename = "doconfig"


    def makedoc(self,docid):
        
        #read table data
        basedb = Basedb();
        engine = basedb.crengine()
        session1 = basedb.createsession(engine)

    
