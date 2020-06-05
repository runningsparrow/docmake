# -*- coding: utf-8 -*-

#attention:
#run command below first:
#pip install -r requirements.txt

from docx import Document

import json
import sys
import os
import xlrd
import re
import sys
#替换文件需要用到的包
import zipfile
import shutil
import random
import string


#解决无法导入不同目录模块的问题
sys.path.append("..")
sys.path.append(".")

from mysqlhelp import Basedb,doconfig
from returndata import returndata1



class docmaker():

    def __init__(self):
        self.name = "docmaker"
        self.tablename = "doconfig"
        self.basedb = Basedb();
        self.engine = self.basedb.crengine()
        self.session1 = self.basedb.createsession(self.engine)

        self.flagtext = '@@text\d'
        self.flagimage = '@@image\d'
        self.flagimage1 = 'image\d'
        self.flagsheet = '@@Sheet\d'    
        self.flagsheet1 = '@@excel\d'


    def makedoc(self,doc_name):
        #read table
        docdata = self.querydocdata(doc_name)

        # print(docdata)

        print(docdata["returndt"][0].doc_template)
        print(docdata["returndt"][0].doc_outpath)
        print(docdata["returndt"][0].doc_label_text)
        print(docdata["returndt"][0].doc_image_dir)
        print(docdata["returndt"][0].doc_excel)


        #open tempalte
        docfile =  os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/template/" + docdata["returndt"][0].doc_template
        print(docfile)
        document = Document(docfile)

        #read json text

        jsonfile = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/text/" + docdata["returndt"][0].doc_label_text
        load_f = open(jsonfile, 'r',encoding='UTF-8') 
        jsondata = json.load(load_f)

        print(jsondata)
        # print(jsondata["test1"].keys())

        #get key list
        # keylist = jsondata["test1"].keys()

        # for key in keylist:
        #     print(key)
        #     print(jsondata["test1"][key])



        #read image dir 
        imagedir = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/image/" + docdata["returndt"][0].doc_image_dir
        imagedirs = []
        imagefiles = []
        imagefilesnames = []
        for item in os.scandir(imagedir):
            if item.is_dir():
                imagedirs.append(item.path)

            elif item.is_file():
                imagefiles.append(item.path)
                imagefilesnames.append(item.name)

        print('\n'.join(imagefiles))
        print('\n'.join(imagefilesnames))



        #read excel
        excelfile = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/excel/" + docdata["returndt"][0].doc_excel
        # excelfile = 'I://workspace1/python/docmake/resouce/excel/test1.xlsx'

        wb = xlrd.open_workbook(excelfile)

        # debug
        # for sheet in wb.sheets():
        #     print(sheet.name)

        #process doc
        for paragraph in document.paragraphs:
            print("===========")
            s = paragraph.text

            searchtext = re.search(self.flagtext,s)

            searchimage = re.search(self.flagimage,s)

            searchsheet = re.search(self.flagsheet,s)
            

            #文本
            if searchtext:

                print("searchtext")

                print(searchtext)

                p = re.compile(self.flagtext)

                # print(p)


                #get key list
                keylist = jsondata["test1"].keys()

                for key in keylist:
                    print("遍历")
                    print(key)
                    if key == paragraph.text:
                        ss = p.sub(jsondata["test1"][key],s)
                        paragraph.text = ss
                        print(paragraph.text)
                    else:
                        print("not match")
                        pass
                
            
            else:

                print("not found text")

            #图片
            if searchimage:

                print("searchimage")

                p = re.compile(self.flagimage)

                

                for imagename in imagefilesnames:
                    
                   

                    imagenamenosuffix = imagename.split(".")

                    print(imagenamenosuffix[0])

                    searchimage1 = re.search(self.flagimage1,imagenamenosuffix[0])

                    if searchimage1:
                        print(imagenamenosuffix[0])
                        print(s[2:8])
                        if s[2:8] == imagenamenosuffix[0]:

                            

                            ss = p.sub('',s)
                            paragraph.text = ss
                            run = paragraph.add_run()
                            run.add_picture(imagedir+'/'+imagename)
         
            else:
    
                print("not found image")

            #电子表格
            if searchsheet:

                print("searchsheet")

                sheetname = s[2:8]

                # print(sheetname)

                for sheet in wb.sheets():
                    # print(sheet.name)

                    if sheet.name == sheetname:

                        print("debug")
                        print(sheet.name)
                        print(sheet.nrows)
                        print(sheet.ncols)

                        
                        table = document.add_table(sheet.nrows,sheet.ncols)

                        for i in range(0,sheet.nrows):
                            for j in range(0,sheet.ncols):
                                print(i)
                                print(j)
                                print(sheet.cell_value(i,j))
                                table.cell(i,j).text = str(sheet.cell_value(i,j))

                        
                        #
                        self.move_table_after(table,paragraph)
                        #
                        paragraph.text = ''
            else:
    
                print("not found sheet")


        #save doc to outpath

        savepath = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/output/" + docdata["returndt"][0].doc_outpath

        document.save(savepath)

        attach_dir = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/attachment/" + docdata["returndt"][0].doc_attach_dir
        self.subattach(savepath,attach_dir)






    #移动表格
    def move_table_after(self,table, paragraph):
        tbl, p = table._tbl, paragraph._p
        p.addnext(tbl)


    #替换word里的附件

    def subattach(self,docfile,attach_dir):

        #建立临时父目录
        if not os.path.exists(os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/tempdir/"):
            os.mkdir(os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/tempdir/")


        #以压缩格式打开word文件
        zipdoc = zipfile.ZipFile(docfile) 
        
        tmpdir = ""

        pathsuffix = os.path.abspath(os.path.dirname(os.path.abspath("__file__"))) + "/resouce/tempdir/"

        #生行8位临时文件夹名

        while True:
            tmpdir= ''.join(random.sample(string.ascii_letters + string.digits, 8))   #生行8位临时文件夹名
            if not os.path.exists(tmpdir):
                break
        tmpdir1 = pathsuffix + tmpdir

        print(tmpdir1)

        os.mkdir(tmpdir1)                  #创建临时目录
        os.chdir(tmpdir1)                   #转到临时目录
        zipdoc.extractall()                     #解压word文件到临时文件夹
        zipdoc.close()                           #关闭word文档，否则后面重新压缩会报错


        #read attach dir 
        attachdir = attach_dir
        attachdirs = []
        attachfiles = []
        attachfilesnames = []
        for item in os.scandir(attachdir):
            if item.is_dir():
                attachdirs.append(item.path)

            elif item.is_file():
                attachfiles.append(item.path)
                attachfilesnames.append(item.name)

        print('\n'.join(attachfiles))
        print('\n'.join(attachfilesnames))

        #获取docx里的嵌入对象
        embeddingdir = tmpdir1 + "/word/embeddings/"
        embeddingdirs = []
        embeddingfiles = []
        embeddingfilesnames = []
        for item in os.scandir(embeddingdir):
            if item.is_dir():
                embeddingdirs.append(item.path)

            elif item.is_file():
                embeddingfiles.append(item.path)
                embeddingfilesnames.append(item.name)

        print('\n'.join(embeddingfiles))
        print('\n'.join(embeddingfilesnames))

        #把正确文件拷贝覆盖模版文件的空附件
        i = 0
        
        for embedingfilename in embeddingfilesnames:
            j = 0
            for attachfilesname in attachfilesnames:
                if embedingfilename == attachfilesname:
                    
                    shutil.copy(attachfiles[j],embeddingfiles[i])


                j = j + 1
            
            i = i + 1

        #
        azip = zipfile.ZipFile(docfile, 'w')    #以压缩格式新建word文档
        for i in os.walk('.'):                             #使用os.walk遍历整个目录及子目录，保证原有的目录结构不变
            for j in i[2]:
                azip.write(os.path.join(i[0],j), compress_type=zipfile.ZIP_DEFLATED)     #将文件逐个打包到word文档中，压缩格式指定为ZIP_DEFLATED
        azip.close()                                       #关闭文件

        os.chdir('..')
        shutil.rmtree(tmpdir1,ignore_errors=True)    #删除临时文件夹



        

    def insertdocdata(self,doc_name,doc_template,doc_outpath,doc_label_text,doc_image_dir,doc_excel,doc_attach_dir,doc_rmrk=""):
        #check exists
        queryret = self.querydocdata(doc_name)
        if len(queryret["returndt"]) != 0:
            print("文档配置已存在")
            rd = returndata1
            rd["returncd"] = 1
            rd["returndt"] = "文档配置已存在"
            
        else:
            print("可创建文档配置")

            #check datacount
            querycount = self.querydocdatacount()
            datacount = len(querycount["returndt"])

            dc = doconfig()
            dc.doc_id = datacount + 1
            dc.doc_name = doc_name
            dc.doc_template = doc_template
            dc.doc_outpath = doc_outpath
            dc.doc_label_text = doc_label_text
            dc.doc_image_dir = doc_image_dir
            dc.doc_excel = doc_excel
            dc.doc_attach_dir = doc_attach_dir
            dc.doc_rmrk = doc_rmrk

            self.session1.add(dc)
            self.session1.commit()
            self.session1.close()

            # print(dc)
            rd = returndata1
            rd["returncd"] = 0
            rd["returndt"] = dc

        # print(rd)

        return rd



    def repairdocdata(self,doc_name,doc_template,doc_outpath,doc_label_text,doc_image_dir,doc_excel,doc_attach_dir,doc_rmrk=""):
        #check exists
        queryret = self.querydocdata(doc_name)

        if len(queryret["returndt"]) == 0:
            print("文档配置已存在")
            rd = returndata1
            rd["returncd"] = 1
            rd["returndt"] = "文档配置未找到，更新未成功"
        else:

            

            #更新数据

            dc = queryret["returndt"][0]

            dc.doc_template = doc_template
            dc.doc_outpath = doc_outpath
            dc.doc_label_text = doc_label_text
            dc.doc_image_dir = doc_image_dir
            dc.doc_excel = doc_excel
            dc.doc_attach_dir = doc_attach_dir
            dc.doc_rmrk = doc_rmrk

            self.session1.commit()
            self.session1.close()

            rd = returndata1
            rd["returncd"] = 0
            rd["returndt"] = dc

        return rd


    def deletedocdata(self,doc_name):

        dc = self.session1.query(doconfig).filter(doconfig.doc_name == doc_name).all()

        if(len(dc)!=0):

            rd = returndata1
            rd["returncd"] = 0
            rd["returndt"] = dc

            self.session1.query(doconfig).filter(doconfig.doc_name == doc_name).delete()
            self.session1.commit()
            self.session1.close()

        else:
           
            rd = returndata1
            rd["returncd"] = 1
            rd["returndt"] = "未找到数据，未删除"

        return rd


    def querydocdata(self,doc_name):
        #read table data   
        dc = self.session1.query(doconfig).filter(doconfig.doc_name == doc_name).all()

        # print("sssss")
        # print(dc)

        rd = returndata1


        rd["returncd"] = 0
        rd["returndt"] = dc


        # print(dc.doc_template)    
        return rd


    def querydocdatacount(self):
        dc = self.session1.query(doconfig)
        # print(len(dc.all()))
        # return dc.all()
        rd = returndata1
        rd["returncd"] = 0
        rd["returndt"] = dc.all()

        print(dc.all())

        return rd
        
           

if __name__ == '__main__':

    # print("start python script")

    para = 'parameter 1111'
    
    # print(sys.argv)
    if len(sys.argv) > 1:

        # print("进程 " +sys.argv[1] +" 执行。") 
        para = sys.argv[1]
    else:

        # print("无参数")
        pass



    dm = docmaker()
    # dm.querydocdata("test1")
    # dm.querydocdatacount();
    # ret = dm.insertdocdata("test4","template","fasdfs","ddddd","doc_image_dir","doc_excel","doc_rmrk")

    # ret = dm.repairdocdata("test4",para,"fasdfs","ddddd","doc_image_dir","doc_excel","doc_rmrk")

    # if(ret["returncd"]) == 0:

    #     # print(ret["returndt"].doc_id)

    #     # print(ret)
    #     # print(ret["returndt"].doc_id)
    #     # print(ret["returndt"].doc_name)
    #     # print(ret["returndt"].doc_template)
    #     # print(ret["returndt"].doc_outpath)
    #     # print(ret["returndt"].doc_label_text)
    #     # print(ret["returndt"].doc_image_dir)
    #     # print(ret["returndt"].doc_excel)
    #     # print(ret["returndt"].doc_rmrk)

    #     print(ret["returndt"])
    
    # else:

    #     print("未插入数据")


    dm.makedoc("test1")

    # dm.querydocdatacount();

    



    
