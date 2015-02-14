# -*- coding: UTF-8 -*- 
import os
import wx
import shutil
import win32com
import datetime
import win32api
import re
from win32com.client import DispatchEx,constants


CESHI = 1 #(测试为1,正常为0)

def DD():
    DIQU = u''
    DANWEI = u''    
    if os.path.exists(r"./re.py") :
        with open(r"./re.py") as o:
            o.readline()
            o.readline()
            DIQU = o.readline().strip()
            DANWEI = o.readline().strip()
    return (unicode(DIQU),unicode(DANWEI))
            
# 柜打印和盒打印模板相同(档号、索引号、案卷题名、总页数、保管期限、文件类别、备注)
def doGuiAnJuanMuLuPrint(self,printList = [('test','test','test','test','test','test','test','test','test','0','test','test','test')],guiOrHe=u''):
    DIQU,DANWEI = DD()
    msword = DispatchEx('Word.Application')
    msword.Visible = CESHI
    #msword.DisplayAlerts = CESHI
    doc = msword.Documents.Add()  
    pre_section = doc.Sections(1)
    new_seciton = doc.Range(pre_section.Range.End-1, pre_section.Range.End-1).Sections.Add()
    new_range = new_seciton.Range
    if guiOrHe == u'gui':
        new_range.Text = unicode(DANWEI[3:])+u"天保工程档案柜层案卷目录"
    else:
        new_range.Text = unicode(DANWEI[3:])+u"天保工程档案盒案卷目录"
    new_range.Font.Size = 15
    new_range.ParagraphFormat.Alignment = 1
    new_table = new_range.Tables.Add(doc.Range(new_range.End-1,new_range.End-1), len(printList)+1, 8)
    #new_table = new_range.Tables.Add(doc.Range(0,0), len(printList)+1, 8)
    new_table.Cell(1,1).Range.Text = u"序号"
    new_table.Cell(1,2).Range.Text = u"档号"
    new_table.Cell(1,3).Range.Text = u"索引号"
    new_table.Cell(1,4).Range.Text = u"案卷题名"
    new_table.Cell(1,5).Range.Text = u"总页数"
    new_table.Cell(1,6).Range.Text = u"保管期限"
    new_table.Cell(1,7).Range.Text = u"文件类别"
    new_table.Cell(1,8).Range.Text = u"备注"
    new_table.Rows(1).Borders(constants.wdBorderBottom).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderTop).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderLeft).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderRight).LineStyle = 1
    hang = len(printList)
    i = 2
    j = 1
    for row in printList:
        if i < hang+2 :
            new_table.Cell(i,1).Range.Text = j
            new_table.Cell(i,2).Range.Text = row[0]+u'-'+row[1]+u'-'+row[2]+u'-'+row[3]
            new_table.Cell(i,3).Range.Text = row[4]+row[5]+row[6]+row[7]
            new_table.Cell(i,4).Range.Text = row[8]
            new_table.Cell(i,5).Range.Text = int(row[9])
            new_table.Cell(i,6).Range.Text = row[10]
            new_table.Cell(i,7).Range.Text = row[11]
            new_table.Cell(i,8).Range.Text = row[12]
            new_table.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1            
            i = i + 1
            j = j + 1
    for i in range(1,8):
        new_table.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1         
    #删除第一页空白
    pre_section = doc.Sections(1)
    doc.Range(pre_section.Range.Start, pre_section.Range.End-1).Delete(1)    
    
    if guiOrHe == u'gui' :
        path = os.getcwd()+"\\打印\\柜案卷目录"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    else :
        path = os.getcwd()+"\\打印\\盒案卷目录"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    saveAndOpen(msword, doc , path)

def doJuanMuLuPrint(self,printList = [('test','test','test','test','test','test','test','test','test','test','test')]):
    DIQU,DANWEI = DD()
    msword = DispatchEx(r'Word.Application')
    msword.Visible = CESHI
    #msword.DisplayAlerts = CESHI 
    doc = msword.Documents.Add()   
    pre_section = doc.Sections(1)
    new_seciton = doc.Range(pre_section.Range.End-1, pre_section.Range.End-1).Sections.Add()
    new_range = new_seciton.Range
    new_range.Text = unicode(DANWEI[3:])+u"天保工程档案(卷内)目录"
    new_range.Font.Size = 15
    new_range.ParagraphFormat.Alignment = 1
    new_table = new_range.Tables.Add(doc.Range(new_range.End-1,new_range.End-1), len(printList)+1, 11)
    new_table.Cell(1,1).Range.Text = u"序号"
    new_table.Cell(1,2).Range.Text = u"区号+柜号"
    new_table.Cell(1,3).Range.Text = u"盒号"
    new_table.Cell(1,4).Range.Text = u"卷号"
    new_table.Cell(1,5).Range.Text = u"文件编号"
    new_table.Cell(1,6).Range.Text = u"发文单位"
    new_table.Cell(1,7).Range.Text = u"文件标题"
    new_table.Cell(1,8).Range.Text = u"发文日期"
    new_table.Cell(1,9).Range.Text = u"页数"
    new_table.Cell(1,10).Range.Text = u"互见号"
    new_table.Cell(1,11).Range.Text = u"备注"
    new_table.Rows(1).Borders(constants.wdBorderBottom).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderTop).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderLeft).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderRight).LineStyle = 1    
    hang = len(printList)
    i = 2
    j = 1
    for row in printList:
        if i < hang+2 :
            new_table.Cell(i,1).Range.Text = j
            new_table.Cell(i,2).Range.Text = row[0]+row[1]
            new_table.Cell(i,3).Range.Text = row[2]
            new_table.Cell(i,4).Range.Text = row[3]
            new_table.Cell(i,5).Range.Text = row[4]
            new_table.Cell(i,6).Range.Text = row[5]
            new_table.Cell(i,7).Range.Text = row[6]
            new_table.Cell(i,8).Range.Text = row[7]
            new_table.Cell(i,9).Range.Text = row[8]
            new_table.Cell(i,10).Range.Text = row[9]
            new_table.Cell(i,11).Range.Text = row[10]
            new_table.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1             
            i = i + 1
            j = j + 1
    for i in range(1,11):
        new_table.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
        new_table.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1         
    #删除第一页空白
    pre_section = doc.Sections(1)
    doc.Range(pre_section.Range.Start, pre_section.Range.End-1).Delete(1)      
    path = os.getcwd()+"\\打印\\卷目录"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    saveAndOpen(msword, doc , path)

def doAnJuanBiaoTiPrint(self,printList=[(u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test')]):
    DIQU,DANWEI = DD()
    DW = u''
    msword = win32com.client.gencache.EnsureDispatch("Word.Application")
    #msword = DispatchEx(r'Word.Application')
    msword.Visible = CESHI
    #msword.DisplayAlerts = CESHI
    doc = msword.Documents.Add()
    for section_index in range(1,len(printList)+1) :
        pre_section = doc.Sections(section_index)
        new_seciton = doc.Range(pre_section.Range.End-1, pre_section.Range.End-1).Sections.Add()
        new_range = new_seciton.Range
        new_table = new_range.Tables.Add(doc.Range(new_range.End-1,new_range.End-1), 4, 1)
        for i in range(1,5):
            new_table.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1            
        new_table.Cell(2,1).Range.InsertAfter(new_table.Cell(2,1).Range.Tables.Add(new_table.Cell(2,1).Range,3,2))
        sub_table_top = new_table.Cell(2,1).Range.Tables.Add(new_table.Cell(2,1).Range,3,2)
        for i in range(1,4):
            sub_table_top.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1 
        for i in range(1,3):
            sub_table_top.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1            
        sub_table_top.Cell(1,1).Range.Text = u"全宗号："
        sub_table_top.Cell(2,1).Range.Text = u"档  号："
        sub_table_top.Cell(3,1).Range.Text = u"索引号："
        DW = printList[section_index-1][0][-1] + unicode(DIQU)
        sub_table_top.Cell(1,2).Range.Text = printList[section_index-1][0]+unicode(DIQU)+u'-'+printList[section_index-1][1]+u'-'+printList[section_index-1][2]#"全宗号值："
        sub_table_top.Cell(2,2).Range.Text = printList[section_index-1][3]+u'-'+printList[section_index-1][4]+u'-'+printList[section_index-1][5]+u'-'+printList[section_index-1][6]#"档号值："
        sub_table_top.Cell(3,2).Range.Text = printList[section_index-1][7]+printList[section_index-1][8]+printList[section_index-1][9]+printList[section_index-1][10]#"索引号值："
        sub_table_top.Columns(1).SetWidth(2.5*28.35,0)
        sub_table_top.Columns(2).SetWidth(4*28.35,0)
        sub_table_top.Rows.Alignment = 0
        sub_table_bottom = new_table.Cell(4,1).Range.Tables.Add(new_table.Cell(4,1).Range,1,2)
        for i in range(1,4):
            sub_table_bottom.Rows.Add()
        for i in range(1,5):
            sub_table_bottom.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1 
        for i in range(1,3):
            sub_table_bottom.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1 
        sub_table_bottom.Cell(1,1).Range.Text = u"立卷单位:"
        sub_table_bottom.Cell(2,1).Range.Text = u"起止日期:"
        sub_table_bottom.Cell(3,1).Range.Text = u"保管期限:"
        sub_table_bottom.Cell(4,1).Range.Text = u"密    级:"
        sub_table_bottom.Cell(1,2).Range.Text = printList[section_index-1][11]#"立卷单位值"
        #设置临时日期变量
        temp_date = unicode(printList[section_index-1][12])
        if temp_date == u'None' :
            temp_date = u''
        else:
            temp_date = temp_date[:4]+u'年'+temp_date[4:6]+u'月'+temp_date[6:8]+u'日' + u'--' +\
                        temp_date[9:13]+u'年'+temp_date[13:15]+u'月'+temp_date[15:]+u'日'
        sub_table_bottom.Cell(2,2).Range.Text = temp_date #"起止日期值"
        sub_table_bottom.Cell(3,2).Range.Text = printList[section_index-1][13]#"保管期限值"
        sub_table_bottom.Cell(4,2).Range.Text = printList[section_index-1][14]#"密级值"
        sub_table_bottom.Columns(1).SetWidth(3*28.35,0)
        sub_table_bottom.Columns(2).SetWidth(8*28.35,0)
        sub_table_bottom.Rows.Alignment = 1
        new_table.Cell(2,1).Range.Font.Size = 12
        new_table.Rows(1).Borders(constants.wdBorderBottom).LineStyle = 0
        new_table.Cell(3,1).Range.Font.Size = 25
        new_table.Cell(3,1).Range.Text = "\n\n"+printList[section_index-1][15]#"案卷题目值"
        new_table.Cell(3,1).Range.ParagraphFormat.Alignment = 1
        new_table.Cell(4,1).Range.Font.Size = 12
        new_table.Rows(1).Height = 3
        new_table.Rows(2).Height = 120
        new_table.Rows(2).Borders(constants.wdBorderBottom).LineStyle =0
        new_table.Rows(3).Height = 330
        new_table.Rows(3).Borders(constants.wdBorderBottom).LineStyle =0
        new_table.Rows(4).Height = 200
        #字号调整sub_table_top.Cell(1,2).Range.Font.Size = 5
    #删除第一页空白
    pre_section = doc.Sections(1)
    doc.Range(pre_section.Range.Start, pre_section.Range.End-1).Delete(1)
    #字体变小
    doc.Range(0,0).Select()
    while msword.Selection.Find.Execute(DW, False, False, True, False, False, True, 0, True, "", 0):
        msword.Selection.Font.Subscript = True  
    path = os.getcwd()+"\\打印\\案卷标题"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    saveAndOpen(msword,doc,path)
    
    

def doCaiWuPrint(self,printList=[(u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test',u'test')]):
    DIQU,DANWEI = DD()
    DW = u''
    msword = DispatchEx(r'Word.Application')
    msword.Visible = CESHI
    #msword.DisplayAlerts = CESHI
    doc = msword.Documents.Add()
    for section_index in range(1,len(printList)+1) :
        pre_section = doc.Sections(section_index)
        new_seciton = doc.Range(pre_section.Range.End-1, pre_section.Range.End-1).Sections.Add()
        new_range = new_seciton.Range
        new_table = new_range.Tables.Add(doc.Range(new_range.End-1,new_range.End-1), 4, 1)
        for i in range(1,5):
            new_table.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            new_table.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1           
        new_table.Cell(2,1).Range.InsertAfter(new_table.Cell(2,1).Range.Tables.Add(new_table.Cell(2,1).Range,3,2))
        sub_table_top = new_table.Cell(2,1).Range.Tables.Add(new_table.Cell(2,1).Range,3,2)
        for i in range(1,4):
            sub_table_top.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_top.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1 
        for i in range(1,3):
            sub_table_top.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_top.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1                 
        sub_table_top.Cell(1,1).Range.Text = u"全宗号："
        sub_table_top.Cell(2,1).Range.Text = u"档  号："
        sub_table_top.Cell(3,1).Range.Text = u"索引号："
        DW = printList[section_index-1][0][-1] + unicode(DIQU)
        sub_table_top.Cell(1,2).Range.Text = printList[section_index-1][0]+unicode(DIQU)+u'-'+printList[section_index-1][1]+u'-'+printList[section_index-1][2]#"全宗号值："
        sub_table_top.Cell(2,2).Range.Text = printList[section_index-1][3]+u'-'+printList[section_index-1][4]+u'-'+printList[section_index-1][5]+u'-'+printList[section_index-1][6]#"档号值："
        sub_table_top.Cell(3,2).Range.Text = printList[section_index-1][7]+printList[section_index-1][8]+printList[section_index-1][9]+printList[section_index-1][10]#"索引号值："
        sub_table_top.Columns(1).SetWidth(2.5*28.35,0)
        sub_table_top.Columns(2).SetWidth(4*28.35,0)
        sub_table_top.Rows.Alignment = 0
        sub_table_bottom = new_table.Cell(4,1).Range.Tables.Add(new_table.Cell(4,1).Range,1,2)
        for i in range(1,4):
            sub_table_bottom.Rows.Add()
        for i in range(1,5):
            sub_table_bottom.Rows(i).Borders(constants.wdBorderBottom).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_bottom.Rows(i).Borders(constants.wdBorderRight).LineStyle = 1 
        for i in range(1,3):
            sub_table_bottom.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderTop).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderLeft).LineStyle = 1
            sub_table_bottom.Columns(i).Borders(constants.wdBorderRight).LineStyle = 1         
        sub_table_bottom.Cell(1,1).Range.Text = u"立卷单位:"
        sub_table_bottom.Cell(2,1).Range.Text = u"起止日期:"
        sub_table_bottom.Cell(3,1).Range.Text = u"保管期限:"
        sub_table_bottom.Cell(4,1).Range.Text = u"密    级:"
        sub_table_bottom.Cell(1,2).Range.Text = printList[section_index-1][11]#"立卷单位值"
        #设置临时日期变量
        temp_date = unicode(printList[section_index-1][12])
        if temp_date == u'None' :
            temp_date = u''
        else:
            temp_date = temp_date[:4]+u'年'+temp_date[4:6]+u'月'+temp_date[6:8]+u'日' + u'--' +\
                        temp_date[9:13]+u'年'+temp_date[13:15]+u'月'+temp_date[15:]+u'日'
        sub_table_bottom.Cell(2,2).Range.Text = temp_date #"起止日期值"
        sub_table_bottom.Cell(3,2).Range.Text = printList[section_index-1][13]#"保管期限值"
        sub_table_bottom.Cell(4,2).Range.Text = printList[section_index-1][14]#"密级值"
        sub_table_bottom.Columns(1).SetWidth(2.5*28.35,0)
        sub_table_bottom.Columns(2).SetWidth(6.5*28.35,0)
        sub_table_bottom.Rows.Alignment = 1
        new_table.Cell(2,1).Range.Font.Size = 10
        new_table.Rows(1).Borders(constants.wdBorderBottom).LineStyle = 0
        new_table.Cell(3,1).Range.Font.Size = 18
        new_table.Cell(3,1).Range.Text = "\n\n"+printList[section_index-1][15]#"案卷题目值"
        new_table.Cell(3,1).Range.ParagraphFormat.Alignment = 1
        new_table.Cell(4,1).Range.Font.Size = 10
        new_table.Rows(1).Height = 3
        new_table.Rows(2).Height = 45
        new_table.Rows(2).Borders(constants.wdBorderBottom).LineStyle =0
        new_table.Rows(3).Height = 160
        new_table.Rows(3).Borders(constants.wdBorderBottom).LineStyle = 0
        new_table.Rows(4).Height = 75
        #字号调整sub_table_top.Cell(1,2).Range.Font.Size = 5
    #删除第一页空白
    pre_section = doc.Sections(1)
    doc.Range(pre_section.Range.Start, pre_section.Range.End-1).Delete(1)
    #字体变小
    doc.Range(0,0).Select()
    while msword.Selection.Find.Execute(DW, False, False, True, False, False, True, 0, True, "", 0):
        msword.Selection.Font.Subscript = True
    path = os.getcwd()+"\\打印\\财务用表"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    saveAndOpen(msword,doc,path) 

#索引号、互见号、档案内文件数、档案内文件总页数、立卷人、立卷日期
def doBeiKaoBiaoPrint(self,printList=[('test','test','test','test','test','test','test','test','test')]):
    DIQU,DANWEI = DD()
    msword = DispatchEx(r'Word.Application')
    msword.Visible = CESHI
    #msword.DisplayAlerts = CESHI
    doc = msword.Documents.Add()
    pre_section = doc.Sections(1)
    new_seciton = doc.Range(pre_section.Range.End-1, pre_section.Range.End-1).Sections.Add()
    new_range = new_seciton.Range
    new_range.Text = u"                       卷内备考\n"
    new_range.Font.Size = 15
    new_range.ParagraphFormat.Alignment = 0
    new_table = new_range.Tables.Add(doc.Range(new_range.End-1,new_range.End-1), 1, 1)
    new_table.Rows(1).Borders(constants.wdBorderBottom).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderTop).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderLeft).LineStyle = 1
    new_table.Rows(1).Borders(constants.wdBorderRight).LineStyle = 1
    new_table.Rows(1).Height = 640
    new_table.Cell(1,1).Range.Font.Size = 12
    new_table.Cell(1,1).Range.Text = u"档号:" + printList[0][0]+printList[0][1]+printList[0][2]+printList[0][3] +u"\n\n\n" +u"互见号:" + printList[0][4] + u"\n\n\n" + \
                                     u"说明:" + u"\n\n\n" +u"    卷内共有\n"+u"            文字材料 " + unicode(printList[0][5]) + u" 件，" + \
                                     u"共 "+ unicode(int(printList[0][6])) + u" 页\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n" + \
                                     u"                                            立卷人:" + printList[0][7] + u"\n\n" + \
                                     u"                                            立卷日期:" + printList[0][8] + u"\n\n" + \
                                     u"                                            检查人:\n\n\n"+ \
                                     u"                                                    年    月    日"
    #删除第一页空白
    pre_section = doc.Sections(1)
    doc.Range(pre_section.Range.Start, pre_section.Range.End-1).Delete(1)    
    path = os.getcwd()+"\\打印\\卷内备考表"+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".doc"
    saveAndOpen(msword,doc,path)
    
#存储、关闭文档并打开文档所在文件夹
def saveAndOpen(msword,doc,path=os.getcwd()):
    temp_path = unicode(path)
    if not os.path.exists(os.path.dirname(temp_path)) :
        os.mkdir(os.path.dirname(temp_path))
    doc.SaveAs(temp_path)
    msword.Quit()
    #想要正确打开必须将utf-8转码为系统识别的gbk
    os.system("explorer.exe %s" % os.path.dirname(temp_path).decode("utf-8").encode('gbk'))

