# -*- coding:utf-8 -*-
import os
import wx
import pickle


DANWEI = u''
class SingleDialog(wx.Dialog):
    def __init__(
            self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition, 
            style=wx.DEFAULT_DIALOG_STYLE):
        pre = wx.PreDialog()
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)
        myList = []
        targetFile = DANWEI + u'.tbdata'
        #载入数据
        if os.path.exists(targetFile) :
            with open(targetFile,'rb') as loadfile :
                myLIst = pickle.load(loadfile)
        
        self.notebook = wx.Notebook(parent=self,size=(800,600))
        self.archivesPanel = wx.Panel(parent=self.notebook)
        
        
        
        
        self.filesPanel = wx.Panel(parent=self.notebook)


        
        self.notebook.AddPage(self.archivesPanel,u"档案")
        self.notebook.AddPage(self.filesPanel,u'文件')
        self.Show()
        self.Center()
    
class myArchiveTable(wx.grid.PyGridTableBase):
    def __init__(self):
        wx.grid.PyGridTableBase.__init__(self)
        #self.dataList = DoCURD.query_all_values(self.CONN, paramDict={"table_name":"archives"})
        #初始化时显示空的20行表格
        self.dataList = []
        needAddRows = 20 - len(self.dataList)
        if needAddRows > 0 :
            for i in range(needAddRows) :
                self.dataList.append(['','','','','','','','','','','','','','','','','','',''])
            
        #列标签
        self.colLabels = [u"档案ID",u"门    类",u'分    类',u'立  卷  单  位',u'案  卷  题  名',u'立  卷  日  期',u'位    置',u'密    级',u'责  任  人',
                          u'区    号',u'柜    号',u'盒    号',u'卷    号',u'互  见  号',u'科    目',u'期    限',u'备    注',u'录  入  人',u'是  否  归  档']     

    # these five are the required methods
    def GetNumberRows(self):
        return len(self.dataList)

    def GetNumberCols(self):
        return len(self.dataList[0])

    def IsEmptyCell(self, row, col):
        return True
    def GetValue(self, row, col):
        return self.dataList[row][col]
        
    def SetValue(self, row, col, value):
        #设置值
        #self.dataList[row][col] = value
        #如果表是档案查询页的表则不允许设置值
        #if self == app.frame.select_data_table :
            #pass
        #else:
            #self.dataList[row][col] = value
        pass
            
    def GetColLabelValue(self,col):
        return self.colLabels[col]
    
    def AppendRows(self,numRows=1):
        self.isModified = True
        gridView = self.GetView()
        gridView.BeginBatch()
        appendMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_APPENDED,numRows)
        gridView.ProcessTableMessage(appendMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)
                   
        return True
    
    def DeleteRows(self,pos=0,numRows=1):
        if self.dataList is None or len(self.dataList) == 0:
            return False
        #for rowNum in range(0,numRows):
            #self.dataList.remove(self.dataList[pos+rowNum]) 
        gridView = self.GetView()
        gridView.BeginBatch()
        deleteMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED,pos,numRows)
        gridView.ProcessTableMessage(deleteMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)
           
        return True

#用于构造文件表格的类
class myFileTable(wx.grid.PyGridTableBase):
    def __init__(self):
        wx.grid.PyGridTableBase.__init__(self)
        #self.dataList = DoCURD.query_all_values(self.CONN, paramDict={"table_name":"archives"})
        #初始化时显示空的20行表格
        self.dataList_File = []
        needAddRows_File = 20 - len(self.dataList_File)
        if needAddRows_File > 0 :
            for i in range(needAddRows_File) :
                self.dataList_File.append(['','','','','','','','','','',''])
            
        #列标签
        self.colLabels_File = [u"文  件  ID",u"档  案  ID",u'案    卷    题    名',u'        文      件      题      目        ',u'    文      件      编      号    ',u'  发  文  单  位  ',
                          u'    形   成   日   期    ',u'页    数',u'密    级',u'备    注',u'录  入  人']     

    # these five are the required methods
    def GetNumberRows(self):
        return len(self.dataList_File)

    def GetNumberCols(self):
        return len(self.dataList_File[0])

    def IsEmptyCell(self, row, col):
        return True
    def GetValue(self, row, col):
        return self.dataList_File[row][col]
        
    def SetValue(self, row, col, value):
        #设置值
        #self.dataList[row][col] = value
        #如果表是档案查询页的表则不允许设置值
        #if self == app.frame.select_data_table :
            #pass
        #else:
            #self.dataList[row][col] = value
        pass 


class SinglePanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        dlg = SingleDialog(self, -1, u"", size=(800,600),
                         style=wx.DEFAULT_DIALOG_STYLE)
        dlg.SetLabel(DANWEI+u'  上报数据')
        dlg.CenterOnScreen()
        val = dlg.ShowModal()
        dlg.Destroy()

    
    
