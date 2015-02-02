#-*- coding:utf-8 -*-
import wx
import wx.grid
import DoCURD

CONN  = u''
ARCHVIEID = u''

class myArchiveShowFilesTable(wx.grid.PyGridTableBase):
    def __init__(self):
        wx.grid.PyGridTableBase.__init__(self)
        #self.dataListForFiles = DoCURD.query_all_values(self.CONN, paramDict={"table_name":"archives"})
        #初始化时显示空的20行表格
        paramDict = {}
        paramDict['table_name'] = 'files'
        paramDict['archiveID'] = ARCHVIEID
        paramDict['wenjiantimu'] = u''
        paramDict['wenjianbianhao'] = u'' 
        self.dataListForFiles = DoCURD.query_some_values_from_filestable2(CONN,paramDict)
        needAddRows = 100 - len(self.dataListForFiles)
        if needAddRows > 0 :
            for i in range(needAddRows) :
                self.dataListForFiles.append(['','','','','','','','','','',''])
            
        #列标签
        self.colLabels = [u"文  件  ID",u"档  案  ID",u'案    卷    题    名',u'        文      件      题      目        ',u'    文      件      编      号    ',u'  发  文  单  位  ',
                          u'    形   成   日   期    ',u'页    数',u'密    级',u'备    注',u'录  入  人']
    # these five are the required methods
    def GetNumberRows(self):
        return len(self.dataListForFiles)

    def GetNumberCols(self):
        return len(self.dataListForFiles[0])

    def IsEmptyCell(self, row, col):
        return True
    def GetValue(self, row, col):
        return self.dataListForFiles[row][col]
        
    def SetValue(self, row, col, value):
        #设置值
        #self.dataListForFiles[row][col] = value
        #如果表是档案查询页的表则不允许设置值
        #if self == app.frame.select_data_table :
            #pass
        #else:
            #self.dataListForFiles[row][col] = value
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
        if self.dataListForFiles is None or len(self.dataListForFiles) == 0:
            return False
        #for rowNum in range(0,numRows):
            #self.dataListForFiles.remove(self.dataListForFiles[pos+rowNum]) 
        gridView = self.GetView()
        gridView.BeginBatch()
        deleteMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED,pos,numRows)
        gridView.ProcessTableMessage(deleteMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)
           
        return True

class ShowDialog(wx.Dialog):
    def __init__(self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition,style=wx.DEFAULT_DIALOG_STYLE):
        pre = wx.PreDialog()
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)    
        #生成数据显示控件和数据源表格
        self.dataListForFiles_forShow = wx.grid.Grid(parent=self,size=(1015,505))
        self.select_data_file_table = myArchiveShowFilesTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_select_file_table_rows_forShow = self.select_data_file_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.dataListForFiles_forShow.SetTable(self.select_data_file_table,takeOwnership=True)
        self.dataListForFiles_forShow.AutoSize()
           
           
class ShowFilesPanel(wx.Panel):
    def __init__(self,parent):
        wx.Panel.__init__(self, parent, -1)
        dlg = ShowDialog(self, -1, u"该档案下所含文件:", size=(800,535),
                         style=wx.DEFAULT_DIALOG_STYLE)
        dlg.CenterOnScreen()
        val = dlg.ShowModal()
        dlg.Destroy()
            

        