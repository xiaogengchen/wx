# -*- coding:utf-8 -*-
import os
import wx
import wx.grid
import wx.combo
import UtilData
import DoCURD
import ShowFiles

DANWEI = u''
DATALIST = []

class SingleDialog(wx.Dialog):
    def __init__(
            self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition, 
            style=wx.DEFAULT_DIALOG_STYLE):
        pre = wx.PreDialog()
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)
        myList = []
        datefile = u''
        #进行数据库连接
        for i in DATALIST:
            if DANWEI in i:
                datafile = i
                break
        self.CONN = DoCURD.connect_db(u".\\上报数据\\"+datafile)
        #生成notebook框架
        self.notebook = wx.Notebook(parent=self,size=(1024,620))
        #生成档案面板(包含在notebook框架中)
        self.archivesPanel = wx.Panel(parent=self.notebook)
        #-----------------------------------------------------------------------
        select_archives_vbox = wx.BoxSizer(wx.VERTICAL)
        select_archives_hbox = wx.BoxSizer(wx.HORIZONTAL)
        select_archives_hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        #区号
        select_archives_quhaoText = wx.StaticText(parent=self.archivesPanel,label=u"区号:")
        #生成  区号  下拉选择框
        self.quhao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,choices=[""]+UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.quhao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.quhao_select_archives_comboBox.SetPopupMaxHeight(self.quhao_select_archives_comboBox.GetCharHeight()*10) 
        #柜号
        select_archives_guihaoText = wx.StaticText(parent=self.archivesPanel,label=u"柜号:")
        self.guihao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,choices=[""]+UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guihao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guihao_select_archives_comboBox.SetPopupMaxHeight(self.guihao_select_archives_comboBox.GetCharHeight()*10)       
        #盒号
        select_archives_hehaoText = wx.StaticText(parent=self.archivesPanel,label=u"盒号:")
        self.hehao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=(43,25),choices=[""]+UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.hehao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.hehao_select_archives_comboBox.SetPopupMaxHeight(self.hehao_select_archives_comboBox.GetCharHeight()*10)
        #卷号
        select_archives_juanhaoText = wx.StaticText(parent=self.archivesPanel,label=u"卷号:")
        self.juanhao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=(43,25),choices=[""]+UtilData.JUANHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.juanhao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.juanhao_select_archives_comboBox.SetPopupMaxHeight(self.juanhao_select_archives_comboBox.GetCharHeight()*10)       
        #***********************************高级检索begine********************************************
        SENIOR_SELECT_SIZE = (43,25)
        #门类
        select_archives_menleiText = wx.StaticText(parent=self.archivesPanel,label=u"门类:")
        self.menleiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MENLEI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.menleiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.menleiText_select_archives_comboBox.SetPopupMaxHeight(self.menleiText_select_archives_comboBox.GetCharHeight()*10)                
        #绑定门类
        self.menleiText_select_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceGuiDangForSeniorSelect)
        #归档(分类)
        select_archives_guidangText = wx.StaticText(parent=self.archivesPanel,label=u"分类:")
        self.guidangText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""],style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guidangText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guidangText_select_archives_comboBox.SetPopupMaxHeight(self.guidangText_select_archives_comboBox.GetCharHeight()*10)                
        #案卷题名
        select_archives_anjuantimingText = wx.StaticText(parent=self.archivesPanel,label=u"案卷题名:")
        self.anjuantiming_select_archives_Ctrl = wx.TextCtrl(parent=self.archivesPanel,size=(134,25))
        #期限
        select_archives_qixianText = wx.StaticText(parent=self.archivesPanel,label=u"期限:")
        self.qixianText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.QIXIAN_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.qixianText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.qixianText_select_archives_comboBox.SetPopupMaxHeight(self.qixianText_select_archives_comboBox.GetCharHeight()*10)        
        #立卷单位
        select_archives_lijuandanweiText = wx.StaticText(parent=self.archivesPanel,label=u"单位:")
        self.lijuandanweiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[i for i in UtilData.DANWEI_List if DANWEI in i],style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.lijuandanweiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.lijuandanweiText_select_archives_comboBox.SetPopupMaxHeight(self.lijuandanweiText_select_archives_comboBox.GetCharHeight()*10)         
        #立卷日期
        select_archives_lijuanriqiText = wx.StaticText(parent=self.archivesPanel,label=u"立卷日期:")
        self.lijuanriqiText_select_archives_pop = UtilData.DateControl(self.archivesPanel, -1)       
        #录入人
        select_archives_inputterText = wx.StaticText(parent=self.archivesPanel,label=u"录入人:")
        self.inputterText_select_archives_Ctrl = wx.TextCtrl(parent=self.archivesPanel)  
        #存放位置
        select_archives_weizhiText = wx.StaticText(parent=self.archivesPanel,label=u"存放位置:")
        self.weizhiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.WEIZHI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.weizhiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.weizhiText_select_archives_comboBox.SetPopupMaxHeight(self.weizhiText_select_archives_comboBox.GetCharHeight()*10)                 
        #互见号
        select_archives_hujianhaoText = wx.StaticText(parent=self.archivesPanel,label=u"互见号:")
        self.hujianhaoText_select_archives_Ctrl = wx.TextCtrl(parent=self.archivesPanel)        
        #档案科目
        select_archives_kemuText = wx.StaticText(parent=self.archivesPanel,label=u"档案科目:")
        self.kemuText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.KEMU_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.kemuText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.kemuText_select_archives_comboBox.SetPopupMaxHeight(self.kemuText_select_archives_comboBox.GetCharHeight()*10)        
        #密级
        select_archives_mijiText = wx.StaticText(parent=self.archivesPanel,label=u"密级:")
        self.mijiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.archivesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.mijiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.mijiText_select_archives_comboBox.SetPopupMaxHeight(self.mijiText_select_archives_comboBox.GetCharHeight()*10)        
        ##档案ID
        #select_archives_archiveIDText = wx.StaticText(parent=self.archivesPanel,label=u"档案ID:")
        #self.archiveIDText_select_archives_Ctrl = wx.TextCtrl(parent=self.archivesPanel,size=(122,25))        
        #查询按钮
        select_archives_button = wx.Button(parent=self.archivesPanel,label=u"查询",size=(75,28))
        select_archives_showFiles_button = wx.Button(parent=self.archivesPanel,label=u"查看",size=(75,28))
        #对查询按钮进行事件绑定
        select_archives_button.Bind(wx.EVT_BUTTON, self.selectArchivesBySomeConditions)
        select_archives_showFiles_button.Bind(wx.EVT_BUTTON, self.showFiles)        
        #向顶部横向布局管理器中添加元素
        select_archives_hbox.Add(select_archives_quhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=15)
        select_archives_hbox.Add(self.quhao_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox.Add(select_archives_guihaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.guihao_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox.Add(select_archives_hehaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.hehao_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox.Add(select_archives_juanhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.juanhao_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox.Add(select_archives_anjuantimingText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.anjuantiming_select_archives_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox.Add(select_archives_hujianhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.hujianhaoText_select_archives_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        select_archives_hbox.Add(select_archives_kemuText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.kemuText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)                
        #select_archives_hbox.Add(select_archives_archiveIDText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        #select_archives_hbox.Add(self.archiveIDText_select_archives_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        select_archives_hbox.Add(select_archives_mijiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox.Add(self.mijiText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        
        select_archives_hbox2.Add(select_archives_menleiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=15)
        select_archives_hbox2.Add(self.menleiText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox2.Add(select_archives_guidangText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.guidangText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox2.Add(select_archives_lijuandanweiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.lijuandanweiText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        select_archives_hbox2.Add(select_archives_qixianText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.qixianText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        select_archives_hbox2.Add(select_archives_lijuanriqiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.lijuanriqiText_select_archives_pop,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5) 
        select_archives_hbox2.Add(select_archives_inputterText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.inputterText_select_archives_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        select_archives_hbox2.Add(select_archives_weizhiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(self.weizhiText_select_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)        
        
        select_archives_hbox2.Add(select_archives_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        select_archives_hbox2.Add(select_archives_showFiles_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        #向纵向布局管理器中添加横向布局管理器
        select_archives_vbox.Add(select_archives_hbox)
        select_archives_vbox.Add(select_archives_hbox2)
        #***********************************高级检索end********************************************
        #生成数据显示控件和数据源表格
        self.select_grid_archives = wx.grid.Grid(parent=self.archivesPanel,size=(1015,505))
        self.select_data_table = myArchivesTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_select_table_rows = self.select_data_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.select_grid_archives.SetTable(self.select_data_table,takeOwnership=True)
        self.select_grid_archives.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.activeRow)
        self.select_grid_archives.AutoSize()
        #向纵向布局管理器中添加数据表格
        select_archives_vbox.Add(self.select_grid_archives,flag=wx.ALIGN_BOTTOM|wx.TOP,border=2)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.archivesPanel.SetSizer(select_archives_vbox)           
        
        #-----------------------------------------------------------------------        
        #生成文件面板(包含在notebook框架中)
        self.filesPanel = wx.Panel(parent=self.notebook)
        #档案查询页中的总垂直布局管理器和顶部横向布局管理器
        select_files_vbox = wx.BoxSizer(wx.VERTICAL)
        select_files_hbox = wx.BoxSizer(wx.HORIZONTAL)
        #档案ID
        #查询档案表中共有多少档案ID,都是什么
        temp_archiveIDList = DoCURD.query_archiveID_for_selectfile(self.CONN,paramDict={"table_name":"archives"})
        archiveIDs = ['']    
        for row in temp_archiveIDList:
            for col in row :
                archiveIDs.append(unicode(col))
        select_files_archiveIDText = wx.StaticText(parent=self.filesPanel,label=u'档案ID:')
        self.archiveID_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.filesPanel,id=-1,choices=archiveIDs,style=wx.CB_READONLY)
        self.archiveID_select_files_comboBox.SetSelection(n=0)
        self.archiveID_select_files_comboBox.SetPopupMaxHeight(self.archiveID_select_files_comboBox.GetCharHeight()*10)
        SENIOR_SELECT_SIZE_FOR_FILES=(80,25)
        #文件编号
        select_files_wenjianbianhaoText = wx.StaticText(parent=self.filesPanel,label=u"文件编号:")
        self.wenjianbianhao_select_files_ctrl = wx.TextCtrl(parent=self.filesPanel,size=SENIOR_SELECT_SIZE_FOR_FILES)
        #文件题目
        select_files_wenjiantimuText = wx.StaticText(parent=self.filesPanel,label=u"文件题目:")
        self.wenjiantimu_select_files_Ctrl = wx.TextCtrl(parent=self.filesPanel,size=SENIOR_SELECT_SIZE_FOR_FILES)
        #发文单位
        select_files_fawendangweiText = wx.StaticText(parent=self.filesPanel,label=u"发文单位:")
        self.fawendanwei_select_files_Ctrl = wx.TextCtrl(parent=self.filesPanel,size=SENIOR_SELECT_SIZE_FOR_FILES)  
        #形成日期
        select_files_xingchengriqiText = wx.StaticText(parent=self.filesPanel,label=u"形成日期:")
        self.xingchengriqiText_select_files_pop = UtilData.DateControl(self.filesPanel)               
        #密级
        select_files_mijiText = wx.StaticText(parent=self.filesPanel,label=u"密级:")
        self.miji_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.filesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MIJI_List,style=wx.CB_READONLY)
        self.miji_select_files_comboBox.SetSelection(n=0)
        self.miji_select_files_comboBox.SetPopupMaxHeight(self.miji_select_files_comboBox.GetCharHeight()*10)        
        #存放位置
        select_files_weizhiText = wx.StaticText(parent=self.filesPanel,label=u"位置:")
        self.weizhi_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.filesPanel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.WEIZHI_List,style=wx.CB_READONLY)
        self.weizhi_select_files_comboBox.SetSelection(n=0)
        self.weizhi_select_files_comboBox.SetPopupMaxHeight(self.weizhi_select_files_comboBox.GetCharHeight()*10)        
        #备注
        select_files_beizhuText = wx.StaticText(parent=self.filesPanel,label=u"备注:")
        self.beizhu_select_files_Ctrl = wx.TextCtrl(parent=self.filesPanel,size=SENIOR_SELECT_SIZE_FOR_FILES)          
        
        #查询按钮
        select_files_button = wx.Button(parent=self.filesPanel,label=u"查询")
        #对查询按钮进行事件绑定
        select_files_button.Bind(wx.EVT_BUTTON, self.selectFilesBySomeConditions)
        #向顶部横向布局管理器中添加元素
        select_files_hbox.Add(select_files_archiveIDText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)
        select_files_hbox.Add(self.archiveID_select_files_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)               
        select_files_hbox.Add(select_files_wenjianbianhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.wenjianbianhao_select_files_ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)
        select_files_hbox.Add(select_files_wenjiantimuText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.wenjiantimu_select_files_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)
        select_files_hbox.Add(select_files_fawendangweiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.fawendanwei_select_files_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1) 
        select_files_hbox.Add(select_files_xingchengriqiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.xingchengriqiText_select_files_pop,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)        
        select_files_hbox.Add(select_files_mijiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.miji_select_files_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)        
        select_files_hbox.Add(select_files_weizhiText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.weizhi_select_files_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1) 
        select_files_hbox.Add(select_files_beizhuText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=3)
        select_files_hbox.Add(self.beizhu_select_files_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)         
        
        select_files_hbox.Add(select_files_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=1)
        #向纵向布局管理器中添加横向布局管理器
        select_files_vbox.Add(select_files_hbox,flag=wx.TOP,border=10)
        #生成数据显示控件和数据源表格
        self.select_grid_files = wx.grid.Grid(parent=self.filesPanel,size=(1015,505))
        self.select_data_file_table = myFilesTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_select_file_table_rows = self.select_data_file_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.select_grid_files.SetTable(self.select_data_file_table,takeOwnership=True)
        self.select_grid_files.AutoSize()
        #向纵向布局管理器中添加数据表格
        select_files_vbox.Add(self.select_grid_files,flag=wx.ALIGN_BOTTOM|wx.TOP,border=15)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.filesPanel.SetSizer(select_files_vbox)                
        #---------------------------------------------------------------------------
        self.notebook.AddPage(self.archivesPanel,u"档案")
        self.notebook.AddPage(self.filesPanel,u'文件')
        self.Show()
        self.Center()
        
    def choiceGuiDangForSeniorSelect(self,event):
        TEMP = self.menleiText_select_archives_comboBox.GetValue().split(u"_")[:1]
        TEMP = "".join(TEMP)
        #print TEMP,type(TEMP)
        GUIDANG_MENLEI_List = []
        if TEMP == u"" :
            GUIDANG_MENLEI_List = [u'']
        if TEMP in u'MENLEI_W_List'.split(u"_") :
            GUIDANG_MENLEI_List = [u'']+UtilData.MENLEI_W_List
        if TEMP in u'MENLEI_J_List'.split(u"_"):
            GUIDANG_MENLEI_List = [u'']+UtilData.MENLEI_J_List   
        if TEMP in u'MENLEI_K_List'.split(u"_"):
            GUIDANG_MENLEI_List = [u'']+UtilData.MENLEI_K_List
        if TEMP in u'MENLEI_Y_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_Y_List 
        if TEMP in u'MENLEI_S_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_S_List
        if TEMP in u'MENLEI_D_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_D_List 
        #self.guidang_archives_comboBox.Clear()
        #根据门类不同的选择重新设置comboBox的choice值
        self.guidangText_select_archives_comboBox.Set(GUIDANG_MENLEI_List)
        #默认第0个元素选中状态
        self.guidangText_select_archives_comboBox.SetSelection(0)  
            
    #选中当前所选单元格所在行且激活当前选中行中所选单元格
    def activeRow(self,event):
        #选中当前所选单元格所在行
        event.GetEventObject().SelectRow(row=event.GetRow())
        #激活当前选中行中所选单元格
        event.GetEventObject().SetGridCursor(row=event.GetRow(),col=event.GetCol())   
    
    def showFiles(self,event):
        active_row_number = self.select_grid_archives.GetGridCursorRow()
        #获得指定单元格的值GetCellValue(row,col)
        ShowFiles.CONN = self.CONN
        ShowFiles.ARCHVIEID = self.select_grid_archives.GetCellValue(active_row_number,0)
        if u''!=ShowFiles.ARCHVIEID:
            ShowFiles.ShowFilesPanel(self)    
    
    #在档案查询页根据给定的一些条件(区号、柜号、盒号、卷号、案卷题名)进行查询
    def selectArchivesBySomeConditions(self,event):
        #为了判断日期是否准确临时生成一个temp_cal变量
        temp_cal = self.lijuanriqiText_select_archives_pop.GetValue().strip().decode('utf-8')
        if temp_cal != u'' :
            if len(temp_cal) != 11 or (temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日") :
                self.statusBar.SetStatusText(u"立卷日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！")
                errorMessage = wx.MessageDialog(parent=None,message=u"立卷日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"档案查询异常信息",style=wx.ICON_ERROR)
                result = errorMessage.ShowModal()
                errorMessage.Destroy()               
                return        
        paramDict = {"table_name":"archives"}
        paramDict['quhao'] = unicode(self.quhao_select_archives_comboBox.GetValue().strip())
        paramDict['guihao'] = unicode(self.guihao_select_archives_comboBox.GetValue().strip())
        paramDict['hehao'] = unicode(self.hehao_select_archives_comboBox.GetValue().strip())
        paramDict['juanhao'] = unicode(self.juanhao_select_archives_comboBox.GetValue().strip())
        paramDict['anjuantiming'] = unicode(self.anjuantiming_select_archives_Ctrl.Value.strip())
        paramDict['menlei'] = unicode(self.menleiText_select_archives_comboBox.GetValue().strip())
        paramDict['guidang'] = unicode(self.guidangText_select_archives_comboBox.GetValue().strip())
        paramDict['qixian'] = unicode(self.qixianText_select_archives_comboBox.GetValue().strip())
        paramDict['danwei'] = unicode(self.lijuandanweiText_select_archives_comboBox.GetValue().strip())
        paramDict['lijuanriqi'] = unicode(self.lijuanriqiText_select_archives_pop.GetValue().strip())
        paramDict['inputter'] = unicode(self.inputterText_select_archives_Ctrl.Value.strip())
        paramDict['hujianhao'] = unicode(self.hujianhaoText_select_archives_Ctrl.Value.strip())
        paramDict['weizhi'] = unicode(self.weizhiText_select_archives_comboBox.GetValue().strip())
        paramDict['kemu'] = unicode(self.kemuText_select_archives_comboBox.GetValue().strip())
        paramDict['miji'] = unicode(self.mijiText_select_archives_comboBox.GetValue().strip())
        #paramDict['archiveID'] = unicode(self.archiveIDText_select_archives_Ctrl.Value.strip())
        #查询数据库中给定条件的档案,结果替换原有数据源
        self.select_data_table.dataList = DoCURD.query_some_values_from_archivetable(self.CONN,paramDict=paramDict)
        #临时存放现有表格行数
        temp_num = len(self.select_data_table.dataList)
        needAddRows = 20 - temp_num
        #用20减去临时存放表格的行数为需要添加的行数
        #需要添加的行数如果为正值证明记录数少于20,少几行补几行,再删除上次表中多余行数(即上次表行数减20后的数)
        if needAddRows > 0 :
            for i in range(needAddRows) :
                self.select_data_table.dataList.append(['','','','','','','','','','','','','','','','','','',''])
            #如果查询到的数据不足一屏则将原有表格容器多余的行数删除(删除的个数为self.select_data_table.GetRowsCount()-默认的20个)
            self.select_data_table.DeleteRows(numRows=(int(self.last_data_select_table_rows)-20))
            #更新上一次表格行数
            self.last_data_select_table_rows = self.select_data_table.GetRowsCount()
        #需要添加的行数如果为负值证明记录数多余20,将现有表的行数和上一次表的行数相比对
        else:
            #如果现在表的行数小于上一次表的行数则删除容器中多余的行
            if int(temp_num) < int(self.last_data_select_table_rows) :
                # #如果超过一屏20个且上次表格行数超过20个则删（现有表格容器的行数-列表元素个数(查询到的记录数)）行
                self.select_data_table.DeleteRows(numRows=abs(temp_num-int(self.last_data_select_table_rows)))
            #如果现在表的行数大于上一次表的行数则给容器增加多出的行
            else:
                 #如果超过一屏20个且上次表格行数不足20个则加（列表元素个数(查询到的记录数)-现有表格容器的行数）行
                self.select_data_table.AppendRows(abs(temp_num-int(self.last_data_select_table_rows)))
            #更新上一次表格行数
            self.last_data_select_table_rows = self.select_data_table.GetRowsCount()            
        #对查询后的数据进行刷新,否则不显示
        self.select_grid_archives.Refresh()
        #自动调整行宽列宽
        self.select_grid_archives.AutoSizeColumns()
        self.select_grid_archives.AutoSizeRows()        
        wx.MessageBox(u"查询到档案 " + unicode(temp_num) + u" 个",u'查询结果')   

    def selectFilesBySomeConditions(self,event):
        temp_cal = self.xingchengriqiText_select_files_pop.GetValue().strip().decode('utf-8')
        if temp_cal != u'' :
            if len(temp_cal)!=11 or temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日" :
                self.statusBar.SetStatusText(u"形成日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！")
                errorMessage = wx.MessageDialog(parent=None,message=u"形成日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
                result = errorMessage.ShowModal()
                errorMessage.Destroy()               
                return        
        paramDict = {"table_name":"files"}
        paramDict['archiveID'] = unicode(self.archiveID_select_files_comboBox.GetValue())
        paramDict['wenjiantimu'] = unicode(self.wenjiantimu_select_files_Ctrl.Value.strip())
        paramDict['wenjianbianhao'] = unicode(self.wenjianbianhao_select_files_ctrl.Value.strip())
        paramDict['fawendanwei'] = unicode(self.fawendanwei_select_files_Ctrl.Value.strip())
        paramDict['xingchengriqi'] = unicode(self.xingchengriqiText_select_files_pop.GetValue().strip())
        paramDict['miji'] = unicode(self.miji_select_files_comboBox.GetValue().strip())
        paramDict['weizhi'] = unicode(self.weizhi_select_files_comboBox.GetValue().strip())
        paramDict['beizhu'] = unicode(self.beizhu_select_files_Ctrl.Value.strip())
        #查询数据库中给定条件的档案,结果替换原有数据源
        if paramDict['archiveID']!="" :
            self.select_data_file_table.dataList_File = DoCURD.query_some_values_from_filestable(self.CONN,paramDict=paramDict)
        else:
            self.select_data_file_table.dataList_File = DoCURD.query_some_values_from_filestable_noArchiveID(self.CONN,paramDict=paramDict)
        #临时存放现有表格行数
        temp_num_file = len(self.select_data_file_table.dataList_File)
        needAddRows_file = 20 - temp_num_file
        #用20减去临时存放表格的行数为需要添加的行数
        #需要添加的行数如果为正值证明记录数少于20,少几行补几行,再删除上次表中多余行数(即上次表行数减20后的数)
        if needAddRows_file > 0 :
            for i in range(needAddRows_file) :
                self.select_data_file_table.dataList_File.append(['','','','','','','','','','',''])
            #如果查询到的数据不足一屏则将原有表格容器多余的行数删除(删除的个数为self.select_data_table.GetRowsCount()-默认的20个)
            self.select_data_file_table.DeleteRows(numRows=(int(self.last_data_select_file_table_rows)-20))
            #更新上一次表格行数
            self.last_data_select_file_table_rows = self.select_data_file_table.GetRowsCount()
        #需要添加的行数如果为负值证明记录数多余20,将现有表的行数和上一次表的行数相比对
        else:
            #如果现在表的行数小于上一次表的行数则删除容器中多余的行
            if int(temp_num_file) < int(self.last_data_select_file_table_rows) :
                # #如果超过一屏20个且上次表格行数超过20个则删（现有表格容器的行数-列表元素个数(查询到的记录数)）行
                self.select_data_file_table.DeleteRows(numRows=abs(temp_num_file-int(self.last_data_select_file_table_rows)))
            #如果现在表的行数大于上一次表的行数则给容器增加多出的行
            else:
                 #如果超过一屏20个且上次表格行数不足20个则加（列表元素个数(查询到的记录数)-现有表格容器的行数）行
                self.select_data_file_table.AppendRows(abs(temp_num_file-int(self.last_data_select_file_table_rows)))
            #更新上一次表格行数
            self.last_data_select_file_table_rows = self.select_data_file_table.GetRowsCount()            
        #对查询后的数据进行刷新,否则不显示
        self.select_grid_files.Refresh()
        #自动调整行宽列宽
        self.select_grid_files.AutoSizeColumns()
        self.select_grid_files.AutoSizeRows()  
        wx.MessageBox(u"查询到文件 " + unicode(temp_num_file) + u" 个",u'查询结果')   
    
class myArchivesTable(wx.grid.PyGridTableBase):
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
class myFilesTable(wx.grid.PyGridTableBase):
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
    
    def GetColLabelValue(self,col):
        return self.colLabels_File[col]
    
    def AppendRows(self,numRows=1):
        self.isModified = True
        gridView = self.GetView()
        type(gridView)
        gridView.BeginBatch()
        appendMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_APPENDED,numRows)
        gridView.ProcessTableMessage(appendMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)
                   
        return True
    
    def DeleteRows(self,pos=0,numRows=1):
        if self.dataList_File is None or len(self.dataList_File) == 0:
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

class SinglePanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        dlg = SingleDialog(self, -1, u"", size=(1024,620),
                         style=wx.DEFAULT_DIALOG_STYLE)
        dlg.SetLabel(DANWEI+u'  上报数据')
        dlg.CenterOnScreen()
        val = dlg.ShowModal()
        dlg.Destroy()

    
    
