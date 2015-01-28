# -*- coding:utf-8 -*-
import  wx
import UtilData
import DoCURD

temp_list_frame = []
CONN = DoCURD.connect_db('./tbdata.db')
class TestDialog(wx.Dialog):
    def __init__(
            self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition, 
            style=wx.DEFAULT_DIALOG_STYLE):

        pre = wx.PreDialog()
        #pre.SetExtraStyle(wx.DIALOG_EX_CONTEXTHELP)
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)
        #生成修改面板
        #self = wx.Frame(parent=None,title=u'档案修改',size=(450,370))
        #生成GridBagSizer布局管理器
        modify_archives_gridBagSizer = wx.GridBagSizer(vgap=5,hgap=5)
        #生成文本标签
        modify_archives_menleiText = wx.StaticText(parent=self,label=u"门        类:",style=wx.ALIGN_CENTER_VERTICAL)
        #add_archives_jibieText = wx.StaticText(parent=self.add_archives_panel,label=u"级别:",style=wx.ALIGN_CENTER_HORIZONTAL)
        modify_archives_guidangText = wx.StaticText(parent=self,label=u"归        档:",style=wx.ALIGN_CENTER_VERTICAL)
        modify_archives_qixianText = wx.StaticText(parent=self,label=u"期        限:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_anjuantimingText = wx.StaticText(parent=self,label=u"案卷题名:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_danweiText = wx.StaticText(parent=self,label=u"单        位:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_lijuanriqiText = wx.StaticText(parent=self,label=u"立卷日期:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_weizhiText = wx.StaticText(parent=self,label=u"位        置:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_mijiText = wx.StaticText(parent=self,label=u"密        级:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_zerenrenText = wx.StaticText(parent=self,label=u"责  任  人:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_quhaoText = wx.StaticText(parent=self,label=u"区        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_guihaoText = wx.StaticText(parent=self,label=u"柜        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_hehaoText = wx.StaticText(parent=self,label=u"盒        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_juanhaoText = wx.StaticText(parent=self,label=u"卷        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_hujianhaoText = wx.StaticText(parent=self,label=u"互  见  号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_kemuText = wx.StaticText(parent=self,label=u"科        目:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_beizhuText = wx.StaticText(parent=self,label=u"备        注:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_inputterText = wx.StaticText(parent=self,label=u"录  入  人:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_archives_font = wx.Font(pointSize=9,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        modify_archives_menleiText.SetFont(modify_archives_font)
        #add_archives_jibieText.SetFont(modify_archives_font)
        modify_archives_guidangText.SetFont(modify_archives_font)
        modify_archives_qixianText.SetFont(modify_archives_font)
        modify_archives_anjuantimingText.SetFont(modify_archives_font)
        modify_archives_danweiText.SetFont(modify_archives_font)
        modify_archives_lijuanriqiText.SetFont(modify_archives_font)
        modify_archives_weizhiText.SetFont(modify_archives_font)
        modify_archives_mijiText.SetFont(modify_archives_font)
        modify_archives_zerenrenText.SetFont(modify_archives_font)
        modify_archives_quhaoText.SetFont(modify_archives_font)
        modify_archives_guihaoText.SetFont(modify_archives_font)
        modify_archives_hehaoText.SetFont(modify_archives_font)
        modify_archives_juanhaoText.SetFont(modify_archives_font)
        modify_archives_hujianhaoText.SetFont(modify_archives_font)
        modify_archives_kemuText.SetFont(modify_archives_font)
        modify_archives_beizhuText.SetFont(modify_archives_font)
        modify_archives_inputterText.SetFont(modify_archives_font)
        
        #将文本标签纳入add_archives_gridBagSizer布局管理器中
        modify_archives_gridBagSizer.Add(modify_archives_menleiText,pos=(0,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        #add_archives_gridBagSizer.Add(modify_archives_jibieText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_guidangText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_anjuantimingText,pos=(2,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_qixianText,pos=(3,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_danweiText,pos=(4,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_lijuanriqiText,pos=(5,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_weizhiText,pos=(6,0),flag=wx.wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_mijiText,pos=(7,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_zerenrenText,pos=(8,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        
        modify_archives_gridBagSizer.Add(modify_archives_quhaoText,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT|wx.TOP,border=15)
        modify_archives_gridBagSizer.Add(modify_archives_guihaoText,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_hehaoText,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_juanhaoText,pos=(3,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_hujianhaoText,pos=(4,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_kemuText,pos=(5,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_beizhuText,pos=(6,2),span=(2,1),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_archives_gridBagSizer.Add(modify_archives_inputterText,pos=(8,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        
        ARCHIVES_SIZE = (120,25)
        #生成  门类  下拉选择框
        #根据不同门类选项存放不同归档信息的列表
        self.modify_menlei_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[1],size=ARCHIVES_SIZE,choices=UtilData.MENLEI_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_menlei_archives_comboBox.SetPopupMaxHeight(self.modify_menlei_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件,目的是根据选择不同的门类提供不同的归档列表值
        self.modify_menlei_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceGuiDangFrame)

        #生成  归档  下拉选择框
        self.modify_guidang_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[2],size=ARCHIVES_SIZE,choices=UtilData.MENLEI_W_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_guidang_archives_comboBox.SetPopupMaxHeight(self.modify_guidang_archives_comboBox.GetCharHeight()*10)
        
        #生成  期限  下拉选择框
        self.modify_qixian_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[15],size=ARCHIVES_SIZE,choices=UtilData.QIXIAN_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_qixian_archives_comboBox.SetPopupMaxHeight(self.modify_qixian_archives_comboBox.GetCharHeight()*10)
        
        #生成 案卷题名 输入框
        self.modify_anjuantiming_archives_ctrl = wx.TextCtrl(parent=self,value=temp_list_frame[4],size=ARCHIVES_SIZE)
        
        #生成  单位  下拉选择框
        self.modify_danwei_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[3],size=ARCHIVES_SIZE,choices=UtilData.DANWEI_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_danwei_archives_comboBox.SetPopupMaxHeight(self.modify_danwei_archives_comboBox.GetCharHeight()*10)
        ##为下拉选择框绑定事件
        #self.modify_danwei_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForAnJuanHaoFrame)        
        
        #生成 立卷日期 下拉选择框
        self.modify_lijuanriqi_archives_pop = UtilData.DateControl(self, -1, pos = (30,30),size=ARCHIVES_SIZE)
        self.modify_lijuanriqi_archives_pop.SetValue(temp_list_frame[5])
        
        #生成 位置 下拉选择框
        self.modify_weizhi_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[6],size=ARCHIVES_SIZE,choices=UtilData.WEIZHI_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_weizhi_archives_comboBox.SetPopupMaxHeight(self.modify_weizhi_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件
        self.modify_weizhi_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceEvent)        
        
        #生成 密级 下拉选择框
        self.modify_miji_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[7],size=ARCHIVES_SIZE,choices=UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_miji_archives_comboBox.SetPopupMaxHeight(self.modify_miji_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件
        self.modify_miji_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceEvent)    
        
        #生成 责任人 输入框
        self.modify_zerenren_archives_ctrl = wx.TextCtrl(parent=self,value=temp_list_frame[8],size=ARCHIVES_SIZE)        
        
        #生成  区号  下拉选择框
        self.modify_quhao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[9],size=ARCHIVES_SIZE,choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_quhao_archives_comboBox.SetPopupMaxHeight(self.modify_quhao_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件
        self.modify_quhao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFrame)

        #生成  柜号  下拉选择框
        self.modify_guihao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[10],size=ARCHIVES_SIZE,choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_guihao_archives_comboBox.SetPopupMaxHeight(self.modify_guihao_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件
        self.modify_guihao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFrame)
        
        #生成  盒号  下拉选择框
        self.modify_hehao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[11],size=ARCHIVES_SIZE,choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_hehao_archives_comboBox.SetPopupMaxHeight(self.modify_hehao_archives_comboBox.GetCharHeight()*10)
        #为盒号选择框绑定事件(自动判断该盒号下的卷号是否唯一)
        self.modify_hehao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFrame)
        
        #取得初始化时区号、卷号、盒号的值,设置卷号控件为None方便初始化数据库查询
        self.modify_initQuhao_archives_comboBox_value = self.modify_quhao_archives_comboBox.GetValue()
        self.modify_initGuihao_archives_comboBox_value = self.modify_guihao_archives_comboBox.GetValue()
        self.modify_initHehao_archives_comboBox_value = self.modify_hehao_archives_comboBox.GetValue()
        self.modify_juanhao_archives_comboBox = None
        #生成卷号初始化列表
        self.temp_ccfjhf = self.createChoicesForJuanHaoFrame(event='')
        self.temp_list =temp_list_frame[12].split() + self.temp_ccfjhf
        #生成  卷号  下拉选择框
        self.modify_juanhao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[12],size=ARCHIVES_SIZE,choices=self.temp_list,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_juanhao_archives_comboBox.SetPopupMaxHeight(self.modify_juanhao_archives_comboBox.GetCharHeight()*10)

        ##取得初始化时单位的值,设置案卷号控件为None方便初始化数据库查询
        #self.modify_initDanwei_archives_comboBox_value = self.modify_danwei_archives_comboBox.GetValue()
        #self.modify_anjuanhao_archives_comboBox = None
        ##生成 案卷号 初始化列表
        #self.temp_accfjhf = self.createChoicesForAnJuanHaoFrame(event='')
        #self.temp_list_anjuanhao =temp_list_frame[4].split() + self.temp_accfjhf
        ##生成  案卷号  下拉选择框
        #self.modify_anjuanhao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[4],size=ARCHIVES_SIZE,choices=self.temp_list_anjuanhao,style=wx.CB_READONLY)
        ##设置下拉列表一次容纳10个元素
        #self.modify_anjuanhao_archives_comboBox.SetPopupMaxHeight(self.modify_anjuanhao_archives_comboBox.GetCharHeight()*10)       
       
        #生成 互见号 输入框
        self.modify_hujianhao_archives_ctrl = wx.TextCtrl(parent=self,value=temp_list_frame[13],size=ARCHIVES_SIZE)
        
        #生成 科目 下拉选择框
        self.modify_kemu_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_frame[14],size=ARCHIVES_SIZE,choices=UtilData.KEMU_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.modify_kemu_archives_comboBox.SetPopupMaxHeight(self.modify_kemu_archives_comboBox.GetCharHeight()*10)
        
        #生成 备注 输入框
        self.modify_beizhu_archives_ctrl = wx.TextCtrl(parent=self,value=temp_list_frame[16],size=ARCHIVES_SIZE)     
        
        #生成 录入人 输入框
        self.modify_inputter_archives_ctrl = wx.TextCtrl(parent=self,value=temp_list_frame[17],size=ARCHIVES_SIZE)        
        
        #生成 修改档案 按钮并进行绑定
        self.modify_newArchive_button = wx.Button(parent=self,label=u'修改档案')
        self.modify_newArchive_button.Bind(wx.EVT_BUTTON, self.modifyOldArchive)

        #将各选择与输入控件加入布局管理器
        modify_archives_gridBagSizer.Add(self.modify_menlei_archives_comboBox,pos=(0,1),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        modify_archives_gridBagSizer.Add(self.modify_guidang_archives_comboBox,pos=(1,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_anjuantiming_archives_ctrl,pos=(2,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_qixian_archives_comboBox,pos=(3,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_danwei_archives_comboBox,pos=(4,1),flag=wx.ALIGN_LEFT)
        #此处放置 立卷日期 pos=(5,1)
        modify_archives_gridBagSizer.Add(self.modify_lijuanriqi_archives_pop,pos=(5,1),flag=wx.ALIGN_LEFT) 
        modify_archives_gridBagSizer.Add(self.modify_weizhi_archives_comboBox,pos=(6,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_miji_archives_comboBox,pos=(7,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_zerenren_archives_ctrl,pos=(8,1),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_quhao_archives_comboBox,pos=(0,3),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        modify_archives_gridBagSizer.Add(self.modify_guihao_archives_comboBox,pos=(1,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_hehao_archives_comboBox,pos=(2,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_juanhao_archives_comboBox,pos=(3,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_hujianhao_archives_ctrl,pos=(4,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_kemu_archives_comboBox,pos=(5,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_beizhu_archives_ctrl,pos=(6,3),span=(2,1),flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        modify_archives_gridBagSizer.Add(self.modify_inputter_archives_ctrl,pos=(8,3),flag=wx.ALIGN_LEFT)
        modify_archives_gridBagSizer.Add(self.modify_newArchive_button,pos=(9,0),span=(1,4),flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_TOP,border=20)   
        #各列尽可能扩展
        modify_archives_gridBagSizer.AddGrowableCol(0)
        modify_archives_gridBagSizer.AddGrowableCol(3)
        self.SetSizer(modify_archives_gridBagSizer)
        self.Show()
        self.Center()
    #根据修改frame的门类的选择动态提供归档类型
    def choiceGuiDangFrame(self,event):
        TEMP = self.modify_menlei_archives_comboBox.GetValue().split(u"_")[:1]
        TEMP = "".join(TEMP)
        #print TEMP,type(TEMP)
        GUIDANG_MENLEI_List = []
        if TEMP in u'MENLEI_W_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_W_List
        if TEMP in u'MENLEI_J_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_J_List   
        if TEMP in u'MENLEI_K_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_K_List
        if TEMP in u'MENLEI_Y_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_Y_List 
        if TEMP in u'MENLEI_S_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_S_List
        if TEMP in u'MENLEI_D_List'.split(u"_"):
            GUIDANG_MENLEI_List = UtilData.MENLEI_D_List 
        self.modify_guidang_archives_comboBox.Set(GUIDANG_MENLEI_List)
        #默认第0个元素选中状态
        self.modify_guidang_archives_comboBox.SetSelection(0)
        
    def createChoicesForJuanHaoFrame(self,event):
        #查询给定区号、柜号、盒号下的卷号列表
        temp_juanhao_list = DoCURD.query_juanhao_values(CONN,paramDict={"table_name":"archives",
                                                                         "quhao":self.modify_quhao_archives_comboBox.GetValue() ,
                                                                         "guihao":self.modify_guihao_archives_comboBox.GetValue() ,
                                                                         "hehao":self.modify_hehao_archives_comboBox.GetValue()})
        temp_juanhaoExist_list = []
        for rows in temp_juanhao_list:
            temp_juanhaoExist_list.append(rows[0])
        #如果元素在全列表中而不在已存在列表中则将该元素返回
        temp_choices_list = [i for i in UtilData.JUANHAO_List if i not in temp_juanhaoExist_list]
        if self.modify_juanhao_archives_comboBox is not None :
            #如果原始盒号值等于现在盒号值,则将原始盒号值下的卷号列表赋值到现在卷号列表下
            if self.modify_initHehao_archives_comboBox_value == self.modify_hehao_archives_comboBox.GetValue() and\
               self.modify_initQuhao_archives_comboBox_value == self.modify_quhao_archives_comboBox.GetValue() and\
               self.modify_initGuihao_archives_comboBox_value == self.modify_guidang_archives_comboBox.GetValue() :
                temp_choices_list = self.temp_list
            #设置卷号控件中数据源为新生成的temp_choices_list
            self.modify_juanhao_archives_comboBox.Set(temp_choices_list)
            #设置第0个元素为默认显示
            self.modify_juanhao_archives_comboBox.SetSelection(0)
        else:
            return temp_choices_list
    
    #def createChoicesForAnJuanHaoFrame(self,event):
        ##查询给定区号、柜号、盒号下的卷号列表
        #temp_anjuanhao_list = DoCURD.query_anjuanhao_values(CONN,paramDict={"table_name":"archives",
                                                                         #"menlei":self.modify_menlei_archives_comboBox.GetValue() ,
                                                                         #"guidang":self.modify_guidang_archives_comboBox.GetValue() ,
                                                                         #"danwei":self.modify_danwei_archives_comboBox.GetValue()})
        #temp_anjuanhaoExist_list = []
        #for rows in temp_anjuanhao_list:
            #temp_anjuanhaoExist_list.append(rows[0])  
        ##如果元素在全列表中而不在已存在列表中则将该元素返回
        #temp_anjuan_choices_list = [i for i in UtilData.JUANHAO_List if i not in temp_anjuanhaoExist_list]
        #if self.modify_anjuanhao_archives_comboBox is not None :
            ##如果原始单位值等于现在单位值,则将原始单位值下的案卷号列表赋值到现在案卷号列表下
            #if self.modify_initDanwei_archives_comboBox_value == self.modify_danwei_archives_comboBox.GetValue():
                #temp_anjuan_choices_list = self.temp_list_anjuanhao              
            ##设置卷号控件中数据源为新生成的temp_choices_list
            #self.modify_anjuanhao_archives_comboBox.Set(temp_anjuan_choices_list)
            ##设置第0个元素为默认显示
            #self.modify_anjuanhao_archives_comboBox.SetSelection(0)
        #else:
            #return temp_anjuan_choices_list
        
    def choiceEvent(self,event):
        pass
    
    def modifyOldArchive(self,event):
        #判断案卷题名、立卷日期、责任人、输入人是否为空
        if self.modify_anjuantiming_archives_ctrl.Value.strip() == "" or self.modify_lijuanriqi_archives_pop.GetValue().strip() == "" or self.modify_zerenren_archives_ctrl.Value.strip() == "" or self.modify_inputter_archives_ctrl.Value.strip() == "" :
            errorMessage = wx.MessageDialog(parent=None,message=u"案卷题名、立卷日期、责任人、录入人不允许为空！",caption=u"修改档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return
        #为了判断日期是否准确临时生成一个temp_cal变量
        temp_cal = self.modify_lijuanriqi_archives_pop.GetValue().strip()
        if temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日" :
            errorMessage = wx.MessageDialog(parent=None,message=u"立卷日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"修改档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return    
        #判断年份是否已经归档，如果年份已经归档则不允许新增
        year_already_RowsList = DoCURD.query_already_year(CONN)
        year_already_List = []
        for row in year_already_RowsList :
            year_already_List.append(row[0])
        if temp_cal[:4] in year_already_List :
            errorMessage = wx.MessageDialog(parent=None,message=unicode(temp_cal[:4])+u"年档案已经归档,不允许修改本档案年份为已归档年份!",caption=u"新增档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()             
            return        
        paramDict={}
        #将修改面板中各控件的值加入参数字典
        paramDict['table_name']=u'archives'
        paramDict['archiveID']=temp_list_frame[0]
        paramDict['menlei']=self.modify_menlei_archives_comboBox.GetValue()
        paramDict['guidang']=self.modify_guidang_archives_comboBox.GetValue()
        paramDict['qixian']=self.modify_qixian_archives_comboBox.GetValue()
        paramDict['anjuantiming']=self.modify_anjuantiming_archives_ctrl.Value.strip()
        paramDict['danwei']=self.modify_danwei_archives_comboBox.GetValue()
        paramDict['lijuanriqi']=self.modify_lijuanriqi_archives_pop.GetValue()
        paramDict['weizhi']=self.modify_weizhi_archives_comboBox.GetValue()
        paramDict['miji']=self.modify_miji_archives_comboBox.GetValue()
        paramDict['zerenren']=self.modify_zerenren_archives_ctrl.Value.strip()
        paramDict['quhao']=self.modify_quhao_archives_comboBox.GetValue()
        paramDict['guihao']=self.modify_guihao_archives_comboBox.GetValue()
        paramDict['hehao']=self.modify_hehao_archives_comboBox.GetValue()
        paramDict['juanhao']=self.modify_juanhao_archives_comboBox.GetValue()
        paramDict['hujianhao']=self.modify_hujianhao_archives_ctrl.Value.strip()
        paramDict['kemu']=self.modify_kemu_archives_comboBox.GetValue()
        paramDict['beizhu']=self.modify_beizhu_archives_ctrl.Value.strip()
        paramDict['inputter']=self.modify_inputter_archives_ctrl.Value.strip()
        #将连接和参数字典传入更新语句,修改档案信息
        updateState = DoCURD.update_some_values_from_archivetable(CONN,paramDict)
        if updateState == u'ok' :
            print updateState
            #提示档案修改成功
            inforMessage = wx.MessageDialog(parent=None,message=u"所选档案修改成功！",caption=u"提示...",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()   
            #调用父窗口(TestPanel)的父窗口(myFrame)下的管理页查询方法
            self.Parent.Parent.manageArchivesBySomeConditions(self)
            self.Parent.Parent.selectArchivesBySomeConditions(self)
        else :
            print updateState
        self.Destroy()
#---------------------------------------------------------------------------

class TestPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        dlg = TestDialog(self, -1, u"档案修改", size=(450,370),
                         style=wx.DEFAULT_DIALOG_STYLE)
        
        dlg.CenterOnScreen()
        val = dlg.ShowModal()    
        dlg.Destroy()
        