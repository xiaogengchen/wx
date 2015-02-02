# -*- coding:utf-8 -*-

import os
import pickle
import shutil
import wx
import wx.combo
import wx.grid
import wx.lib.filebrowsebutton
import UtilData
import DoCURD
import DialogDemo
import DialogDemo_File
import datetime
import outputPrint
import Regist
import ShowFiles
from win32com.shell import shell, shellcon


CONNECTION = DoCURD.connect_db('./tbdata.db')    
class MyFrame(wx.Frame):
    #框架界面初始化,用户各元素外观布局^(wx.MAXIMIZE_BOX|wx.RESIZE_BORDER)
    def __init__(self):
        wx.Frame.__init__(self,parent=None,title=u"天保档案管理系统   v1.0",size=(1024,768),style=wx.DEFAULT_FRAME_STYLE^(wx.MAXIMIZE_BOX|wx.RESIZE_BORDER))
        #初始化用户名
        self.USERNAME = ''
        #初始化数据库
        self.CONN = CONNECTION
        #frame居中
        self.Center(wx.BOTH)
        #生成panel
        self.panel = wx.Panel(parent=self)
        #常量----logo、title、welcome路径和尺寸
        LOGO = {"path":"logo.bmp","size":(120,120)}
        TITLE = {"path":"title.bmp","size":(899,120)}
        WELCOME = {"path":"welcome.bmp","size":(1018,570)}
        #主布局管理器纵向排列
        self.mainBoxSizer = wx.BoxSizer(wx.VERTICAL)
        #    11、创建上层布局管理器横向排列
        topBoxSizer = wx.BoxSizer(wx.HORIZONTAL)
        #    12、静态图片logo,父容器为主面板
        logo = wx.StaticBitmap(parent=self.panel,bitmap=wx.Bitmap(LOGO.get("path"),wx.BITMAP_TYPE_BMP),size=LOGO.get("size"))
        #    13、向上层布局管理器添加logo
        topBoxSizer.Add(logo,flag=wx.LEFT)
        #    14、静态图片title，父容器为主面板
        title = wx.StaticBitmap(parent=self.panel,bitmap=wx.Bitmap(TITLE.get("path"),wx.BITMAP_TYPE_BMP),size=TITLE.get("size"))
        #    15、向上层面板布局管理器添加title
        topBoxSizer.Add(title,flag=wx.EXPAND)
        #    16、向主布局管理器中添加上层布局管理器
        self.mainBoxSizer.Add(topBoxSizer)    
        #主标签框架
        self.main_Notebook = wx.Notebook(parent=self.panel,size=(1021,612))
        #子框架尺寸和子框架下表格的尺寸
        SUB_FRAME_SIZE = (1015,565)     
        GRID_SIZE = (1015,505)
        #**********欢迎页开始**********
        self.welcome_panel = wx.Panel(parent=self.main_Notebook)
        
        #背景图片
        welcomeImage = wx.StaticBitmap(parent=self.welcome_panel,bitmap=wx.Bitmap(WELCOME.get("path"),wx.BITMAP_TYPE_BMP),size=WELCOME.get("size"))

        #欢迎页标题及字体样式  label=u"档 案 管 理 系 统"
        welcome_title = wx.StaticText(parent=self.welcome_panel,label=u"",style=wx.ALIGN_CENTER)
        font_title = wx.Font(pointSize=35,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        welcome_title.SetFont(font_title)
        #用户名
        usernameLabel_welcome = wx.StaticText(parent=welcomeImage,label=u"用户名: ",style=wx.ALIGN_CENTER)
        font_usernameAndPassword = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.NORMAL)
        usernameLabel_welcome.SetFont(font_usernameAndPassword)
        self.usernameText_welcome = wx.TextCtrl(parent=welcomeImage,size=(120,25))
        #密码
        passwordLabel_welcome = wx.StaticText(parent=welcomeImage,label=u"密    码: ",style=wx.ALIGN_CENTER)
        passwordLabel_welcome.SetFont(font_usernameAndPassword)
        self.passwordText_welcome = wx.TextCtrl(parent=welcomeImage,size=(120,25),style=wx.TE_PASSWORD)
        #登陆按钮
        #loginButton = wx.Button(parent=self.welcome_panel,label=u"登陆")
        loginButton = wx.Button(parent=welcomeImage,label=u"登陆")
        loginButton.Bind(wx.EVT_BUTTON, self.checkLogin)
        #生成欢迎页中垂直方向的布局管理器
        welcome_vboxsizer = wx.BoxSizer(wx.VERTICAL)
        #横向用户名管理器
        username_welcome_hboxsizer = wx.BoxSizer(wx.HORIZONTAL)
        username_welcome_hboxsizer.Add(usernameLabel_welcome,flag=wx.ALIGN_CENTER)
        username_welcome_hboxsizer.Add(self.usernameText_welcome)
        #横向密码管理器
        password_welcome_hboxsizer = wx.BoxSizer(wx.HORIZONTAL)
        password_welcome_hboxsizer.Add(passwordLabel_welcome,flag=wx.ALIGN_CENTER)
        password_welcome_hboxsizer.Add(self.passwordText_welcome)
        #向纵向布局管理器中添加欢迎页的标题、用户名管理器、密码管理器、登陆按钮
        welcome_vboxsizer.Add(welcome_title,flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP,border=50)
        welcome_vboxsizer.Add(username_welcome_hboxsizer,flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP,border=70)
        welcome_vboxsizer.Add(password_welcome_hboxsizer,flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP,border=20)
        welcome_vboxsizer.Add(loginButton,flag=wx.ALIGN_CENTER_HORIZONTAL|wx.TOP,border=20)
        #在欢迎面板中设置垂直方向的布局管理器
        self.welcome_panel.SetSizer(welcome_vboxsizer)
        #向主框架中加入欢迎页
        self.main_Notebook.AddPage(self.welcome_panel,u"欢迎")
        #**********欢迎页结束**********      
        
        #**********档案管理页开始**********
        self.archivesManage_panel = wx.Panel(parent=self.main_Notebook)
        #档案管理框架
        self.archivesManage_Notebook = wx.Notebook(parent=self.archivesManage_panel,size=SUB_FRAME_SIZE)
        #档案管理框架中新增页
        self.add_archives_panel = wx.Panel(parent=self.archivesManage_Notebook)
        #生成GridBagSizer布局管理器
        add_archives_gridBagSizer = wx.GridBagSizer(vgap=20,hgap=10)
        #生成文本标签
        add_archives_menleiText = wx.StaticText(parent=self.add_archives_panel,label=u"门        类:",style=wx.ALIGN_CENTER_VERTICAL)
        #add_archives_jibieText = wx.StaticText(parent=self.add_archives_panel,label=u"级        别:",style=wx.ALIGN_CENTER_HORIZONTAL)
        add_archives_guidangText = wx.StaticText(parent=self.add_archives_panel,label=u"分        类:",style=wx.ALIGN_CENTER_VERTICAL)
        add_archives_qixianText = wx.StaticText(parent=self.add_archives_panel,label=u"期        限:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_anjuantimingText = wx.StaticText(parent=self.add_archives_panel,label=u"案卷题名:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_danweiText = wx.StaticText(parent=self.add_archives_panel,label=u"单        位:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_lijuanriqiText = wx.StaticText(parent=self.add_archives_panel,label=u"立卷日期:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_weizhiText = wx.StaticText(parent=self.add_archives_panel,label=u"位        置:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_mijiText = wx.StaticText(parent=self.add_archives_panel,label=u"密        级:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_zerenrenText = wx.StaticText(parent=self.add_archives_panel,label=u"责  任  人:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_quhaoText = wx.StaticText(parent=self.add_archives_panel,label=u"区        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_guihaoText = wx.StaticText(parent=self.add_archives_panel,label=u"柜        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_hehaoText = wx.StaticText(parent=self.add_archives_panel,label=u"盒        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_juanhaoText = wx.StaticText(parent=self.add_archives_panel,label=u"卷        号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_hujianhaoText = wx.StaticText(parent=self.add_archives_panel,label=u"互  见  号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_kemuText = wx.StaticText(parent=self.add_archives_panel,label=u"科        目:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_beizhuText = wx.StaticText(parent=self.add_archives_panel,label=u"备        注:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_inputterText = wx.StaticText(parent=self.add_archives_panel,label=u"录  入  人:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        add_archives_font = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        add_archives_menleiText.SetFont(add_archives_font)
        add_archives_guidangText.SetFont(add_archives_font)
        add_archives_qixianText.SetFont(add_archives_font)
        add_archives_anjuantimingText.SetFont(add_archives_font)
        add_archives_danweiText.SetFont(add_archives_font)
        add_archives_lijuanriqiText.SetFont(add_archives_font)
        add_archives_weizhiText.SetFont(add_archives_font)
        add_archives_mijiText.SetFont(add_archives_font)
        add_archives_zerenrenText.SetFont(add_archives_font)
        add_archives_quhaoText.SetFont(add_archives_font)
        add_archives_guihaoText.SetFont(add_archives_font)
        add_archives_hehaoText.SetFont(add_archives_font)
        add_archives_juanhaoText.SetFont(add_archives_font)
        add_archives_hujianhaoText.SetFont(add_archives_font)
        add_archives_kemuText.SetFont(add_archives_font)
        add_archives_beizhuText.SetFont(add_archives_font)
        add_archives_inputterText.SetFont(add_archives_font)
        #将文本标签纳入add_archives_gridBagSizer布局管理器中
        add_archives_gridBagSizer.Add(add_archives_menleiText,pos=(0,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        #add_archives_gridBagSizer.Add(add_archives_jibieText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_guidangText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_anjuantimingText,pos=(2,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_qixianText,pos=(3,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_danweiText,pos=(4,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_lijuanriqiText,pos=(5,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_weizhiText,pos=(6,0),flag=wx.wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_mijiText,pos=(7,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_zerenrenText,pos=(8,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_quhaoText,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT|wx.TOP,border=15)
        add_archives_gridBagSizer.Add(add_archives_guihaoText,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_hehaoText,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_juanhaoText,pos=(3,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_hujianhaoText,pos=(4,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_kemuText,pos=(5,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_beizhuText,pos=(6,2),span=(2,1),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_archives_gridBagSizer.Add(add_archives_inputterText,pos=(8,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        ARCHIVES_SIZE = (180,30)
        #生成  门类  下拉选择框
        #根据不同门类选项存放不同归档信息的列表
        self.menlei_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.MENLEI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.menlei_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.menlei_archives_comboBox.SetPopupMaxHeight(self.menlei_archives_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件,目的是根据选择不同的门类提供不同的归档列表值
        self.menlei_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceGuiDang)        
        #生成  归档  下拉选择框
        self.guidang_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=(465,30),choices=UtilData.MENLEI_W_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guidang_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guidang_archives_comboBox.SetPopupMaxHeight(self.guidang_archives_comboBox.GetCharHeight()*10)
        #生成  期限  下拉选择框
        self.qixian_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.QIXIAN_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.qixian_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.qixian_archives_comboBox.SetPopupMaxHeight(self.qixian_archives_comboBox.GetCharHeight()*10)
        #生成 案卷题名 输入框
        self.anjuantiming_archives_ctrl = wx.TextCtrl(parent=self.add_archives_panel,size=(465,30))
        #生成  单位  下拉选择框
        self.danwei_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.DANWEI_List,style=wx.CB_READONLY)
        #设置下拉列表一次容纳10个元素
        self.danwei_archives_comboBox.SetPopupMaxHeight(self.danwei_archives_comboBox.GetCharHeight()*10)       
        #生成 立卷日期 下拉选择框
        self.lijuanriqi_archives_pop = UtilData.DateControl(self.add_archives_panel, -1, pos = (30,30),size=ARCHIVES_SIZE)
        today = datetime.datetime.now().strftime('%Y%m%d')
        #设置立卷日期默认值
        self.lijuanriqi_archives_pop.SetValue(unicode(today[:4]+"年"+today[4:6]+"月"+today[6:]+"日"))
        #生成 位置 下拉选择框
        self.weizhi_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.WEIZHI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.weizhi_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.weizhi_archives_comboBox.SetPopupMaxHeight(self.weizhi_archives_comboBox.GetCharHeight()*10)
        #生成 密级 下拉选择框
        self.miji_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.miji_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.miji_archives_comboBox.SetPopupMaxHeight(self.miji_archives_comboBox.GetCharHeight()*10)  
        #生成 责任人 输入框
        self.zerenren_archives_ctrl = wx.TextCtrl(parent=self.add_archives_panel,size=ARCHIVES_SIZE)        
        #生成  区号  下拉选择框
        self.quhao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.quhao_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.quhao_archives_comboBox.SetPopupMaxHeight(self.quhao_archives_comboBox.GetCharHeight()*10)
        self.quhao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHao)
        #生成  柜号  下拉选择框
        self.guihao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guihao_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guihao_archives_comboBox.SetPopupMaxHeight(self.guihao_archives_comboBox.GetCharHeight()*10)
        self.guihao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHao)
        #生成  盒号  下拉选择框
        self.hehao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.hehao_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.hehao_archives_comboBox.SetPopupMaxHeight(self.hehao_archives_comboBox.GetCharHeight()*10)
        #为盒号选择框绑定事件(自动判断该盒号下的卷号是否唯一)
        self.hehao_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHao)
        #生成  卷号  下拉选择框
        self.juanhao_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.JUANHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.juanhao_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.juanhao_archives_comboBox.SetPopupMaxHeight(self.juanhao_archives_comboBox.GetCharHeight()*10)
        #生成 互见号 输入框
        self.hujianhao_archives_ctrl = wx.TextCtrl(parent=self.add_archives_panel,size=ARCHIVES_SIZE)
        #生成 科目 下拉选择框
        self.kemu_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_archives_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.KEMU_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.kemu_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.kemu_archives_comboBox.SetPopupMaxHeight(self.kemu_archives_comboBox.GetCharHeight()*10)
        #生成 备注 输入框
        self.beizhu_archives_ctrl = wx.TextCtrl(parent=self.add_archives_panel,size=ARCHIVES_SIZE)     
        #生成 录入人 输入框
        self.inputter_archives_ctrl = wx.TextCtrl(parent=self.add_archives_panel,size=ARCHIVES_SIZE) 
        self.newArchive_button = wx.Button(parent=self.add_archives_panel,label=u'新增档案')
        self.newArchive_button.Bind(wx.EVT_BUTTON, self.addNewArchive)
        #将各选择与输入控件加入布局管理器
        add_archives_gridBagSizer.Add(self.menlei_archives_comboBox,pos=(0,1),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        add_archives_gridBagSizer.Add(self.guidang_archives_comboBox,pos=(1,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.anjuantiming_archives_ctrl,pos=(2,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.qixian_archives_comboBox,pos=(3,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.danwei_archives_comboBox,pos=(4,1),flag=wx.ALIGN_LEFT)
        #此处放置 立卷日期 pos=(5,1)
        add_archives_gridBagSizer.Add(self.lijuanriqi_archives_pop,pos=(5,1),flag=wx.ALIGN_LEFT) 
        add_archives_gridBagSizer.Add(self.weizhi_archives_comboBox,pos=(6,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.miji_archives_comboBox,pos=(7,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.zerenren_archives_ctrl,pos=(8,1),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.quhao_archives_comboBox,pos=(0,3),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        add_archives_gridBagSizer.Add(self.guihao_archives_comboBox,pos=(1,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.hehao_archives_comboBox,pos=(2,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.juanhao_archives_comboBox,pos=(3,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.hujianhao_archives_ctrl,pos=(4,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.kemu_archives_comboBox,pos=(5,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.beizhu_archives_ctrl,pos=(6,3),span=(2,1),flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        add_archives_gridBagSizer.Add(self.inputter_archives_ctrl,pos=(8,3),flag=wx.ALIGN_LEFT)
        add_archives_gridBagSizer.Add(self.newArchive_button,pos=(9,0),span=(1,4),flag=wx.ALIGN_CENTER)   
        #各列尽可能扩展
        add_archives_gridBagSizer.AddGrowableCol(0)
        add_archives_gridBagSizer.AddGrowableCol(3)
        #设置add_archives_gridBagSizer为add_archives_panel的布局管理器
        self.add_archives_panel.SetSizer(add_archives_gridBagSizer)
        #档案管理框架中查询页
        self.select_archives_panel = wx.Panel(parent=self.archivesManage_Notebook)
        #档案查询页中的总垂直布局管理器和顶部横向布局管理器
        select_archives_vbox = wx.BoxSizer(wx.VERTICAL)
        select_archives_hbox = wx.BoxSizer(wx.HORIZONTAL)
        select_archives_hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        #区号
        select_archives_quhaoText = wx.StaticText(parent=self.select_archives_panel,label=u"区号:")
        #生成  区号  下拉选择框
        self.quhao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,choices=[""]+UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.quhao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.quhao_select_archives_comboBox.SetPopupMaxHeight(self.quhao_select_archives_comboBox.GetCharHeight()*10) 
        #柜号
        select_archives_guihaoText = wx.StaticText(parent=self.select_archives_panel,label=u"柜号:")
        self.guihao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,choices=[""]+UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guihao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guihao_select_archives_comboBox.SetPopupMaxHeight(self.guihao_select_archives_comboBox.GetCharHeight()*10)       
        #盒号
        select_archives_hehaoText = wx.StaticText(parent=self.select_archives_panel,label=u"盒号:")
        self.hehao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=(43,25),choices=[""]+UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.hehao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.hehao_select_archives_comboBox.SetPopupMaxHeight(self.hehao_select_archives_comboBox.GetCharHeight()*10)
        #卷号
        select_archives_juanhaoText = wx.StaticText(parent=self.select_archives_panel,label=u"卷号:")
        self.juanhao_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=(43,25),choices=[""]+UtilData.JUANHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.juanhao_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.juanhao_select_archives_comboBox.SetPopupMaxHeight(self.juanhao_select_archives_comboBox.GetCharHeight()*10)       
        #***********************************高级检索begine********************************************
        SENIOR_SELECT_SIZE = (43,25)
        #门类
        select_archives_menleiText = wx.StaticText(parent=self.select_archives_panel,label=u"门类:")
        self.menleiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MENLEI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.menleiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.menleiText_select_archives_comboBox.SetPopupMaxHeight(self.menleiText_select_archives_comboBox.GetCharHeight()*10)                
        #绑定门类
        self.menleiText_select_archives_comboBox.Bind(wx.EVT_COMBOBOX, self.choiceGuiDangForSeniorSelect)
        #归档(分类)
        select_archives_guidangText = wx.StaticText(parent=self.select_archives_panel,label=u"分类:")
        self.guidangText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""],style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guidangText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guidangText_select_archives_comboBox.SetPopupMaxHeight(self.guidangText_select_archives_comboBox.GetCharHeight()*10)                
        #案卷题名
        select_archives_anjuantimingText = wx.StaticText(parent=self.select_archives_panel,label=u"案卷题名:")
        self.anjuantiming_select_archives_Ctrl = wx.TextCtrl(parent=self.select_archives_panel,size=(134,25))
        #期限
        select_archives_qixianText = wx.StaticText(parent=self.select_archives_panel,label=u"期限:")
        self.qixianText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.QIXIAN_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.qixianText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.qixianText_select_archives_comboBox.SetPopupMaxHeight(self.qixianText_select_archives_comboBox.GetCharHeight()*10)        
        #立卷单位
        select_archives_lijuandanweiText = wx.StaticText(parent=self.select_archives_panel,label=u"单位:")
        self.lijuandanweiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.DANWEI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.lijuandanweiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.lijuandanweiText_select_archives_comboBox.SetPopupMaxHeight(self.lijuandanweiText_select_archives_comboBox.GetCharHeight()*10)         
        #立卷日期
        select_archives_lijuanriqiText = wx.StaticText(parent=self.select_archives_panel,label=u"立卷日期:")
        self.lijuanriqiText_select_archives_pop = UtilData.DateControl(self.select_archives_panel, -1)       
        #录入人
        select_archives_inputterText = wx.StaticText(parent=self.select_archives_panel,label=u"录入人:")
        self.inputterText_select_archives_Ctrl = wx.TextCtrl(parent=self.select_archives_panel)  
        #存放位置
        select_archives_weizhiText = wx.StaticText(parent=self.select_archives_panel,label=u"存放位置:")
        self.weizhiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.WEIZHI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.weizhiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.weizhiText_select_archives_comboBox.SetPopupMaxHeight(self.weizhiText_select_archives_comboBox.GetCharHeight()*10)                 
        #互见号
        select_archives_hujianhaoText = wx.StaticText(parent=self.select_archives_panel,label=u"互见号:")
        self.hujianhaoText_select_archives_Ctrl = wx.TextCtrl(parent=self.select_archives_panel)        
        #档案科目
        select_archives_kemuText = wx.StaticText(parent=self.select_archives_panel,label=u"档案科目:")
        self.kemuText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.KEMU_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.kemuText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.kemuText_select_archives_comboBox.SetPopupMaxHeight(self.kemuText_select_archives_comboBox.GetCharHeight()*10)        
        #密级
        select_archives_mijiText = wx.StaticText(parent=self.select_archives_panel,label=u"密级:")
        self.mijiText_select_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_archives_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.mijiText_select_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.mijiText_select_archives_comboBox.SetPopupMaxHeight(self.mijiText_select_archives_comboBox.GetCharHeight()*10)        
        ##档案ID
        #select_archives_archiveIDText = wx.StaticText(parent=self.select_archives_panel,label=u"档案ID:")
        #self.archiveIDText_select_archives_Ctrl = wx.TextCtrl(parent=self.select_archives_panel,size=(122,25))        
        #查询按钮
        select_archives_button = wx.Button(parent=self.select_archives_panel,label=u"查询",size=(75,28))
        select_archives_showFiles_button = wx.Button(parent=self.select_archives_panel,label=u"查看",size=(75,28))
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
        self.select_grid_archives = wx.grid.Grid(parent=self.select_archives_panel,size=GRID_SIZE)
        self.select_data_table = myArchiveTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_select_table_rows = self.select_data_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.select_grid_archives.SetTable(self.select_data_table,takeOwnership=True)
        self.select_grid_archives.AutoSize()
        #向纵向布局管理器中添加数据表格
        select_archives_vbox.Add(self.select_grid_archives,flag=wx.ALIGN_BOTTOM)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.select_archives_panel.SetSizer(select_archives_vbox)
        #档案管理框架中管理页
        self.manage_archives_panel = wx.Panel(parent=self.archivesManage_Notebook)
        #档案查询页中的总垂直布局管理器和顶部横向布局管理器
        manage_archives_vbox = wx.BoxSizer(wx.VERTICAL)
        manage_archives_hbox = wx.BoxSizer(wx.HORIZONTAL)
        #管理 区号
        manage_archives_quhaoText = wx.StaticText(parent=self.manage_archives_panel,label=u"区号:")
        #生成  区号  下拉选择框
        self.quhao_manage_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.manage_archives_panel,id=-1,choices=[""]+UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.quhao_manage_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.quhao_manage_archives_comboBox.SetPopupMaxHeight(self.quhao_manage_archives_comboBox.GetCharHeight()*10)
        #管理 柜号
        manage_archives_guihaoText = wx.StaticText(parent=self.manage_archives_panel,label=u"柜号:")
        self.guihao_manage_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.manage_archives_panel,id=-1,choices=[""]+UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guihao_manage_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guihao_manage_archives_comboBox.SetPopupMaxHeight(self.guihao_manage_archives_comboBox.GetCharHeight()*10)        
        #管理 盒号
        manage_archives_hehaoText = wx.StaticText(parent=self.manage_archives_panel,label=u"盒号:")
        self.hehao_manage_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.manage_archives_panel,id=-1,choices=[""]+UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.hehao_manage_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.hehao_manage_archives_comboBox.SetPopupMaxHeight(self.hehao_manage_archives_comboBox.GetCharHeight()*10)
        #管理 卷号
        manage_archives_juanhaoText = wx.StaticText(parent=self.manage_archives_panel,label=u"卷号:")
        self.juanhao_manage_archives_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.manage_archives_panel,id=-1,choices=[""]+UtilData.JUANHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.juanhao_manage_archives_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.juanhao_manage_archives_comboBox.SetPopupMaxHeight(self.juanhao_manage_archives_comboBox.GetCharHeight()*10)
        #管理 案卷题名
        manage_archives_anjuantimingText = wx.StaticText(parent=self.manage_archives_panel,label=u"案卷题名:")
        self.anjuantiming_manage_archives_Ctrl = wx.TextCtrl(parent=self.manage_archives_panel)
        #管理 查询按钮
        manage_archives_button = wx.Button(parent=self.manage_archives_panel,label=u"查询")
        #对查询按钮进行事件绑定
        manage_archives_button.Bind(wx.EVT_BUTTON, self.manageArchivesBySomeConditions)
        #管理 修改按钮
        manage_archives_modifyButton = wx.Button(parent=self.manage_archives_panel,label=u'修改')
        #对修改按钮进行事件绑定
        manage_archives_modifyButton.Bind(wx.EVT_BUTTON, self.manageArchives_modify)
        #管理 删除按钮
        manage_archives_deleteButton = wx.Button(parent=self.manage_archives_panel,label=u'删除')
        #对删除按钮进行事件绑定
        manage_archives_deleteButton.Bind(wx.EVT_BUTTON, self.manageArchives_delete)        
        #向顶部横向布局管理器中添加元素
        manage_archives_hbox.Add(manage_archives_quhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(self.quhao_manage_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_archives_hbox.Add(manage_archives_guihaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(self.guihao_manage_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_archives_hbox.Add(manage_archives_hehaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(self.hehao_manage_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_archives_hbox.Add(manage_archives_juanhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(self.juanhao_manage_archives_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_archives_hbox.Add(manage_archives_anjuantimingText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(self.anjuantiming_manage_archives_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_archives_hbox.Add(manage_archives_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_archives_hbox.Add(manage_archives_modifyButton,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=120)
        manage_archives_hbox.Add(manage_archives_deleteButton,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        #向纵向布局管理器中添加横向布局管理器
        manage_archives_vbox.Add(manage_archives_hbox)
        #生成数据显示控件和数据源表格
        self.manage_grid_archives = wx.grid.Grid(parent=self.manage_archives_panel,size=GRID_SIZE)
        self.manage_data_table = myArchiveTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_manage_table_rows = self.manage_data_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.manage_grid_archives.SetTable(self.manage_data_table,takeOwnership=True)
        self.manage_grid_archives.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.activeRow)
        self.manage_grid_archives.AutoSize()
        #向纵向布局管理器中添加数据表格
        manage_archives_vbox.Add(self.manage_grid_archives,flag=wx.ALIGN_BOTTOM)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.manage_archives_panel.SetSizer(manage_archives_vbox)        
        #向档案管理框架中加新增、查询、管理页,父元素为档案管理页
        self.archivesManage_Notebook.AddPage(self.add_archives_panel,u"档案新增")
        self.archivesManage_Notebook.AddPage(self.select_archives_panel,u"档案查询")
        self.archivesManage_Notebook.AddPage(self.manage_archives_panel,u"档案操作")
        #向主框架中加入档案管理页
        self.main_Notebook.AddPage(self.archivesManage_panel,u"档案管理")
        #**********档案管理页结束**********
        #self.fileManage_Notebook = wx.Notebook(parent=self.panel,size=(600,300))
        #**********文件管理页开始**********
        self.fileManage_panel = wx.Panel(parent=self.main_Notebook)
        #文件管理框架
        self.fileManage_Notebook = wx.Notebook(parent=self.fileManage_panel,size=SUB_FRAME_SIZE)
        #文件管理框架中新增页
        self.add_files_panel = wx.Panel(parent=self.fileManage_Notebook)
        #生成GridBagSizer布局管理器
        add_files_gridBagSizer = wx.GridBagSizer(vgap=30,hgap=10)
        add_files_hboxsizer = wx.BoxSizer(wx.HORIZONTAL)
        #生成文本标签
        add_files_quhaoText = wx.StaticText(parent=self.add_files_panel,label=u"区        号:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_guihaoText = wx.StaticText(parent=self.add_files_panel,label=u"柜        号:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_hehaoText = wx.StaticText(parent=self.add_files_panel,label=u"盒        号:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_juanhaoText = wx.StaticText(parent=self.add_files_panel,label=u"卷        号:",style=wx.ALIGN_CENTER_VERTICAL)  
        add_files_wenjiantimuText = wx.StaticText(parent=self.add_files_panel,label=u"文件题目:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_wenjianbianhaoText = wx.StaticText(parent=self.add_files_panel,label=u"文件编号:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_fawendanweiText = wx.StaticText(parent=self.add_files_panel,label=u"发文单位:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_xingchengriqiText = wx.StaticText(parent=self.add_files_panel,label=u"形成日期:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_yeshuText = wx.StaticText(parent=self.add_files_panel,label=u"页        数:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_mijiText = wx.StaticText(parent=self.add_files_panel,label=u"密        级:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_neirongText = wx.StaticText(parent=self.add_files_panel,label=u"附        件:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_beizhuText = wx.StaticText(parent=self.add_files_panel,label=u"备        注:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_inputterText = wx.StaticText(parent=self.add_files_panel,label=u"录  入  人:",style=wx.ALIGN_CENTER_VERTICAL)
        add_files_font = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        add_files_quhaoText.SetFont(add_files_font)
        add_files_guihaoText.SetFont(add_files_font)
        add_files_hehaoText.SetFont(add_files_font)
        add_files_juanhaoText.SetFont(add_files_font)
        add_files_wenjiantimuText.SetFont(add_files_font)
        add_files_wenjianbianhaoText.SetFont(add_files_font)
        add_files_fawendanweiText.SetFont(add_files_font)
        add_files_xingchengriqiText.SetFont(add_files_font)
        add_files_yeshuText.SetFont(add_files_font)
        add_files_mijiText.SetFont(add_files_font)
        add_files_neirongText.SetFont(add_files_font)
        add_files_beizhuText.SetFont(add_files_font)
        add_files_inputterText.SetFont(add_files_font)   
        #将文本标签纳入add_archives_gridBagSizer布局管理器中
        add_files_gridBagSizer.Add(add_files_quhaoText,pos=(0,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        add_files_gridBagSizer.Add(add_files_guihaoText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_hehaoText,pos=(2,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_juanhaoText,pos=(3,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_wenjianbianhaoText,pos=(4,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_wenjiantimuText,pos=(5,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)        
        add_files_gridBagSizer.Add(add_files_neirongText,pos=(6,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_fawendanweiText,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        add_files_gridBagSizer.Add(add_files_xingchengriqiText,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_yeshuText,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_mijiText,pos=(3,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_beizhuText,pos=(4,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        add_files_gridBagSizer.Add(add_files_inputterText,pos=(5,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT) 
        #生成  区号  下拉选择框
        self.quhao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=(100,30),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.quhao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.quhao_files_comboBox.SetPopupMaxHeight(self.quhao_files_comboBox.GetCharHeight()*10)
        self.quhao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createjuanhaoByDanwei)
        #生成  柜号  下拉选择框
        self.guihao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=(100,30),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.guihao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.guihao_files_comboBox.SetPopupMaxHeight(self.guihao_files_comboBox.GetCharHeight()*10)
        self.guihao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createjuanhaoByDanwei)
        #生成  盒号 下拉选择框
        self.hehao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=(100,30),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.hehao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.hehao_files_comboBox.SetPopupMaxHeight(self.hehao_files_comboBox.GetCharHeight()*10)
        self.hehao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createjuanhaoByDanwei)
        #生成  卷号  下拉选择框
        temp_list = DoCURD.query_juanhao_values(self.CONN,{"table_name":"archives","quhao":self.quhao_files_comboBox.GetValue(),"guihao":self.guihao_files_comboBox.GetValue(),"hehao":self.hehao_files_comboBox.GetValue()})
        self.temp_juanhao_list = []
        for row in temp_list:
            self.temp_juanhao_list.append(row[0])
        self.temp_juanhao_list.sort()        
        self.juanhao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=(100,30),choices=self.temp_juanhao_list,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        if len(self.temp_juanhao_list) != 0 :
            self.juanhao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.juanhao_files_comboBox.SetPopupMaxHeight(self.juanhao_files_comboBox.GetCharHeight()*10)
        #选择卷号后自动显示案卷题名
        self.add_files_anjuantimingText = wx.StaticText(parent=self.add_files_panel,label=u'',style=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        self.showAnjuantiming('')
        add_files_hboxsizer.Add(self.juanhao_files_comboBox)
        add_files_hboxsizer.Add(self.add_files_anjuantimingText,flag=wx.ALIGN_CENTER_VERTICAL)
        #将下拉卷号绑定事件，显示案卷题名
        self.juanhao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.showAnjuantiming)
        #生成 文件编号 输入框
        self.wenjianbianhao_files_ctrl = wx.TextCtrl(parent=self.add_files_panel,size=(450,30))
        #生成 文件题目 输入框
        self.wenjiantimu_files_ctrl = wx.TextCtrl(parent=self.add_files_panel,size=(450,30))
        #生成文件内容上传按钮
        self.neirong_file_FileBrowseButton = wx.lib.filebrowsebutton.FileBrowseButton(
                                                                                     parent=self.add_files_panel, 
                                                                                     size=(745,60), 
                                                                                     labelText='', 
                                                                                     buttonText=u"上传", 
                                                                                     toolTip=u"请点击导入按钮,并选择文件的电子稿", 
                                                                                     dialogTitle=u"请选择要导入的电子稿..."
                                                                                     )
        #生成 发文单位 输入框
        self.fawendanwei_files_ctrl = wx.TextCtrl(parent=self.add_files_panel,size=ARCHIVES_SIZE)
        #生成 形成日期 下拉选择框
        self.xingchengriqi_files_pop = UtilData.DateControl(self.add_files_panel, -1, pos = (30,30),size=ARCHIVES_SIZE)
        today = datetime.datetime.now().strftime('%Y%m%d')
        #设置 形成日期 默认值
        self.xingchengriqi_files_pop.SetValue(unicode(today[:4]+"年"+today[4:6]+"月"+today[6:]+"日"))
        #生成 页数 下拉选择框
        self.yeshu_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=ARCHIVES_SIZE,choices=[unicode(i) for i in range(1,1001)],style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.yeshu_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.yeshu_files_comboBox.SetPopupMaxHeight(self.yeshu_files_comboBox.GetCharHeight()*10)
        #生成 密级 下拉选择框
        self.miji_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.add_files_panel,id=-1,size=ARCHIVES_SIZE,choices=UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.miji_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.miji_files_comboBox.SetPopupMaxHeight(self.miji_files_comboBox.GetCharHeight()*10)
        #生成 备注 输入框
        self.beizhu_files_ctrl = wx.TextCtrl(parent=self.add_files_panel,size=ARCHIVES_SIZE)     
        #生成 录入人 输入框
        self.inputter_files_ctrl = wx.TextCtrl(parent=self.add_files_panel,size=ARCHIVES_SIZE) 
        #生成 新增档案 按钮并进行绑定
        self.newFile_button = wx.Button(parent=self.add_files_panel,label=u'添加文件')
        self.newFile_button.Bind(wx.EVT_BUTTON, self.addNewFile)
        #将各选择与输入控件加入布局管理器
        add_files_gridBagSizer.Add(self.quhao_files_comboBox,pos=(0,1),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        add_files_gridBagSizer.Add(self.guihao_files_comboBox,pos=(1,1),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.hehao_files_comboBox,pos=(2,1),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(add_files_hboxsizer,pos=(3,1),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.wenjianbianhao_files_ctrl,pos=(4,1),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.wenjiantimu_files_ctrl,pos=(5,1),flag=wx.ALIGN_LEFT) 
        #此处放置 内容 pos=(6,1)
        add_files_gridBagSizer.Add(self.neirong_file_FileBrowseButton,pos=(6,1),span=(1,3),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.fawendanwei_files_ctrl,pos=(0,3),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        add_files_gridBagSizer.Add(self.xingchengriqi_files_pop,pos=(1,3),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.yeshu_files_comboBox,pos=(2,3),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.miji_files_comboBox,pos=(3,3),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.beizhu_files_ctrl,pos=(4,3),flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        add_files_gridBagSizer.Add(self.inputter_files_ctrl,pos=(5,3),flag=wx.ALIGN_LEFT)
        add_files_gridBagSizer.Add(self.newFile_button,pos=(7,0),span=(1,4),flag=wx.ALIGN_CENTER)    
        #各列尽可能扩展
        add_files_gridBagSizer.AddGrowableCol(0)
        add_files_gridBagSizer.AddGrowableCol(3)
        #设置add_files_gridBagSizer为add_files_panel的布局管理器
        self.add_files_panel.SetSizer(add_files_gridBagSizer)
        #文件管理框架中查询页
        self.select_files_panel = wx.Panel(parent=self.fileManage_Notebook)
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
        select_files_archiveIDText = wx.StaticText(parent=self.select_files_panel,label=u'档案ID:')
        self.archiveID_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_files_panel,id=-1,choices=archiveIDs,style=wx.CB_READONLY)
        self.archiveID_select_files_comboBox.SetSelection(n=0)
        self.archiveID_select_files_comboBox.SetPopupMaxHeight(self.archiveID_select_files_comboBox.GetCharHeight()*10)
        SENIOR_SELECT_SIZE_FOR_FILES=(80,25)
        #文件编号
        select_files_wenjianbianhaoText = wx.StaticText(parent=self.select_files_panel,label=u"文件编号:")
        self.wenjianbianhao_select_files_ctrl = wx.TextCtrl(parent=self.select_files_panel,size=SENIOR_SELECT_SIZE_FOR_FILES)
        #文件题目
        select_files_wenjiantimuText = wx.StaticText(parent=self.select_files_panel,label=u"文件题目:")
        self.wenjiantimu_select_files_Ctrl = wx.TextCtrl(parent=self.select_files_panel,size=SENIOR_SELECT_SIZE_FOR_FILES)
        #发文单位
        select_files_fawendangweiText = wx.StaticText(parent=self.select_files_panel,label=u"发文单位:")
        self.fawendanwei_select_files_Ctrl = wx.TextCtrl(parent=self.select_files_panel,size=SENIOR_SELECT_SIZE_FOR_FILES)  
        #形成日期
        select_files_xingchengriqiText = wx.StaticText(parent=self.select_files_panel,label=u"形成日期:")
        self.xingchengriqiText_select_files_pop = UtilData.DateControl(self.select_files_panel)               
        #密级
        select_files_mijiText = wx.StaticText(parent=self.select_files_panel,label=u"密级:")
        self.miji_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_files_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.MIJI_List,style=wx.CB_READONLY)
        self.miji_select_files_comboBox.SetSelection(n=0)
        self.miji_select_files_comboBox.SetPopupMaxHeight(self.miji_select_files_comboBox.GetCharHeight()*10)        
        #存放位置
        select_files_weizhiText = wx.StaticText(parent=self.select_files_panel,label=u"位置:")
        self.weizhi_select_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.select_files_panel,id=-1,size=SENIOR_SELECT_SIZE,choices=[""]+UtilData.WEIZHI_List,style=wx.CB_READONLY)
        self.weizhi_select_files_comboBox.SetSelection(n=0)
        self.weizhi_select_files_comboBox.SetPopupMaxHeight(self.weizhi_select_files_comboBox.GetCharHeight()*10)        
        #备注
        select_files_beizhuText = wx.StaticText(parent=self.select_files_panel,label=u"备注:")
        self.beizhu_select_files_Ctrl = wx.TextCtrl(parent=self.select_files_panel,size=SENIOR_SELECT_SIZE_FOR_FILES)          
        
        #查询按钮
        select_files_button = wx.Button(parent=self.select_files_panel,label=u"查询")
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
        select_files_vbox.Add(select_files_hbox)
        #生成数据显示控件和数据源表格
        self.select_grid_files = wx.grid.Grid(parent=self.select_files_panel,size=GRID_SIZE)
        self.select_data_file_table = myFileTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_select_file_table_rows = self.select_data_file_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.select_grid_files.SetTable(self.select_data_file_table,takeOwnership=True)
        self.select_grid_files.AutoSize()
        #向纵向布局管理器中添加数据表格
        select_files_vbox.Add(self.select_grid_files,flag=wx.ALIGN_BOTTOM)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.select_files_panel.SetSizer(select_files_vbox)        
        #文件管理框架中管理页
        self.manage_files_panel = wx.Panel(parent=self.fileManage_Notebook)
        #档案查询页中的总垂直布局管理器和顶部横向布局管理器
        manage_files_vbox = wx.BoxSizer(wx.VERTICAL)
        manage_files_hbox = wx.BoxSizer(wx.HORIZONTAL)
        #档案ID
          #查询档案表中共有多少档案ID,都是什么
        manage_temp_archiveIDList = DoCURD.query_archiveID_for_selectfile(self.CONN,paramDict={"table_name":"archives"})
        manage_archiveIDs = []    
        for row in manage_temp_archiveIDList:
            for col in row :
                manage_archiveIDs.append(unicode(col))
        manage_files_archiveIDText = wx.StaticText(parent=self.manage_files_panel,label=u'档案ID:')
        self.archiveID_manage_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.manage_files_panel,id=-1,choices=[""]+manage_archiveIDs,style=wx.CB_READONLY)
        self.archiveID_manage_files_comboBox.SetSelection(n=0)
        self.archiveID_manage_files_comboBox.SetPopupMaxHeight(self.archiveID_manage_files_comboBox.GetCharHeight()*10)
        #文件编号
        manage_files_wenjianbianhaoText = wx.StaticText(parent=self.manage_files_panel,label=u"文件编号:")
        self.wenjianbianhao_manage_files_ctrl = wx.TextCtrl(parent=self.manage_files_panel)
        #文件题目
        manage_files_wenjiantimuText = wx.StaticText(parent=self.manage_files_panel,label=u"文件题目:")
        self.wenjiantimu_manage_files_Ctrl = wx.TextCtrl(parent=self.manage_files_panel)
        #查询按钮
        selectManage_files_button = wx.Button(parent=self.manage_files_panel,label=u"查询")
        #对查询按钮进行事件绑定
        selectManage_files_button.Bind(wx.EVT_BUTTON, self.selectManageFilesBySomeConditions)
        #打开按钮(默认按钮为失效状态,只有当前选中的文件有电子附件,才能变成可用状态)
        self.openManage_files_button = wx.Button(parent=self.manage_files_panel,label=u"打开")
        self.openManage_files_button.Enable(False)
        self.OPENFILE = u''
        #打开按钮进行事件绑定
        self.openManage_files_button.Bind(wx.EVT_BUTTON, self.openfile)         
        #修改按钮
        modifyManage_files_button = wx.Button(parent=self.manage_files_panel,label=u"修改")
        #修改询按钮进行事件绑定
        modifyManage_files_button.Bind(wx.EVT_BUTTON, self.modifyManageFilesBySomeConditions)  
        #删除按钮
        deleteManage_files_button = wx.Button(parent=self.manage_files_panel,label=u"删除")
        #对删除按钮进行事件绑定
        deleteManage_files_button.Bind(wx.EVT_BUTTON, self.deleteManageFilesBySomeConditions)         
        #向顶部横向布局管理器中添加元素
        manage_files_hbox.Add(manage_files_archiveIDText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_files_hbox.Add(self.archiveID_manage_files_comboBox,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)               
        manage_files_hbox.Add(manage_files_wenjianbianhaoText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_files_hbox.Add(self.wenjianbianhao_manage_files_ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_files_hbox.Add(manage_files_wenjiantimuText,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_files_hbox.Add(self.wenjiantimu_manage_files_Ctrl,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=5)
        manage_files_hbox.Add(selectManage_files_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_files_hbox.Add(self.openManage_files_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=120)
        manage_files_hbox.Add(modifyManage_files_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        manage_files_hbox.Add(deleteManage_files_button,flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT,border=10)
        #向纵向布局管理器中添加横向布局管理器
        manage_files_vbox.Add(manage_files_hbox)
        #生成数据显示控件和数据源表格
        self.manage_grid_files = wx.grid.Grid(parent=self.manage_files_panel,size=GRID_SIZE)
        self.manage_data_file_table = myFileTable()
        #获得表格初始化后的行数,以last命名,为了将来动态删除多余行用
        self.last_data_manage_file_table_rows = self.manage_data_file_table.GetRowsCount()
        #将数据源表格设置到数据显示控件中
        self.manage_grid_files.SetTable(self.manage_data_file_table,takeOwnership=True)
        self.manage_grid_files.Bind(wx.grid.EVT_GRID_CMD_CELL_LEFT_CLICK, self.activeRow_ManageFiles)
        self.manage_grid_files.AutoSize()
        #向纵向布局管理器中添加数据表格
        manage_files_vbox.Add(self.manage_grid_files,flag=wx.ALIGN_BOTTOM)
        #将纵向布局管理器设置为档案查询页的布局管理器
        self.manage_files_panel.SetSizer(manage_files_vbox)                
        #向文件管理框架中添加新增、查询、管理页,父元素为文件管理页
        self.fileManage_Notebook.AddPage(self.add_files_panel,u"文件新增")
        self.fileManage_Notebook.AddPage(self.select_files_panel,u"文件查询")
        self.fileManage_Notebook.AddPage(self.manage_files_panel,u"文件操作")
        #向主框架中加入文件管理页
        self.main_Notebook.AddPage(self.fileManage_panel,u"文件管理")   
        #**********文件管理页结束**********
        
        #**********打印页开始**********
        self.print_panel = wx.Panel(parent=self.main_Notebook)
        #打印框架
        self.print_Notebook = wx.Notebook(parent=self.print_panel,size=SUB_FRAME_SIZE)
        #打印柜中档案表
        self.printGuiArchives_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_vbox = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号  文本及下拉选择框
        quhaoText_print = wx.StaticText(parent=self.printGuiArchives_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print = wx.StaticText(parent=self.printGuiArchives_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        guihaoText_print.SetFont(font_print)
        quhaoText_print.SetFont(font_print)
        self.guihao_print_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.printGuiArchives_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)        
        self.guihao_print_comboBox.SetPopupMaxHeight(self.guihao_print_comboBox.GetCharHeight()*10) 
        self.quhao_print_comboBox = wx.combo.OwnerDrawnComboBox(parent=self.printGuiArchives_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        guihaoPrint_Button = wx.Button(parent=self.printGuiArchives_panel,label=u'打印')
        guihaoPrint_Button.Bind(wx.EVT_BUTTON, self.printByGuiHao)
        quhao_print_hbox.Add(quhaoText_print,flag=wx.ALIGN_CENTER)
        quhao_print_hbox.Add(self.quhao_print_comboBox,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox.Add(guihaoText_print,flag=wx.ALIGN_CENTER)
        guihao_print_hbox.Add(self.guihao_print_comboBox,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)    
        guihao_print_vbox.Add(quhao_print_hbox,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        guihao_print_vbox.Add(guihao_print_hbox,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        guihao_print_vbox.Add(guihaoPrint_Button,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        self.printGuiArchives_panel.SetSizer(guihao_print_vbox)
        #打印盒中档案表
        self.printHeArchives_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox_for_he = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox_for_he = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_hbox_for_he = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_vbox_for_he = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号 盒号 文本及下拉选择框
        quhaoText_print_for_he = wx.StaticText(parent=self.printHeArchives_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print_for_he = wx.StaticText(parent=self.printHeArchives_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        hehaoText_print_for_he = wx.StaticText(parent=self.printHeArchives_panel,label=u"请选择盒号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        quhaoText_print_for_he.SetFont(font_print)
        guihaoText_print_for_he.SetFont(font_print)
        hehaoText_print_for_he.SetFont(font_print)
        self.quhao_print_comboBox_for_he = wx.combo.OwnerDrawnComboBox(parent=self.printHeArchives_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        self.guihao_print_comboBox_for_he = wx.combo.OwnerDrawnComboBox(parent=self.printHeArchives_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)       
        self.guihao_print_comboBox_for_he.SetPopupMaxHeight(self.guihao_print_comboBox_for_he.GetCharHeight()*10) 
        self.hehao_print_comboBox_for_he = wx.combo.OwnerDrawnComboBox(parent=self.printHeArchives_panel,id=-1,size=(120,28),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        self.hehao_print_comboBox_for_he.SetPopupMaxHeight(self.hehao_print_comboBox_for_he.GetCharHeight()*10) 
        hehaoPrint_Button_for_he = wx.Button(parent=self.printHeArchives_panel,label=u'打印')
        #柜打印和盒打印模板相同
        hehaoPrint_Button_for_he.Bind(wx.EVT_BUTTON, self.printByHeHao)
        quhao_print_hbox_for_he.Add(quhaoText_print_for_he,flag=wx.ALIGN_CENTER)
        quhao_print_hbox_for_he.Add(self.quhao_print_comboBox_for_he,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox_for_he.Add(guihaoText_print_for_he,flag=wx.ALIGN_CENTER)
        guihao_print_hbox_for_he.Add(self.guihao_print_comboBox_for_he,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_hbox_for_he.Add(hehaoText_print_for_he,flag=wx.ALIGN_CENTER)
        hehao_print_hbox_for_he.Add(self.hehao_print_comboBox_for_he,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_vbox_for_he.Add(quhao_print_hbox_for_he,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        hehao_print_vbox_for_he.Add(guihao_print_hbox_for_he,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        hehao_print_vbox_for_he.Add(hehao_print_hbox_for_he,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        hehao_print_vbox_for_he.Add(hehaoPrint_Button_for_he,flag=wx.ALIGN_CENTER|wx.TOP,border=30)        
        self.printHeArchives_panel.SetSizer(hehao_print_vbox_for_he)
        #打印卷中文件表
        self.printJuanArchives_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox_for_juan = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox_for_juan = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_hbox_for_juan = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_hbox_for_juan = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_vbox_for_juan = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号 盒号 文本及下拉选择框
        quhaoText_print_for_juan = wx.StaticText(parent=self.printJuanArchives_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print_for_juan = wx.StaticText(parent=self.printJuanArchives_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        hehaoText_print_for_juan = wx.StaticText(parent=self.printJuanArchives_panel,label=u"请选择盒号 :",style=wx.ALIGN_CENTER)
        juanhaoText_print_for_juan = wx.StaticText(parent=self.printJuanArchives_panel,label=u"请选择卷号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        quhaoText_print_for_juan.SetFont(font_print)
        guihaoText_print_for_juan.SetFont(font_print)
        hehaoText_print_for_juan.SetFont(font_print)
        juanhaoText_print_for_juan.SetFont(font_print)
        self.quhao_print_comboBox_for_juan = wx.combo.OwnerDrawnComboBox(parent=self.printJuanArchives_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        self.guihao_print_comboBox_for_juan = wx.combo.OwnerDrawnComboBox(parent=self.printJuanArchives_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)       
        self.guihao_print_comboBox_for_juan.SetPopupMaxHeight(self.guihao_print_comboBox_for_juan.GetCharHeight()*10) 
        self.hehao_print_comboBox_for_juan = wx.combo.OwnerDrawnComboBox(parent=self.printJuanArchives_panel,id=-1,size=(120,28),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        self.hehao_print_comboBox_for_juan.SetPopupMaxHeight(self.hehao_print_comboBox_for_juan.GetCharHeight()*10) 
        self.juanhao_print_comboBox_for_juan = wx.combo.OwnerDrawnComboBox(parent=self.printJuanArchives_panel,id=-1,size=(120,28),choices=UtilData.JUANHAO_List,style=wx.CB_READONLY)
        self.juanhao_print_comboBox_for_juan.SetPopupMaxHeight(self.juanhao_print_comboBox_for_juan.GetCharHeight()*10) 
        juanhaoPrint_Button_for_juan = wx.Button(parent=self.printJuanArchives_panel,label=u'打印')
        juanhaoPrint_Button_for_juan.Bind(wx.EVT_BUTTON, self.printByJuanHao)
        quhao_print_hbox_for_juan.Add(quhaoText_print_for_juan,flag=wx.ALIGN_CENTER)
        quhao_print_hbox_for_juan.Add(self.quhao_print_comboBox_for_juan,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox_for_juan.Add(guihaoText_print_for_juan,flag=wx.ALIGN_CENTER)
        guihao_print_hbox_for_juan.Add(self.guihao_print_comboBox_for_juan,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_hbox_for_juan.Add(hehaoText_print_for_juan,flag=wx.ALIGN_CENTER)
        hehao_print_hbox_for_juan.Add(self.hehao_print_comboBox_for_juan,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        juanhao_print_hbox_for_juan.Add(juanhaoText_print_for_juan,flag=wx.ALIGN_CENTER)
        juanhao_print_hbox_for_juan.Add(self.juanhao_print_comboBox_for_juan,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        juanhao_print_vbox_for_juan.Add(quhao_print_hbox_for_juan,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        juanhao_print_vbox_for_juan.Add(guihao_print_hbox_for_juan,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_juan.Add(hehao_print_hbox_for_juan,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_juan.Add(juanhao_print_hbox_for_juan,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_juan.Add(juanhaoPrint_Button_for_juan,flag=wx.ALIGN_CENTER|wx.TOP,border=30)        
        self.printJuanArchives_panel.SetSizer(juanhao_print_vbox_for_juan)
        #打印案卷标题
        self.printAnJuanTitle_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox_for_anjuanbiaoti = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox_for_anjuanbiaoti = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_hbox_for_anjuanbiaoti = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_hbox_for_anjuanbiaoti = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_vbox_for_anjuanbiaoti = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号 盒号 文本及下拉选择框
        quhaoText_print_for_anjuanbiaoti = wx.StaticText(parent=self.printAnJuanTitle_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print_for_anjuanbiaoti = wx.StaticText(parent=self.printAnJuanTitle_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        hehaoText_print_for_anjuanbiaoti = wx.StaticText(parent=self.printAnJuanTitle_panel,label=u"请选择盒号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        quhaoText_print_for_anjuanbiaoti.SetFont(font_print)
        guihaoText_print_for_anjuanbiaoti.SetFont(font_print)
        hehaoText_print_for_anjuanbiaoti.SetFont(font_print)
        self.quhao_print_comboBox_for_anjuanbiaoti = wx.combo.OwnerDrawnComboBox(parent=self.printAnJuanTitle_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        self.guihao_print_comboBox_for_anjuanbiaoti = wx.combo.OwnerDrawnComboBox(parent=self.printAnJuanTitle_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)       
        self.guihao_print_comboBox_for_anjuanbiaoti.SetPopupMaxHeight(self.guihao_print_comboBox_for_anjuanbiaoti.GetCharHeight()*10) 
        self.hehao_print_comboBox_for_anjuanbiaoti = wx.combo.OwnerDrawnComboBox(parent=self.printAnJuanTitle_panel,id=-1,size=(120,28),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        self.hehao_print_comboBox_for_anjuanbiaoti.SetPopupMaxHeight(self.hehao_print_comboBox_for_anjuanbiaoti.GetCharHeight()*10) 
        juanhaoPrint_Button_for_anjuanbiaoti = wx.Button(parent=self.printAnJuanTitle_panel,label=u'打印')
        juanhaoPrint_Button_for_anjuanbiaoti.Bind(wx.EVT_BUTTON, self.printByAnJuanBiaoTi)
        quhao_print_hbox_for_anjuanbiaoti.Add(quhaoText_print_for_anjuanbiaoti,flag=wx.ALIGN_CENTER)
        quhao_print_hbox_for_anjuanbiaoti.Add(self.quhao_print_comboBox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox_for_anjuanbiaoti.Add(guihaoText_print_for_anjuanbiaoti,flag=wx.ALIGN_CENTER)
        guihao_print_hbox_for_anjuanbiaoti.Add(self.guihao_print_comboBox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_hbox_for_anjuanbiaoti.Add(hehaoText_print_for_anjuanbiaoti,flag=wx.ALIGN_CENTER)
        hehao_print_hbox_for_anjuanbiaoti.Add(self.hehao_print_comboBox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        juanhao_print_vbox_for_anjuanbiaoti.Add(quhao_print_hbox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        juanhao_print_vbox_for_anjuanbiaoti.Add(guihao_print_hbox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_anjuanbiaoti.Add(hehao_print_hbox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_anjuanbiaoti.Add(juanhao_print_hbox_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_anjuanbiaoti.Add(juanhaoPrint_Button_for_anjuanbiaoti,flag=wx.ALIGN_CENTER|wx.TOP,border=30)                
        self.printAnJuanTitle_panel.SetSizer(juanhao_print_vbox_for_anjuanbiaoti)
        #打印财务用表
        self.printCaiWu_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox_for_caiwu = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox_for_caiwu = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_hbox_for_caiwu = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_vbox_for_caiwu = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号 盒号 文本及下拉选择框
        quhaoText_print_for_caiwu = wx.StaticText(parent=self.printCaiWu_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print_for_caiwu = wx.StaticText(parent=self.printCaiWu_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        hehaoText_print_for_caiwu = wx.StaticText(parent=self.printCaiWu_panel,label=u"请选择盒号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        quhaoText_print_for_caiwu.SetFont(font_print)
        guihaoText_print_for_caiwu.SetFont(font_print)
        hehaoText_print_for_caiwu.SetFont(font_print)
        self.quhao_print_comboBox_for_caiwu = wx.combo.OwnerDrawnComboBox(parent=self.printCaiWu_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        self.guihao_print_comboBox_for_caiwu = wx.combo.OwnerDrawnComboBox(parent=self.printCaiWu_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)       
        self.guihao_print_comboBox_for_caiwu.SetPopupMaxHeight(self.guihao_print_comboBox_for_anjuanbiaoti.GetCharHeight()*10) 
        self.hehao_print_comboBox_for_caiwu = wx.combo.OwnerDrawnComboBox(parent=self.printCaiWu_panel,id=-1,size=(120,28),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        self.hehao_print_comboBox_for_caiwu.SetPopupMaxHeight(self.hehao_print_comboBox_for_anjuanbiaoti.GetCharHeight()*10) 
        juanhaoPrint_Button_for_caiwu = wx.Button(parent=self.printCaiWu_panel,label=u'打印')
        juanhaoPrint_Button_for_caiwu.Bind(wx.EVT_BUTTON, self.printCaiWu)
        quhao_print_hbox_for_caiwu.Add(quhaoText_print_for_caiwu,flag=wx.ALIGN_CENTER)
        quhao_print_hbox_for_caiwu.Add(self.quhao_print_comboBox_for_caiwu,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox_for_caiwu.Add(guihaoText_print_for_caiwu,flag=wx.ALIGN_CENTER)
        guihao_print_hbox_for_caiwu.Add(self.guihao_print_comboBox_for_caiwu,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_hbox_for_caiwu.Add(hehaoText_print_for_caiwu,flag=wx.ALIGN_CENTER)
        hehao_print_hbox_for_caiwu.Add(self.hehao_print_comboBox_for_caiwu,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        juanhao_print_vbox_for_caiwu.Add(quhao_print_hbox_for_caiwu,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        juanhao_print_vbox_for_caiwu.Add(guihao_print_hbox_for_caiwu,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_caiwu.Add(hehao_print_hbox_for_caiwu,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        juanhao_print_vbox_for_caiwu.Add(juanhaoPrint_Button_for_caiwu,flag=wx.ALIGN_CENTER|wx.TOP,border=30)        
        self.printCaiWu_panel.SetSizer(juanhao_print_vbox_for_caiwu)
        #打印备考表
        self.printBeikao_panel = wx.Panel(parent=self.print_Notebook)
        quhao_print_hbox_for_beikao = wx.BoxSizer(wx.HORIZONTAL)
        guihao_print_hbox_for_beikao = wx.BoxSizer(wx.HORIZONTAL)
        hehao_print_hbox_for_beikao = wx.BoxSizer(wx.HORIZONTAL)
        juanhao_print_hbox_for_beikao = wx.BoxSizer(wx.HORIZONTAL)
        beikao_print_vbox_for_beikao = wx.BoxSizer(wx.VERTICAL)
        #生成  区号、柜号 盒号 文本及下拉选择框
        quhaoText_print_for_beikao = wx.StaticText(parent=self.printBeikao_panel,label=u"请选择区号 :",style=wx.ALIGN_CENTER)
        guihaoText_print_for_beikao = wx.StaticText(parent=self.printBeikao_panel,label=u"请选择柜号 :",style=wx.ALIGN_CENTER)
        hehaoText_print_for_beikao = wx.StaticText(parent=self.printBeikao_panel,label=u"请选择盒号 :",style=wx.ALIGN_CENTER)
        juanhaoText_print_for_beikao = wx.StaticText(parent=self.printBeikao_panel,label=u"请选择卷号 :",style=wx.ALIGN_CENTER)
        font_print = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        quhaoText_print_for_beikao.SetFont(font_print)
        guihaoText_print_for_beikao.SetFont(font_print)
        hehaoText_print_for_beikao.SetFont(font_print)
        juanhaoText_print_for_beikao.SetFont(font_print)
        self.quhao_print_comboBox_for_beikao = wx.combo.OwnerDrawnComboBox(parent=self.printBeikao_panel,id=-1,size=(120,28),choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        self.guihao_print_comboBox_for_beikao = wx.combo.OwnerDrawnComboBox(parent=self.printBeikao_panel,id=-1,size=(120,28),choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)       
        self.guihao_print_comboBox_for_beikao.SetPopupMaxHeight(self.guihao_print_comboBox_for_beikao.GetCharHeight()*10) 
        self.hehao_print_comboBox_for_beikao = wx.combo.OwnerDrawnComboBox(parent=self.printBeikao_panel,id=-1,size=(120,28),choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        self.hehao_print_comboBox_for_beikao.SetPopupMaxHeight(self.hehao_print_comboBox_for_beikao.GetCharHeight()*10) 
        self.juanhao_print_comboBox_for_beikao = wx.combo.OwnerDrawnComboBox(parent=self.printBeikao_panel,id=-1,size=(120,28),choices=UtilData.JUANHAO_List,style=wx.CB_READONLY)
        self.juanhao_print_comboBox_for_beikao.SetPopupMaxHeight(self.juanhao_print_comboBox_for_beikao.GetCharHeight()*10) 
        juanhaoPrint_Button_for_beikao = wx.Button(parent=self.printBeikao_panel,label=u'打印')
        juanhaoPrint_Button_for_beikao.Bind(wx.EVT_BUTTON, self.printBeikao)
        quhao_print_hbox_for_beikao.Add(quhaoText_print_for_beikao,flag=wx.ALIGN_CENTER)
        quhao_print_hbox_for_beikao.Add(self.quhao_print_comboBox_for_beikao,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        guihao_print_hbox_for_beikao.Add(guihaoText_print_for_beikao,flag=wx.ALIGN_CENTER)
        guihao_print_hbox_for_beikao.Add(self.guihao_print_comboBox_for_beikao,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        hehao_print_hbox_for_beikao.Add(hehaoText_print_for_beikao,flag=wx.ALIGN_CENTER)
        hehao_print_hbox_for_beikao.Add(self.hehao_print_comboBox_for_beikao,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        juanhao_print_hbox_for_beikao.Add(juanhaoText_print_for_beikao,flag=wx.ALIGN_CENTER)
        juanhao_print_hbox_for_beikao.Add(self.juanhao_print_comboBox_for_beikao,flag=wx.ALIGN_CENTER|wx.LEFT,border=5)
        beikao_print_vbox_for_beikao.Add(quhao_print_hbox_for_beikao,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        beikao_print_vbox_for_beikao.Add(guihao_print_hbox_for_beikao,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        beikao_print_vbox_for_beikao.Add(hehao_print_hbox_for_beikao,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        beikao_print_vbox_for_beikao.Add(juanhao_print_hbox_for_beikao,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        beikao_print_vbox_for_beikao.Add(juanhaoPrint_Button_for_beikao,flag=wx.ALIGN_CENTER|wx.TOP,border=30)            
        self.printBeikao_panel.SetSizer(beikao_print_vbox_for_beikao)
        
        self.print_Notebook.AddPage(self.printGuiArchives_panel,u"打印柜中档案表")
        self.print_Notebook.AddPage(self.printHeArchives_panel,u'打印盒中档案表')
        self.print_Notebook.AddPage(self.printJuanArchives_panel,u'打印卷中文件表')
        self.print_Notebook.AddPage(self.printAnJuanTitle_panel,u'打印案卷标题')
        self.print_Notebook.AddPage(self.printCaiWu_panel,u'打印财务用表')
        self.print_Notebook.AddPage(self.printBeikao_panel,u'打印备考表')
        #向主框架页中加入打印页
        self.main_Notebook.AddPage(self.print_panel,u"打印")    
        #**********打印页结束**********
        #**********档案归档页开始**********
        self.placeOnFile_panel = wx.Panel(parent=self.main_Notebook)
        #档案归档框架
        self.placeOnFile_Notebook = wx.Notebook(parent=self.placeOnFile_panel,size=SUB_FRAME_SIZE)
        #归档页
        self.pigeonhole_placeOnFile_panel = wx.Panel(parent=self.placeOnFile_Notebook)
        pigeonhole_HBox = wx.BoxSizer(wx.HORIZONTAL)
        quxiao_HBox = wx.BoxSizer(wx.HORIZONTAL)
        pigeonhole_VBox = wx.BoxSizer(wx.VERTICAL)
        guidangnianxianText = wx.StaticText(parent=self.pigeonhole_placeOnFile_panel,label=u"请选择归档年限: ",style=wx.ALIGN_CENTER)
        quxiaonianxianText = wx.StaticText(parent=self.pigeonhole_placeOnFile_panel,label=u"请选择取消年限: ",style=wx.ALIGN_CENTER)
        font_guidangnianxianText = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        guidangnianxianText.SetFont(font_guidangnianxianText)
        quxiaonianxianText.SetFont(font_guidangnianxianText)
        self.guidangnianxianCombo = wx.combo.OwnerDrawnComboBox(parent=self.pigeonhole_placeOnFile_panel,id=-1,size=(100,28),choices=[],style=wx.CB_READONLY)
        #取得数据库中已有的未归档的档案的年限形成一个列表
        self.createGuidangList(event='')
        guidangnianxianButton = wx.Button(parent=self.pigeonhole_placeOnFile_panel,label=u'归    档')
        self.quxiaonianxianCombo = wx.combo.OwnerDrawnComboBox(parent=self.pigeonhole_placeOnFile_panel,id=-1,size=(100,28),choices=[],style=wx.CB_READONLY)
        self.createQuxiaoguidangList(event='')
        quxiaonianxianButton = wx.Button(parent=self.pigeonhole_placeOnFile_panel,label=u'取消归档')
        #绑定归档
        guidangnianxianButton.Bind(wx.EVT_BUTTON, self.guidang)
        #绑定取消归档
        quxiaonianxianButton.Bind(wx.EVT_BUTTON, self.quxiaoguidang)
        quxiao_HBox.Add(quxiaonianxianText,flag=wx.ALIGN_CENTER)
        quxiao_HBox.Add(self.quxiaonianxianCombo,flag=wx.ALIGN_CENTER)
        quxiao_HBox.Add(quxiaonianxianButton,flag=wx.ALIGN_CENTER|wx.LEFT,border=20)
        pigeonhole_HBox.Add(guidangnianxianText,flag=wx.ALIGN_CENTER)
        pigeonhole_HBox.Add(self.guidangnianxianCombo,flag=wx.ALIGN_CENTER)
        pigeonhole_HBox.Add(guidangnianxianButton,flag = wx.ALIGN_CENTER|wx.LEFT,border=20)
        pigeonhole_VBox.Add(pigeonhole_HBox,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        pigeonhole_VBox.Add(quxiao_HBox,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        self.pigeonhole_placeOnFile_panel.SetSizer(pigeonhole_VBox)
        #预销毁页
        #self.preDelete_panel = wx.Panel(parent=self.placeOnFile_Notebook)
        #销毁页
        self.delete_panel = wx.Panel(parent=self.placeOnFile_Notebook)
        destroy_HBox = wx.BoxSizer(wx.HORIZONTAL)
        destroy_VBox = wx.BoxSizer(wx.VERTICAL)
        destorynianxianText = wx.StaticText(parent=self.delete_panel,label=u"请选择销毁年限: ",style=wx.ALIGN_CENTER)
        font_guidangnianxianText = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        destorynianxianText.SetFont(font_guidangnianxianText)
        self.destorynianxianCombo = wx.combo.OwnerDrawnComboBox(parent=self.delete_panel,id=-1,size=(100,28),choices=[],style=wx.CB_READONLY)
        #取得数据库中已有的已归档的档案的年限形成一个列表
        self.createCanDestroyList(event='')
        destorynianxianButton = wx.Button(parent=self.delete_panel,label=u'销    毁')
        #绑定销毁
        destorynianxianButton.Bind(wx.EVT_BUTTON, self.destory)
        destroy_HBox.Add(destorynianxianText,flag=wx.ALIGN_CENTER)
        destroy_HBox.Add(self.destorynianxianCombo,flag=wx.ALIGN_CENTER)
        destroy_HBox.Add(destorynianxianButton,flag = wx.ALIGN_CENTER|wx.LEFT,border=20)
        destroy_VBox.Add(destroy_HBox,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        self.delete_panel.SetSizer(destroy_VBox)  
        #向档案归档框架中添加归档页、预销毁页、销毁页
        self.placeOnFile_Notebook.AddPage(self.pigeonhole_placeOnFile_panel,u"归档、取消归档")
        #self.placeOnFile_Notebook.AddPage(self.preDelete_panel,u"预销毁")
        self.placeOnFile_Notebook.AddPage(self.delete_panel,u"销毁")
        #向主框架中添加档案归档页
        self.main_Notebook.AddPage(self.placeOnFile_panel,u"档案归档、销毁")
        #**********档案归档页结束**********
        #**********系统页开始**********
        self.system_panel = wx.Panel(parent=self.main_Notebook)
        #系统管理框架
        self.system_Notebook = wx.Notebook(parent=self.system_panel,size=SUB_FRAME_SIZE)
        #增加用户页
        self.addUser_system_panel = wx.Panel(parent = self.system_Notebook)
        #生成gridBagSizer布局管理器
        gridBagSizer_addUser = wx.GridBagSizer(hgap=15,vgap=15)
        #生成静态文本（原始密码，新的密码，再次输入）
        usernameText = wx.StaticText(parent=self.addUser_system_panel,label=u"新用户名:",style=wx.ALIGN_CENTER)
        passwordText = wx.StaticText(parent=self.addUser_system_panel,label=u"新的密码:",style=wx.ALIGN_CENTER)
        rePasswordText = wx.StaticText(parent=self.addUser_system_panel,label=u"再次输入:",style=wx.ALIGN_CENTER)
        font_password = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        usernameText.SetFont(font_password)
        passwordText.SetFont(font_password)
        rePasswordText.SetFont(font_password)
        #生成输入控件
        self.usernameCtrl = wx.TextCtrl(parent=self.addUser_system_panel)
        self.passwordCtrl = wx.TextCtrl(parent=self.addUser_system_panel,style=wx.TE_PASSWORD)
        self.rePasswordCtrl = wx.TextCtrl(parent=self.addUser_system_panel,style=wx.TE_PASSWORD)
        addUserButton = wx.Button(parent=self.addUser_system_panel,label=u"添加用户")
        #对添加用户按钮绑定changePassword方法
        addUserButton.Bind(wx.EVT_BUTTON, self.addUser)
        #将静态文本和输入控件加入布局管理器
        gridBagSizer_addUser.Add(usernameText,pos=(0,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=80)
        gridBagSizer_addUser.Add(self.usernameCtrl,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT|wx.TOP,border=80)
        gridBagSizer_addUser.Add(passwordText,pos=(1,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_addUser.Add(self.passwordCtrl,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_addUser.Add(rePasswordText,pos=(2,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_addUser.Add(self.rePasswordCtrl,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_addUser.Add(addUserButton,pos=(3,0),span=(1,4),flag=wx.ALIGN_CENTER)
        #第0列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_addUser.AddGrowableCol(0)
        #第2列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_addUser.AddGrowableCol(2)
        #将gridBagSizer设置为modify_system_panel的布局管理器
        self.addUser_system_panel.SetSizer(gridBagSizer_addUser) 
        
        #修改密码页
        self.modify_system_panel = wx.Panel(parent=self.system_Notebook)
        #生成gridBagSizer布局管理器
        gridBagSizer_modifyPass = wx.GridBagSizer(hgap=15,vgap=15)
        #生成静态文本（原始密码，新的密码，再次输入）
        beforePassText = wx.StaticText(parent=self.modify_system_panel,label=u"原始密码:",style=wx.ALIGN_CENTER)
        nowPassText = wx.StaticText(parent=self.modify_system_panel,label=u"新的密码:",style=wx.ALIGN_CENTER)
        rePassText = wx.StaticText(parent=self.modify_system_panel,label=u"再次输入:",style=wx.ALIGN_CENTER)
        font_pass = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        beforePassText.SetFont(font_pass)
        nowPassText.SetFont(font_pass)
        rePassText.SetFont(font_pass)
        #生成输入控件
        self.beforPassCtrl = wx.TextCtrl(parent=self.modify_system_panel,style=wx.TE_PASSWORD)
        self.nowPassCtrl = wx.TextCtrl(parent=self.modify_system_panel,style=wx.TE_PASSWORD)
        self.rePassCtrl = wx.TextCtrl(parent=self.modify_system_panel,style=wx.TE_PASSWORD)
        confirmButton = wx.Button(parent=self.modify_system_panel,label=u"确定修改")
        #对确定修改按钮绑定changePassword方法
        confirmButton.Bind(wx.EVT_BUTTON, self.changePassword)
        #将静态文本和输入控件加入布局管理器
        gridBagSizer_modifyPass.Add(beforePassText,pos=(0,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=80)
        gridBagSizer_modifyPass.Add(self.beforPassCtrl,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT|wx.TOP,border=80)
        gridBagSizer_modifyPass.Add(nowPassText,pos=(1,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_modifyPass.Add(self.nowPassCtrl,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_modifyPass.Add(rePassText,pos=(2,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_modifyPass.Add(self.rePassCtrl,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_modifyPass.Add(confirmButton,pos=(3,0),span=(1,4),flag=wx.ALIGN_CENTER)
        #第0列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_modifyPass.AddGrowableCol(0)
        #第2列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_modifyPass.AddGrowableCol(2)
        #将gridBagSizer设置为modify_system_panel的布局管理器
        self.modify_system_panel.SetSizer(gridBagSizer_modifyPass)
        
        #重置用户页
        self.reset_system_panel = wx.Panel(parent=self.system_Notebook)
        #生成gridBagSizer布局管理器
        gridBagSizer_reset = wx.GridBagSizer(hgap=15,vgap=15)
        #生成静态文本（原始密码，新的密码，再次输入）
        beforePassText_reset = wx.StaticText(parent=self.reset_system_panel,label=u"用  户  名:",style=wx.ALIGN_CENTER)
        nowPassText_reset = wx.StaticText(parent=self.reset_system_panel,label=u"新的密码:",style=wx.ALIGN_CENTER)
        rePassText_reset = wx.StaticText(parent=self.reset_system_panel,label=u"再次输入:",style=wx.ALIGN_CENTER)
        font_pass_reset = wx.Font(pointSize=15,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        beforePassText_reset.SetFont(font_pass_reset)
        nowPassText_reset.SetFont(font_pass_reset)
        rePassText_reset.SetFont(font_pass_reset)
        #生成输入控件
        user_list = DoCURD.query_user_for_resetUser(self.CONN)
        user_list = [i[0] for i in user_list]
        self.beforPassComb_reset = wx.combo.OwnerDrawnComboBox(parent=self.reset_system_panel,id=-1,size=(112,25),choices=user_list,style=wx.CB_READONLY)
        self.nowPassCtrl_reset = wx.TextCtrl(parent=self.reset_system_panel,style=wx.TE_PASSWORD)
        self.rePassCtrl_reset = wx.TextCtrl(parent=self.reset_system_panel,style=wx.TE_PASSWORD)
        confirmButton = wx.Button(parent=self.reset_system_panel,label=u"确定重置")
        #对确定修改按钮绑定changePassword方法
        confirmButton.Bind(wx.EVT_BUTTON, self.resetUser)
        #将静态文本和输入控件加入布局管理器
        gridBagSizer_reset.Add(beforePassText_reset,pos=(0,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=80)
        gridBagSizer_reset.Add(self.beforPassComb_reset,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT|wx.TOP,border=80)
        gridBagSizer_reset.Add(nowPassText_reset,pos=(1,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_reset.Add(self.nowPassCtrl_reset,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_reset.Add(rePassText_reset,pos=(2,0),span=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        gridBagSizer_reset.Add(self.rePassCtrl_reset,pos=(2,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_LEFT)
        gridBagSizer_reset.Add(confirmButton,pos=(3,0),span=(1,4),flag=wx.ALIGN_CENTER)
        #第0列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_reset.AddGrowableCol(0)
        #第2列可扩展(让该列尽可能占取水平方向上最大的空间)
        gridBagSizer_reset.AddGrowableCol(2)
        #将gridBagSizer设置为modify_system_panel的布局管理器
        self.reset_system_panel.SetSizer(gridBagSizer_reset)        
        
        #数据备份、恢复页
        self.outputAndImport_system_panel = wx.Panel(parent=self.system_Notebook)
        outputAndImport_VBox = wx.BoxSizer(wx.VERTICAL)
        outputButton = wx.Button(parent=self.outputAndImport_system_panel,label=u"数据备份",size=(120,50))
        importButton = wx.Button(parent=self.outputAndImport_system_panel,label=u"数据恢复",size=(120,50)) 
        #设置按钮字体
        font_button = wx.Font(pointSize=18,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        outputButton.SetFont(font_button)
        importButton.SetFont(font_button)
        outputAndImport_VBox.Add(outputButton,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        outputAndImport_VBox.Add(importButton,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
        self.outputAndImport_system_panel.SetSizer(outputAndImport_VBox)          
        #给按钮加入事件
        outputButton.Bind(wx.EVT_BUTTON, self.outputData)
        importButton.Bind(wx.EVT_BUTTON, self.importData)
        
        #数据上报页
        self.shangbao_system_panel = wx.Panel(parent=self.system_Notebook)
        shangbao_VBox = wx.BoxSizer(wx.VERTICAL)
        shangbaoButton = wx.Button(parent=self.shangbao_system_panel,label=u"数据上报",size=(120,50))
        #设置按钮字体
        font_button = wx.Font(pointSize=18,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        shangbaoButton.SetFont(font_button)
        shangbao_VBox.Add(shangbaoButton,flag=wx.ALIGN_CENTER|wx.TOP,border=150)
        self.shangbao_system_panel.SetSizer(shangbao_VBox)          
        #给按钮加入事件
        shangbaoButton.Bind(wx.EVT_BUTTON, self.shangbao)        
       
        #向系统管理框架中添加修改密码页、退出页
        self.system_Notebook.AddPage(self.addUser_system_panel,u"添加用户")
        self.system_Notebook.AddPage(self.modify_system_panel,u"修改密码")
        #if "admin" == self.USERNAME :
        self.system_Notebook.AddPage(self.reset_system_panel,u"重置用户")
        self.system_Notebook.AddPage(self.outputAndImport_system_panel,u"数据备份、数据恢复")
        self.system_Notebook.AddPage(self.shangbao_system_panel,u'数据上报')
        #向主框架中添加系统页
        self.main_Notebook.AddPage(self.system_panel,u"系统")
        #**********系统页结束**********
        
        self.mainBoxSizer.Add(self.main_Notebook)
        #为面板设置主布局管理器
        self.panel.SetSizer(self.mainBoxSizer)
        #生成状态栏
        self.statusBar = self.CreateStatusBar() 
        #显示frame
        self.Show()
    
#$$$$$$$$$$被绑定的事件区结束$$$$$$$$$$ 
    def shangbao(self,event):
        ilist = shell.SHGetSpecialFolderLocation(0, shellcon.CSIDL_DESKTOP)
        deskpath = shell.SHGetPathFromIDList(ilist) 
        dw = u''
        if os.path.exists(u'./re.py') :
            with open(u'./re.py') as r :
                r.readline().strip()
                r.readline().strip()
                r.readline().strip()
                dw = r.readline().strip()
        targetpath = unicode(deskpath)+u"\\"+dw[3:]+u"_"+unicode(datetime.datetime.now().strftime('%Y%m%d'))+u".db"
        try:
            shutil.copyfile(os.getcwd()+u"\\tbdata.db", targetpath)
            wx.MessageBox(u"上报数据已在桌面生成完毕！",u"提示")
        except Exception,e:
            print u"上报数据生成失败:",e

    def showFiles(self,event):
        active_row_number = self.select_grid_archives.GetGridCursorRow()
        #获得指定单元格的值GetCellValue(row,col)
        ShowFiles.CONN = self.CONN
        ShowFiles.ARCHVIEID = self.select_grid_archives.GetCellValue(active_row_number,0)
        if u''!=ShowFiles.ARCHVIEID:
            ShowFiles.ShowFilesPanel(self)

    def showAnjuantiming(self,event):
        quhao = self.quhao_files_comboBox.GetValue()
        guihao = self.guihao_files_comboBox.GetValue()
        hehao = self.hehao_files_comboBox.GetValue()
        juanhao = self.juanhao_files_comboBox.GetValue()
        tempList = DoCURD.query_anjuantiming_for_fileadd(self.CONN,{'quhao':quhao,'guihao':guihao,'hehao':hehao,'juanhao':juanhao})
        self.add_files_anjuantimingText.SetLabel(tempList[0][0])
    def resetUser(self,event):
        state = DoCURD.update_password(self.CONN,{'password':self.nowPassCtrl_reset.Value.strip(),'username':self.beforPassComb_reset.GetValue()})
        if 'ok' == state :
            wx.MessageBox(u"用户 "+self.beforPassComb_reset.GetValue()+" 密码重置成功！",u'提示')
        pass
    def printByGuiHao(self,event):
        quhao = self.quhao_print_comboBox.GetValue().strip()
        guihao = self.guihao_print_comboBox.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        if quhao==u'' or guihao==u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号和柜号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        printList = DoCURD.query_for_guihao_print(self.CONN, paramDict=paramDict)     
        outputPrint.doGuiAnJuanMuLuPrint('',printList=printList,guiOrHe=u"gui")
    def printByHeHao(self,event):
        quhao = self.quhao_print_comboBox_for_he.GetValue().strip()
        guihao = self.guihao_print_comboBox_for_he.GetValue().strip()
        hehao = self.hehao_print_comboBox_for_he.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        paramDict['hehao'] = hehao
        if quhao==u'' or guihao==u'' or hehao==u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号、柜号、盒号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        printList = DoCURD.query_for_hehao_print(self.CONN, paramDict=paramDict)     
        outputPrint.doGuiAnJuanMuLuPrint('',printList=printList,guiOrHe=u"he") 
        
    def printByJuanHao(self,event):
        quhao = self.quhao_print_comboBox_for_juan.GetValue().strip()
        guihao = self.guihao_print_comboBox_for_juan.GetValue().strip()
        hehao = self.hehao_print_comboBox_for_juan.GetValue().strip()
        juanhao = self.juanhao_print_comboBox_for_juan.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        paramDict['hehao'] = hehao
        paramDict['juanhao'] = juanhao
        if quhao==u'' or guihao==u'' or hehao==u'' or juanhao == u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号、柜号、盒号、卷号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        printList = DoCURD.query_files_for_juan_print(self.CONN, paramDict=paramDict)     
        outputPrint.doJuanMuLuPrint('',printList=printList)         
    def printBeikao(self,event):
        quhao = self.quhao_print_comboBox_for_beikao.GetValue().strip()
        guihao = self.guihao_print_comboBox_for_beikao.GetValue().strip()
        hehao = self.hehao_print_comboBox_for_beikao.GetValue().strip()
        juanhao = self.juanhao_print_comboBox_for_beikao.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        paramDict['hehao'] = hehao
        paramDict['juanhao'] = juanhao
        if quhao==u'' or guihao==u'' or hehao==u'' or juanhao == u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号、柜号、盒号、卷号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        
        printList = DoCURD.query_beikaobiao_for_print(self.CONN, paramDict=paramDict)  
        outputPrint.doBeiKaoBiaoPrint('',printList=printList)        
        pass
    def printByAnJuanBiaoTi(self,event):
        quhao = self.quhao_print_comboBox_for_anjuanbiaoti.GetValue().strip()
        guihao = self.guihao_print_comboBox_for_anjuanbiaoti.GetValue().strip()
        hehao = self.hehao_print_comboBox_for_anjuanbiaoti.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        paramDict['hehao'] = hehao
        if quhao==u'' or guihao==u'' or hehao==u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号、柜号、盒号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        printList = DoCURD.query_anjuanbiaoti_print(self.CONN, paramDict=paramDict)     
        outputPrint.doAnJuanBiaoTiPrint('',printList=printList)        
    
    def printCaiWu(self,event):
        quhao = self.quhao_print_comboBox_for_caiwu.GetValue().strip()
        guihao = self.guihao_print_comboBox_for_caiwu.GetValue().strip()
        hehao = self.hehao_print_comboBox_for_caiwu.GetValue().strip()
        paramDict = {}
        paramDict['quhao'] = quhao
        paramDict['guihao'] = guihao
        paramDict['hehao'] = hehao
        if quhao==u'' or guihao==u'' or hehao==u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"请选择区号、柜号、盒号后再打印!",caption=u"打印操作异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #需要打印的列表
        printList = []
        #财务数据同案卷标题数据
        printList = DoCURD.query_anjuanbiaoti_print(self.CONN, paramDict=paramDict)     
        outputPrint.doCaiWuPrint('',printList=printList)       
    
    def addUser(self,event):
        #用户名、密码若为空，提示、返回
        if self.usernameCtrl.Value.strip() == u'' or \
           self.passwordCtrl.Value.strip() == u'' or \
           self.rePasswordCtrl.Value.strip() == u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"用户名或密码不能为空!",caption=u"添加用户异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()        
            return
        #如果两次密码输入不一致,提示、返回
        if self.passwordCtrl.Value.strip() != self.rePasswordCtrl.Value.strip() :
            errorMessage = wx.MessageDialog(parent=None,message=u"两次输入的密码不一致,请核实后再次输入...",caption=u"添加用户异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()  
            return
        #添加前先通过是否能查到密码的方式判断该用户是否存在,如果存在则提示、返回
        isIn = DoCURD.query_password_for_modifyPassword(self.CONN,paramDict={"username":self.usernameCtrl.Value.strip()})
        if len(isIn) != 0 :
            errorMessage = wx.MessageDialog(parent=None,message=u"该用户已经存在,禁止添加!",caption=u"添加用户异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()              
            return
        #添加用户
        try:
            DoCURD.add_user(self.CONN, paramDict={"username":self.usernameCtrl.Value.strip(),"password":self.passwordCtrl.Value.strip()})
            inforMessage = wx.MessageDialog(parent=None,message=u"用户 "+self.usernameCtrl.Value.strip()+" 添加成功！",caption=u"添加用户",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()             
        except Exception,e:
            print e.args[0]
    #数据备份       
    def outputData(self,event):
        dlg = wx.FileDialog(
            self, message=u"数据备份 ...", defaultDir=os.getcwd(), 
            defaultFile="", wildcard="*.db", style=wx.SAVE
            )
    
        if dlg.ShowModal() == wx.ID_OK:
            targetPath = dlg.GetPath()
            fromPath = os.getcwd()+"\\tbdata.db"
            if targetPath == fromPath :
                inforMessage = wx.MessageDialog(parent=None,message=u"数据备份目标文件夹不得与数据源文件夹相同！",caption=u"备份异常信息",style=wx.ICON_INFORMATION)
                result = inforMessage.ShowModal()
                inforMessage.Destroy()                
                return
            try:
                #数据备份
                shutil.copyfile(fromPath, targetPath)
                inforMessage = wx.MessageDialog(parent=None,message=u"数据已备份完毕！",caption=u"备份成功",style=wx.ICON_INFORMATION)
                result = inforMessage.ShowModal()
                inforMessage.Destroy()                
            except Exception,e:
                print e.args[0]
        dlg.Destroy()   
    
    #数据导入
    def importData(self,event):
        targetPath = os.getcwd() + "\\tbdata.db"
        dlg = wx.FileDialog(
            self, message=u"请选择需要恢复的数据文件(提示:数据恢复后软件将自动退出)...",
            defaultDir=os.getcwd(), 
            defaultFile="",
            wildcard="*.db",
            style=wx.OPEN | wx.CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            fromPath = dlg.GetPath()
            try:
                #数据导入
                shutil.copyfile(fromPath, targetPath)
                inforMessage = wx.MessageDialog(parent=None,message=u"数据已恢复完毕,软件自动退出后请重新登陆！",caption=u"恢复成功",style=wx.ICON_INFORMATION)
                result = inforMessage.ShowModal()
                inforMessage.Destroy() 
                #自动退出
                app.Exit()
            except Exception,e:
                print u"异常:",e           
    
        dlg.Destroy()        

    #销毁
    def destory(self,event):
        #销毁确认是否真的删除,默认选否
        questionMessage = wx.MessageDialog(parent=None,message=u"是否销毁所选年限下的档案及其下的所有文件？\n(销毁操作不包括永久类型的档案)",caption=u"档案销毁",style=wx.YES_NO|wx.NO_DEFAULT|wx.ICON_QUESTION)
        result = questionMessage.ShowModal()    
        #如果点击确认按钮则删除档案及其下文件
        if result != wx.ID_YES:
            questionMessage.Destroy()
            return
        questionMessage.Destroy() 
        #提示选择销毁年限
        if self.destorynianxianCombo.GetValue().strip() == "" :
            inforMessage = wx.MessageDialog(parent=None,message=u"请选择需要销毁档案的年限！",caption=u"销毁档案提示",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()   
            return
        #要销毁的年份、销毁人、销毁日期
        year = self.destorynianxianCombo.GetValue().strip()
        people = self.USERNAME
        temp_today = datetime.datetime.now().strftime('%Y%m%d')
        today = unicode(temp_today[:4]+"年"+temp_today[4:6]+"月"+temp_today[6:]+"日")
        paramDict_destory = {'target_table':'destroyedArchives',
                             'people':people,
                             'date':today,
                             'from_table':'archives',
                             'qixian':UtilData.QIXIAN_List[0],#永久
                             'year':year}
        #查找要销毁的列表
        destroyArchivesTurpleList = DoCURD.query_archiveID_for_destory(self.CONN,paramDict=paramDict_destory)
        #如果要销毁的列表为空则返回
        if len(destroyArchivesTurpleList) == 0 :
            return         
        #复制档案表中需要销毁档案到已销毁档案表中
        DoCURD.insert_values_into_destroyedArchives(self.CONN,paramDict=paramDict_destory)
        #更新目标表和原始表的名称
        paramDict_destory['target_table'] = 'destroyedFiles'
        paramDict_destory['from_table'] = 'files'
        #复制文件表中需要销毁的文件到已销毁文件表中
        DoCURD.insert_values_into_destroyedFiles(self.CONN,paramDict=paramDict_destory)
        destroyArchivesList = []
        #迭代元组列表形成要销毁档案的列表
        for row in destroyArchivesTurpleList:
            destroyArchivesList.append(row[0])
        try:
            #销毁物理文件
            for archiveID in destroyArchivesList:
                fileID_filepath_list = DoCURD.query_fileID_filepath_from_files(self.CONN,{"table_name":"files","archiveID":archiveID})
                #如果不是空档案(档案下有文件)
                if len(fileID_filepath_list) != 0 :
                    for row in fileID_filepath_list:
                        #如果filepath值不为无，也即是有电子附件
                        if row[1] != u"无" :
                            #查看该电子附件是否还被别的文件引用
                            fileID_list = DoCURD.query_filepath_from_files(self.CONN, {"table_name":"files","filepath":row[1]})
                            #如果没被别的文件引用(即只被当前文件引用1次)则删除
                            if len(fileID_list) == 1:
                                os.remove(row[1])     
                                #删除空目录
                                if os.path.isdir(os.path.dirname(row[1])) and len(os.listdir(os.path.dirname(row[1])))==0:
                                    os.rmdir(os.path.dirname(row[1]))
                    #删除文件表中该档案下的文件(数据库)
                    DoCURD.delete_row_from_filetable_by_ArchiveID(self.CONN,paramDict={'table_name':'files','archiveID':archiveID})           
                #删除档案表的档案
                DoCURD.delete_row_from_archivetable(self.CONN,paramDict={'table_name':'archives','archiveID':archiveID}) 
            #删除后刷新文件页中的查询和操作视图
            self.selectFilesBySomeConditions(self)
            self.selectManageFilesBySomeConditions(self) 
            #删除后刷新档案页中的查询和操作视图
            self.manageArchivesBySomeConditions(self)  
            self.selectArchivesBySomeConditions(self)
            self.createCanDestroyList(event)
            self.createQuxiaoguidangList(event)
            #提示销毁完毕
            inforMessage = wx.MessageDialog(parent=None,message=u"销毁完毕！",caption=u"提示",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()            
        except Exception,e:
            print e.args[0]

    #生成未归档的档案的年限列表
    def createGuidangList(self,event):
        yearRowsList = DoCURD.query_noGuidang_year(self.CONN)
        temp_year_list = []
        for row in yearRowsList:
            temp_year_list.append(row[0])
        self.guidangnianxianCombo.Set(temp_year_list)
    def createQuxiaoguidangList(self,event):
        yearRowsList = DoCURD.query_already_year(self.CONN)
        temp_year_list = []
        for row in yearRowsList:
            temp_year_list.append(row[0])
        self.quxiaonianxianCombo.Set(temp_year_list)
    def createCanDestroyList(self,event):
        yearRowsList = DoCURD.query_already_year(self.CONN)
        temp_year_list = []
        for row in yearRowsList:
            temp_year_list.append(row[0])
        self.destorynianxianCombo.Set(temp_year_list)
    def guidang(self,event):
        if self.guidangnianxianCombo.GetValue().strip() == "" :
            inforMessage = wx.MessageDialog(parent=None,message=u"请选择归档年限！",caption=u"归档提示",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()  
            return
        #取得要归档的年份
        year = self.guidangnianxianCombo.GetValue()
        state = DoCURD.update_guidang(self.CONN, paramDict={'year':year})
        if u"ok"==state :
            #提示归档成功
            inforMessage = wx.MessageDialog(parent=None,message=unicode(year)+u"年档案已归档完毕！",caption=u"归档成功",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()  
            self.createGuidangList(event)
            self.createQuxiaoguidangList(event)
            self.createCanDestroyList(event)
        else:
            #提示归档失败
            inforMessage = wx.MessageDialog(parent=None,message=unicode(year)+u"年档案已归档失败！",caption=u"归档失败",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()
    def quxiaoguidang(self,event):
        if self.quxiaonianxianCombo.GetValue().strip() == "" :
            inforMessage = wx.MessageDialog(parent=None,message=u"请选择取消归档年限！",caption=u"取消归档提示",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()
            return
        #取得要取消归档的年份
        year = self.quxiaonianxianCombo.GetValue()
        state = DoCURD.update_quxiao_guidang(self.CONN,paramDict={'year':year})
        if u"ok"==state :
            #提示已取消归档
            inforMessage = wx.MessageDialog(parent=None,message=unicode(year)+u"年档案已取消归档！",caption=u"取消归档成功",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()  
            self.createGuidangList(event)
            self.createQuxiaoguidangList(event)
            self.createCanDestroyList(event)
        else:
            #提示取消归档失败
            inforMessage = wx.MessageDialog(parent=None,message=unicode(year)+u"年档案取消归档失败！",caption=u"取消归档失败",style=wx.ICON_INFORMATION)
            result = inforMessage.ShowModal()
            inforMessage.Destroy()        
       
    def createAchivesIDforFileSelectAndManage(self,event):
        #查询档案表中共有多少档案ID,都是什么
        temp_archiveIDList = DoCURD.query_archiveID_for_selectfile(self.CONN,paramDict={"table_name":"archives"})
        archiveIDs = ['']    
        for row in temp_archiveIDList:
            for col in row :
                archiveIDs.append(unicode(col))
        #更新文件查询和管理中的archiveID
        self.archiveID_select_files_comboBox.Set(archiveIDs)
        self.archiveID_manage_files_comboBox.Set(archiveIDs)
        
    #下拉对话框事件选择
    def choiceEvent(self,event):
        #获得选择的值
        pass
    #按钮关闭处理器
    def closeButton(self,event):
        self.Close(True) 
    #窗体关闭处理器
    def closeWindow(self,event):
        self.Destroy()
    def doMe(self,event):
        #删除欢迎页
        self.main_Notebook.DeletePage(0)
    #button1的事件按钮    
    def doChange(self,event):
        pass
    #退出页面事件
    def closeBook(self,event):
        self.Close(True)
        self.Destroy()
    #修改密码
    def changePassword(self,event):
        #用户名、密码若为空，提示、返回
        if self.beforPassCtrl.Value.strip() == u'' or \
           self.nowPassCtrl.Value.strip() == u'' or \
           self.rePassCtrl.Value.strip() == u'' :
            errorMessage = wx.MessageDialog(parent=None,message=u"密码内容不能为空!",caption=u"修改密码异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()        
            return    
        #如果两次密码输入不一致，则提示、返回
        if self.nowPassCtrl.Value.strip() != self.rePassCtrl.Value.strip() :   
            errorMessage = wx.MessageDialog(parent=None,message=u"新密码两次输入不一致,请核实后重新输入",caption=u"密码修改异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()  
            return
        #根据用户名找到原密码
        temp_passwordTurpleList = DoCURD.query_password_for_modifyPassword(self.CONN,paramDict={"username":self.USERNAME})
        #如果输入的密码为原密码则可设置新密码
        if temp_passwordTurpleList[0][0]==self.beforPassCtrl.Value.strip() :
            state=DoCURD.update_password(self.CONN, paramDict={"username":self.USERNAME,"password":self.nowPassCtrl.Value.strip()})
            if u'ok' == state :
                inforMessage = wx.MessageDialog(parent=None,message=u"用户 "+self.USERNAME+u" 密码修改成功！",caption=u"密码修改",style=wx.ICON_INFORMATION)
                result = inforMessage.ShowModal()
                inforMessage.Destroy()                   
        #否则，提示原密码错误
        else:
            errorMessage = wx.MessageDialog(parent=None,message=u"原始密码错误,请核实后重新输入",caption=u"密码修改异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
    
    #登陆校验
    def checkLogin(self,event):
        #查询给定的用户名对应的密码
        temp_passwordList = DoCURD.query_password_for_modifyPassword(self.CONN,paramDict={"username":self.usernameText_welcome.Value.strip()})
        #如果返回列表长度为0,则证明该用户不存在
        if len(temp_passwordList) == 0 :
            errorMessage = wx.MessageDialog(parent=None,message=u"该用户不存在!",caption=u"登陆异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return 
        #用户名密码通过则解冻各页面
        if temp_passwordList[0][0]==self.passwordText_welcome.Value.strip() :
            for i in range(1,int(self.main_Notebook.GetPageCount())):
                self.main_Notebook.GetPage(i).Thaw()
            #设置新增页输入人的默认值
            self.USERNAME = self.usernameText_welcome.Value.strip()
            self.inputter_archives_ctrl.SetValue(self.USERNAME)
            self.inputter_files_ctrl.SetValue(self.USERNAME)
            self.Title = u"天保档案管理系统  v1.0"+u"  [当前用户: " + self.USERNAME + u']'
            #删除登陆页
            self.main_Notebook.DeletePage(0)
            if 'admin'!=self.USERNAME:
                self.system_Notebook.DeletePage(0)
                self.system_Notebook.DeletePage(1)
        else:
            errorMessage = wx.MessageDialog(parent=None,message=u"用户名或密码错误,请核实后重新登陆",caption=u"登陆异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()
          
    #根据门类的选择动态提供归档类型
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

            
    #根据门类的选择动态提供归档类型
    def choiceGuiDang(self,event):
        TEMP = self.menlei_archives_comboBox.GetValue().split(u"_")[:1]
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
        #self.guidang_archives_comboBox.Clear()
        #根据门类不同的选择重新设置comboBox的choice值
        self.guidang_archives_comboBox.Set(GUIDANG_MENLEI_List)
        #默认第0个元素选中状态
        self.guidang_archives_comboBox.SetSelection(0)
        
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
        #self.guidang_archives_comboBox.Clear()
        #根据门类不同的选择重新设置comboBox的choice值
        self.modify_guidang_archives_comboBox.Set(GUIDANG_MENLEI_List)
        #默认第0个元素选中状态
        self.modify_guidang_archives_comboBox.SetSelection(0)
        
    def addNewArchive(self,event):
        #判断案卷题名、立卷日期、责任人、输入人是否为空
        if self.anjuantiming_archives_ctrl.Value.strip() == "" or self.lijuanriqi_archives_pop.GetValue().strip() == "" or self.zerenren_archives_ctrl.Value.strip() == "" or self.inputter_archives_ctrl.Value.strip() == "" :
            self.statusBar.SetStatusText(unicode("案卷题名、立卷日期、责任人、录入人不允许为空！"))
            errorMessage = wx.MessageDialog(parent=None,message=u"案卷题名、立卷日期、责任人、录入人不允许为空！",caption=u"新增档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return
        #为了判断日期是否准确临时生成一个temp_cal变量
        temp_cal = self.lijuanriqi_archives_pop.GetValue().strip().decode('utf-8')
        if len(temp_cal) != 11 or temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日" :
            self.statusBar.SetStatusText(unicode("立卷日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！"))
            errorMessage = wx.MessageDialog(parent=None,message=u"立卷日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"新增档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return
        #判断年份是否已经归档，如果年份已经归档则不允许新增
        year_already_RowsList = DoCURD.query_already_year(self.CONN)
        year_already_List = []
        for row in year_already_RowsList :
            year_already_List.append(row[0])
        if temp_cal[:4] in year_already_List :
            self.statusBar.SetStatusText(unicode(temp_cal[:4])+u"年档案已经归档,不允许新增!")
            errorMessage = wx.MessageDialog(parent=None,message=unicode(temp_cal[:4])+u"年档案已经归档,不允许新增!",caption=u"新增档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()             
            return
        #将用户输入的数据处理后插入数据库中
        paramDict = {}
        paramDict["table_name"]="archives"
        paramDict["archiveID"]="Null"
        paramDict['menlei']=unicode(self.menlei_archives_comboBox.GetValue())
        paramDict['guidang']=unicode(self.guidang_archives_comboBox.GetValue())
        paramDict['qixian']=unicode(self.qixian_archives_comboBox.GetValue())
        paramDict['anjuantiming']=unicode(self.anjuantiming_archives_ctrl.Value)
        paramDict['danwei']=unicode(self.danwei_archives_comboBox.GetValue())
        paramDict['lijuanriqi']=unicode(self.lijuanriqi_archives_pop.GetValue())
        paramDict['weizhi']=unicode(self.weizhi_archives_comboBox.GetValue())
        paramDict['miji']=unicode(self.miji_archives_comboBox.GetValue())
        paramDict['zerenren']=unicode(self.zerenren_archives_ctrl.Value)
        paramDict['quhao']=unicode(self.quhao_archives_comboBox.GetValue())
        paramDict['guihao']=unicode(self.guihao_archives_comboBox.GetValue())
        paramDict['hehao']=unicode(self.hehao_archives_comboBox.GetValue())
        paramDict['juanhao']=unicode(self.juanhao_archives_comboBox.GetValue())
        paramDict['hujianhao']=unicode(self.hujianhao_archives_ctrl.Value)
        paramDict['kemu']=unicode(self.kemu_archives_comboBox.GetValue())
        paramDict['beizhu']=unicode(self.beizhu_archives_ctrl.Value)
        paramDict['inputter']=unicode(self.inputter_archives_ctrl.Value)
        #paramDict['anjuanhao']=unicode(self.anjuanhao_archives_comboBox.GetValue())

        #查询给定区号、柜号、盒号下的卷号列表
        temp_juanhao_list = DoCURD.query_juanhao_values(self.CONN,paramDict=paramDict)
        temp_juanhaoExist_set = set()
        for rows in temp_juanhao_list:
            for col in rows :
                temp_juanhaoExist_set.add(col)
        #如果要新增的卷号在给定的区号、柜号、盒号中已经存在,则禁止插入到数据库中,返回
        if paramDict.get("juanhao") in temp_juanhaoExist_set:
            self.statusBar.SetStatusText(unicode(paramDict.get("quhao") + " 区 " + paramDict.get("guihao") + " 柜 " + paramDict.get("hehao") + " 盒中 " + paramDict.get("juanhao")+" 卷已存在！"))
            errorMessage = wx.MessageDialog(parent=None,message=unicode(paramDict.get("quhao") + " 区 " + paramDict.get("guihao") + " 柜 " + paramDict.get("hehao") + " 盒中 " + paramDict.get("juanhao")+" 卷已存在！"),caption=u"新增档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return 
        #通过了以上的验证则将数据插入到数据库中
        state = DoCURD.insert_values(self.CONN,paramDict=paramDict)
        self.statusBar.SetStatusText(unicode(state)) 
        #更新档案查询页的数据
        #self.query_all_archives()
        #新增档案后自动更新卷号等
        self.createChoicesForJuanHao(event='')    
        self.select_grid_archives.Refresh()
        self.createjuanhaoByDanwei(event='')
        self.createAchivesIDforFileSelectAndManage(event='')
        self.createGuidangList(event='')
        
    def createChoicesForJuanHao(self,event):
        #查询给定区号、柜号、盒号下的卷号列表
        temp_juanhao_list = DoCURD.query_juanhao_values(self.CONN,paramDict={"table_name":"archives",
                                                                         "quhao":self.quhao_archives_comboBox.GetValue() ,
                                                                         "guihao":self.guihao_archives_comboBox.GetValue() ,
                                                                         "hehao":self.hehao_archives_comboBox.GetValue()})

        temp_juanhaoExist_list = []
        for rows in temp_juanhao_list:
            temp_juanhaoExist_list.append(rows[0]) 
        #如果元素在全列表中而不在已存在列表中则将该元素返回
        temp_choices_list = [i for i in UtilData.JUANHAO_List if i not in temp_juanhaoExist_list]
        #设置卷号控件中数据源为新生成的temp_choices_list
        self.juanhao_archives_comboBox.Set(temp_choices_list)
        #设置第0个元素为默认显示
        self.juanhao_archives_comboBox.SetSelection(0)
        
    #def createChoicesForAnJuanHao(self,event):
        ##查询给定区号、柜号、盒号下的卷号列表
        #temp_juanhao_list = DoCURD.query_juanhao_values(self.CONN,paramDict={"table_name":"archives",
                                                                         #"menlei":self.menlei_archives_comboBox.GetValue() ,
                                                                         #"guidang":self.guidang_archives_comboBox.GetValue() ,
                                                                         #"danwei":self.danwei_archives_comboBox.GetValue()})
        #temp_juanhaoExist_list = []
        #for rows in temp_juanhao_list:
            #temp_juanhaoExist_list.append(rows[0]) 
        ##如果元素在全列表中而不在已存在列表中则将该元素返回
        #temp_juanhao_choices_list = [i for i in UtilData.JUANHAO_List if i not in temp_juanhaoExist_list]
        ##设置卷号控件中数据源为新生成的temp_choices_list
        #self.juanhao_archives_comboBox.Set(temp_anjuanhao_choices_list)
        ##设置第0个元素为默认显示
        #self.juanhao_archives_comboBox.SetSelection(0)    
    
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
        self.statusBar.SetStatusText(u"查询到档案 " + unicode(temp_num) + u" 个")
    #档案管理页的查询按钮
    def manageArchivesBySomeConditions(self,event):
        paramDict = {"table_name":"archives"}
        paramDict['quhao'] = unicode(self.quhao_manage_archives_comboBox.GetValue().strip())
        paramDict['guihao'] = unicode(self.guihao_manage_archives_comboBox.GetValue().strip())
        paramDict['hehao'] = unicode(self.hehao_manage_archives_comboBox.GetValue().strip())
        paramDict['juanhao'] = unicode(self.juanhao_manage_archives_comboBox.GetValue().strip())
        paramDict['anjuantiming'] = unicode(self.anjuantiming_manage_archives_Ctrl.Value.strip())
        paramDict['menlei'] = u""
        paramDict['guidang'] = u""
        paramDict['qixian'] = u""
        paramDict['danwei'] = u""
        paramDict['lijuanriqi'] = u""
        paramDict['inputter'] = u""
        paramDict['hujianhao'] = u""
        paramDict['weizhi'] = u""
        paramDict['kemu'] = u""
        paramDict['miji'] = u""
        #paramDict['archiveID'] = u''
        #查询数据库中给定条件的档案,结果替换原有数据源
        self.manage_data_table.dataList = DoCURD.query_some_values_from_archivetable(self.CONN,paramDict=paramDict)
        #整理查询到的二维表，将其中的行由元组转换为列表
        #tempList = []
        #for row in self.manage_data_table.dataList:
            #tempList.append(list(row))
        #self.manage_data_table.dataList = tempList
        #临时存放现有表格行数
        temp_num = len(self.manage_data_table.dataList)
        needAddRows = 20 - temp_num
        #用20减去临时存放表格的行数为需要添加的行数
        #需要添加的行数如果为正值证明记录数少于20,少几行补几行,再删除上次表中多余行数(即上次表行数减20后的数)
        if needAddRows > 0 :
            for i in range(needAddRows) :
                self.manage_data_table.dataList.append(['','','','','','','','','','','','','','','','','','',''])
            #如果查询到的数据不足一屏则将原有表格容器多余的行数删除(删除的个数为self.manage_data_table.GetRowsCount()-默认的20个)
            self.manage_data_table.DeleteRows(numRows=(int(self.last_data_manage_table_rows)-20))
            #更新上一次表格行数
            self.last_data_manage_table_rows = self.manage_data_table.GetRowsCount()
        #需要添加的行数如果为负值证明记录数多余20,将现有表的行数和上一次表的行数相比对
        else:
            #如果现在表的行数小于上一次表的行数则删除容器中多余的行
            if int(temp_num) < int(self.last_data_manage_table_rows) :
                # #如果超过一屏20个且上次表格行数超过20个则删（现有表格容器的行数-列表元素个数(查询到的记录数)）行
                self.manage_data_table.DeleteRows(numRows=abs(temp_num-int(self.last_data_manage_table_rows)))
            #如果现在表的行数大于上一次表的行数则给容器增加多出的行
            else:
                 #如果超过一屏20个且上次表格行数不足20个则加（列表元素个数(查询到的记录数)-现有表格容器的行数）行
                self.manage_data_table.AppendRows(abs(temp_num-int(self.last_data_manage_table_rows)))
            #更新上一次表格行数
            self.last_data_manage_table_rows = self.manage_data_table.GetRowsCount()            
        #对查询后的数据进行刷新,否则不显示
        self.manage_grid_archives.Refresh()
        #自动调整行宽列宽
        self.manage_grid_archives.AutoSizeColumns()
        self.manage_grid_archives.AutoSizeRows()        
        self.statusBar.SetStatusText(u"查询到档案 " + unicode(temp_num) + u" 个")
        
        
    #档案修改事件    
    def manageArchives_modify(self,event):
        #获得当前选中单元格所处的行号
        active_row_number = self.manage_grid_archives.GetGridCursorRow()
        #获得当前表中列数
        active_row_cols = self.manage_grid_archives.GetNumberCols()
        
        #获得指定单元格的值GetCellValue(row,col)
        #生成临时存储行信息的列表并将选中的行中的信息复制到temp_list_frame中,temp_list_frame[0]中为archiveID
        temp_list_frame = []
        for col in range(active_row_cols):
            temp_list_frame.append(self.manage_grid_archives.GetCellValue(active_row_number,col))
        #如果空表修改无效
        if temp_list_frame[0].strip() == u"" :
            return   
        #如果档案已经归档则不允许修改
        if temp_list_frame[-1].strip() != u"否" :
            self.statusBar.SetStatusText(u"该档案已经归档,不允许修改！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该档案已经归档,不允许修改！",caption=u"修改档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()                 
            return        
        #生成修改对话框
        DialogDemo.temp_list_frame = temp_list_frame
        DialogDemo.TestPanel(parent=self)  
        
    #档案删除事件
    def manageArchives_delete(self,event):
        #获得当前选中单元格所处的行号
        active_row_number = self.manage_grid_archives.GetGridCursorRow()
        #获得指定单元格的值GetCellValue(row,col)
       
        paramDict={}
        paramDict['table_name'] = 'archives'
        paramDict['archiveID'] = self.manage_grid_archives.GetCellValue(active_row_number,0)
        #如果空表删除无效
        if paramDict['archiveID'].strip() == u"" :
            return 
        #如果档案已经归档则不允许删除
        if self.manage_grid_archives.GetCellValue(active_row_number,18).strip() != u"否" :
            self.statusBar.SetStatusText(u"该档案已经归档,不允许删除！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该档案已经归档,不允许删除！",caption=u"修改档案异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()                 
            return         
        #删除前确认是否真的删除,默认选否
        questionMessage = wx.MessageDialog(parent=None,message=u"是否删除该档案及该档案下所有文件？",caption=u"删除档案",style=wx.YES_NO|wx.NO_DEFAULT|wx.ICON_QUESTION)
        result = questionMessage.ShowModal()    
        #如果点击确认按钮则删除档案及其下文件
        if result == wx.ID_YES:
            #删除物理文件
            fileID_filepath_list = DoCURD.query_fileID_filepath_from_files(self.CONN,{"table_name":"files","archiveID":paramDict.get("archiveID")})
            #如果不是空档案(档案下有文件)
            if len(fileID_filepath_list) != 0 :
                for row in fileID_filepath_list:
                    #如果filepath值不为无，也即是有电子附件
                    if row[1] != u"无" :
                        #查看该电子附件是否还被别的文件引用
                        fileID_list = DoCURD.query_filepath_from_files(self.CONN, {"table_name":"files","filepath":row[1]})
                        #如果没被别的文件引用(即只被当前文件引用1次)则删除
                        if len(fileID_list) == 1:
                            os.remove(row[1])     
                            #删除空目录
                            if os.path.isdir(os.path.dirname(row[1])) and len(os.listdir(os.path.dirname(row[1])))==0:
                                os.rmdir(os.path.dirname(row[1]))
                #删除文件表中该档案下的文件(数据库)
                DoCURD.delete_row_from_filetable_by_ArchiveID(self.CONN,paramDict={'table_name':'files','archiveID':paramDict.get('archiveID')})      
                #删除后刷新文件表
                self.selectManageFilesBySomeConditions(self)  
                self.selectFilesBySomeConditions(self)
            #删除档案表的档案
            DoCURD.delete_row_from_archivetable(self.CONN,paramDict=paramDict) 
            #删除后刷新档案表
            self.manageArchivesBySomeConditions(self)  
            self.selectArchivesBySomeConditions(self) 
        questionMessage.Destroy()       
        
    #选中当前所选单元格所在行且激活当前选中行中所选单元格
    def activeRow(self,event):
        #选中当前所选单元格所在行
        self.manage_grid_archives.SelectRow(row=event.GetRow())
        #激活当前选中行中所选单元格
        self.manage_grid_archives.SetGridCursor(row=event.GetRow(),col=event.GetCol())
    
    #文件新增页 添加文件按钮绑定事件
    def addNewFile(self,event):
        #如果isGuidang值为'是',则表示要加入的档案已经归档，文件不允许新增入档
        isGuidang = DoCURD.query_isGuidang_values_for_addfile(self.CONN,paramDict={"table_name":"archives",
                                                                  "quhao":self.quhao_files_comboBox.GetValue().strip(),
                                                                  "guihao":self.guihao_files_comboBox.GetValue().strip(),
                                                                  "hehao":self.hehao_files_comboBox.GetValue().strip(),
                                                                  "juanhao":self.juanhao_files_comboBox.GetValue().strip()
                                                                  }
                                                  )
        if u'是'==isGuidang[0][0] :
            self.statusBar.SetStatusText(u"不允许往已经归档的档案中添加文件！")
            errorMessage = wx.MessageDialog(parent=None,message=u"不允许往已经归档的档案中添加文件！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return             
        if  self.juanhao_files_comboBox.GetValue().strip() == u"" :
            self.statusBar.SetStatusText(u"该盒号下暂时没有档案,请在该单位下新建档案后再添加文件！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该盒号下暂时没有档案,请在该单位下新建档案后再添加文件！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return 
        #文件题目、文件编号、发文单位不允许为空
        if self.wenjianbianhao_files_ctrl.Value.strip()=="" or self.wenjiantimu_files_ctrl.Value.strip()=="" or self.fawendanwei_files_ctrl.Value.strip()=="":
            self.statusBar.SetStatusText(u"文件题目、文件编号、发文单位不允许为空！")
            errorMessage = wx.MessageDialog(parent=None,message=u"文件题目、文件编号、发文单位不允许为空！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #要输入的文件编号如果已经存在则提示并返回
        temp_archiveID_for_wenjianbianhao = DoCURD.query_archiveID_values(self.CONN, paramDict={"table_name":"archives","quhao":self.quhao_files_comboBox.GetValue(),
                                                                                                                   "guihao":self.guihao_files_comboBox.GetValue(),
                                                                                                                   "hehao":self.hehao_files_comboBox.GetValue(),
                                                                                                                   "juanhao":self.juanhao_files_comboBox.GetValue()})
        temp_wenjianbianhao = DoCURD.query_wenjianbianhao_values(self.CONN,{"table_name":"files","archiveID":temp_archiveID_for_wenjianbianhao[0][0]})
        temp_wenjianbianhao_list = []
        for row in temp_wenjianbianhao :
            temp_wenjianbianhao_list.append(row[0])
        #如果输入的文件编号已经在数据库中则不允许录入
        if self.wenjianbianhao_files_ctrl.Value.strip() in temp_wenjianbianhao_list :
            self.statusBar.SetStatusText(u"该文件编号已经存在,请核实后再次输入！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该文件编号已经存在,请核实后再次输入！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return 
        #为了判断日期是否准确临时生成一个temp_cal变量
        temp_cal = self.xingchengriqi_files_pop.GetValue().strip().decode('utf-8')
        if len(temp_cal)!=11 or temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日" :
            self.statusBar.SetStatusText(u"形成日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！")
            errorMessage = wx.MessageDialog(parent=None,message=u"形成日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return
        #首先查询档案表中已经有的卷号
        paramDict={}
        paramDict['table_name'] = 'archives'
        paramDict['quhao'] = self.quhao_files_comboBox.GetValue().strip()
        paramDict['guihao'] = self.guihao_files_comboBox.GetValue().strip()
        paramDict['hehao'] = self.hehao_files_comboBox.GetValue().strip()
        juanhaoAndarchiveID_exist_list = DoCURD.query_juanhaoArchiveID_values_for_addfile(self.CONN,paramDict=paramDict)
        #取出内容控件中的值作为新文件的值
        newFile = self.neirong_file_FileBrowseButton.GetValue()
        #临时档案ID
        temp_archiveID = -1
        for row in juanhaoAndarchiveID_exist_list:
            for col in row[:1] :
                #如果查询到的第一个元素(已存在卷号)等于要存储的卷号
                if col == self.juanhao_files_comboBox.GetValue().strip() :
                    #则将第二个元素(档案ID)赋值到临时档案ID变量中
                    temp_archiveID = row[-1]      
        #如果临时档案ID值不为-1,也就是文件新增页输入的档号已经存在,则可将文件执行添加操作
        if temp_archiveID != -1 :
            #如果内容控件中是空的,则将目标路径设为无
            if newFile.strip() == "" :
                targetPath = u'无'
                targetFile = u"无"
            #如果内容控件中不是空的,则对内容控件的值进行判断
              #值正确，上传
              #值错误，提示错误信息
            else :
                #判断要存储电子稿的文件夹是否存在(elec+temp_archiveID)
                targetPath = os.getcwd()+"\\elec\\"+str(temp_archiveID)
                #如果存在则不需要新建,如果不存在则需要新建
                if not os.path.exists(targetPath):
                    os.makedirs(targetPath)
                try:
                    targetFile = targetPath+"\\"+os.path.basename(newFile)
                    #复制附件中路径指向的文件到存储电子稿的文件夹
                    if targetFile != newFile :
                        shutil.copyfile(newFile,targetFile)
                except Exception,e:
                    errorMessage = wx.MessageDialog(parent=None,message=u"请输入正确的文件路径或使用上传按钮选择文件路径！",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
                    result = errorMessage.ShowModal()
                    errorMessage.Destroy()
                    return
            #文件存储位置设定
            paramDict.clear()
            paramDict['table_name'] = 'files'
            paramDict['fileID'] = 'NULL'
            paramDict['archiveID'] = temp_archiveID
            paramDict['wenjiantimu'] = unicode(self.wenjianbianhao_files_ctrl.Value.strip())
            paramDict['wenjianbianhao'] = unicode(self.wenjianbianhao_files_ctrl.Value.strip())
            paramDict['fawendanwei'] = unicode(self.fawendanwei_files_ctrl.Value.strip())
            paramDict['xingchengriqi'] = unicode(self.xingchengriqi_files_pop.GetValue().strip())
            paramDict['yeshu'] = unicode(self.yeshu_files_comboBox.GetValue().strip())
            paramDict['miji'] = unicode(self.miji_files_comboBox.GetValue().strip())
            paramDict['beizhu'] = unicode(self.beizhu_files_ctrl.Value.strip())
            paramDict['inputter'] = unicode(self.inputter_files_ctrl.Value.strip())
            paramDict['filepath'] = unicode(targetFile)
            #如果一切OK,则将值插入到文件表中
            state = DoCURD.insert_values_into_files(self.CONN, paramDict=paramDict) 
            #新增文件后刷新文件管理页和查询页
            self.selectFilesBySomeConditions(event='')
            self.selectManageFilesBySomeConditions(evnet='')
            #状态栏提示文件新增成功
            self.statusBar.SetStatusText(unicode(state))
        #如果无，则提示要添加到的档案不存在，请核实后再添加文件
        else:
            errorMessage = wx.MessageDialog(parent=None,message=u"指定的档号索引: "+paramDict.get("区号")+u"-"+paramDict.get("guihao")+u"-"+paramDict.get("hehao")+u"-"+self.juanhao_files_comboBox.GetValue().strip()+u" 不存在,请核实后再添加文件!",caption=u"新增文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()
            return
    
    #文件新增页 门类下拉选择框绑定事件（设置归档下拉框的值）
    def choiceFileGuiDang(self,event):
        TEMP = self.menlei_files_comboBox.GetValue().split(u"_")[:1]
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
        #self.guidang_archives_comboBox.Clear()
        #根据门类不同的选择重新设置comboBox的choice值
        self.guidang_files_comboBox.Set(GUIDANG_MENLEI_List)
        #默认第0个元素选中状态
        self.guidang_files_comboBox.SetSelection(0)
    
    #文件查询页中的查询
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
        self.statusBar.SetStatusText(u"查询到文件 " + unicode(temp_num_file) + u" 个")
    
    #文件管理页中的查询
    def selectManageFilesBySomeConditions(self,evnet):
        paramDict = {"table_name":"files"}
        paramDict['archiveID'] = unicode(self.archiveID_manage_files_comboBox.GetValue())
        paramDict['wenjiantimu'] = unicode(self.wenjiantimu_manage_files_Ctrl.GetValue())
        paramDict['wenjianbianhao'] = unicode(self.wenjianbianhao_manage_files_ctrl.GetValue())
        #查询数据库中给定条件的档案,结果替换原有数据源
        if paramDict['archiveID']!="" :
            self.manage_data_file_table.dataList_File = DoCURD.query_some_values_from_filestable2(self.CONN,paramDict=paramDict)
        else:
            self.manage_data_file_table.dataList_File = DoCURD.query_some_values_from_filestable_noArchiveID2(self.CONN,paramDict=paramDict)
        #临时存放现有表格行数
        temp_num_file = len(self.manage_data_file_table.dataList_File)
        needAddRows_file = 20 - temp_num_file
        #用20减去临时存放表格的行数为需要添加的行数
        #需要添加的行数如果为正值证明记录数少于20,少几行补几行,再删除上次表中多余行数(即上次表行数减20后的数)
        if needAddRows_file > 0 :
            for i in range(needAddRows_file) :
                self.manage_data_file_table.dataList_File.append(['','','','','','','','','','',''])
            #如果查询到的数据不足一屏则将原有表格容器多余的行数删除(删除的个数为self.select_data_table.GetRowsCount()-默认的20个)
            self.manage_data_file_table.DeleteRows(numRows=(int(self.last_data_manage_file_table_rows)-20))
            #更新上一次表格行数
            self.last_data_manage_file_table_rows = self.manage_data_file_table.GetRowsCount()
        #需要添加的行数如果为负值证明记录数多余20,将现有表的行数和上一次表的行数相比对
        else:
            #如果现在表的行数小于上一次表的行数则删除容器中多余的行
            if int(temp_num_file) < int(self.last_data_manage_file_table_rows) :
                # #如果超过一屏20个且上次表格行数超过20个则删（现有表格容器的行数-列表元素个数(查询到的记录数)）行
                self.manage_data_file_table.DeleteRows(numRows=abs(temp_num_file-int(self.last_data_manage_file_table_rows)))
            #如果现在表的行数大于上一次表的行数则给容器增加多出的行
            else:
                 #如果超过一屏20个且上次表格行数不足20个则加（列表元素个数(查询到的记录数)-现有表格容器的行数）行
                self.manage_data_file_table.AppendRows(abs(temp_num_file-int(self.last_data_manage_file_table_rows)))
            #更新上一次表格行数
            self.last_data_manage_file_table_rows = self.manage_data_file_table.GetRowsCount()            
        #对查询后的数据进行刷新,否则不显示
        self.manage_grid_files.Refresh()
        #自动调整行宽列宽
        self.manage_grid_files.AutoSizeColumns()
        self.manage_grid_files.AutoSizeRows()        
        self.statusBar.SetStatusText(u"查询到文件 " + unicode(temp_num_file) + u" 个")
    
    #文件管理页中的修改
    def modifyManageFilesBySomeConditions(self,event):
        #获得当前选中单元格所处的行号
        active_row_number = self.manage_grid_files.GetGridCursorRow()
        #获得当前表中列数
        active_row_cols = self.manage_grid_files.GetNumberCols()
        #获得当前选中文件所在的档案ID
        nowArchiveID = self.manage_grid_files.GetCellValue(active_row_number,1)
        if len(self.manage_grid_files.GetCellValue(active_row_number,0)) == 0 :
            return        
        #查询该档案的状态(是否归档),如归档则提示后返回
        isGuidang = DoCURD.query_isGuidang_values(self.CONN,{"table_name":"archives","archiveID":nowArchiveID.strip()})
        if  len(isGuidang)==0 :
            return
        if  isGuidang[0][0] != u"否" :
            errorMessage = wx.MessageDialog(parent=None,message=u"该文件所在的档案已归档,不允许修改!",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return
        #获得指定单元格的值GetCellValue(row,col)
        #生成临时存储行信息的列表并将选中的行中的信息复制到temp_list_file_frame中,temp_list_file_frame[0]中为fileID
        temp_list_file_frame = []
        for col in range(active_row_cols):
            temp_list_file_frame.append(self.manage_grid_files.GetCellValue(active_row_number,col))
        #如果空表修改无效
        if temp_list_file_frame[0].strip() == u"" :
            return 
        #生成修改对话框
        DialogDemo_File.temp_list_file_frame = temp_list_file_frame
        DialogDemo_File.TestPanel(parent=self)  
        
        #修改后一定要确认一下附件是否还被别的文件引用,如无引用则删除文件
    
    #文件管理页中的删除
    def deleteManageFilesBySomeConditions(self,event):
        #获得当前选中单元格所处的行号
        active_row_number = self.manage_grid_files.GetGridCursorRow()
        #获得当前表中列数
        active_row_cols = self.manage_grid_files.GetNumberCols()
        #获得当前选中文件所在的档案ID
        nowArchiveID = self.manage_grid_files.GetCellValue(active_row_number,1)
        if len(self.manage_grid_files.GetCellValue(active_row_number,0)) == 0 :
            return
        #查询该档案的状态(是否归档),如归档则提示后返回
        isGuidang = DoCURD.query_isGuidang_values(self.CONN,{"table_name":"archives","archiveID":nowArchiveID.strip()})
        if  len(isGuidang)==0 :
            return        
        if  isGuidang[0][0] != u"否" :
            errorMessage = wx.MessageDialog(parent=None,message=u"该文件所在的档案已归档,不允许删除!",caption=u"删除文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return
        #获得指定单元格的值GetCellValue(row,col)   
        #获取要删除行的fileID
        fileID = self.manage_grid_files.GetCellValue(self.manage_grid_files.GetSelectedRows()[0],0).strip()     
        #如果当前选择不是空行继续进行
        if fileID != u"" :
            #获取所选行fileID对应的附件路径
            filepath = DoCURD.query_neirong_from_files(self.CONN, paramDict={"table_name":"files","fileID":fileID})
            #如果附件路径不为无则先删除附件,如果附件路径为无也就是没有附件则省略删除附件的步骤
            if u"无"!=filepath[0][0]:
                #如果附件在数据库中没有任何指向则删除
                temp_list = DoCURD.query_filepath_from_files(self.CONN,{"table_name":u"files","filepath":unicode(filepath[0][0])})
                #如果查询到文件路径在数据库中不止存储了一次,也就是说有多个文件都包含这个附件,则不删除附件
                #反之,如果文件路径在数据库中只存储了一次，也就是说只有一个文件包含了这个附件，则删除附件
                if len(temp_list) == 1 and os.path.exists(os.path.dirname(filepath[0][0])) :
                    os.remove(filepath[0][0])     
                    #删除空目录
                    if os.path.isdir(os.path.dirname(filepath[0][0])) and len(os.listdir(os.path.dirname(filepath[0][0])))==0:
                        os.rmdir(os.path.dirname(filepath[0][0]))
            #删除记录
            DoCURD.delete_row_from_filetable(self.CONN, paramDict={"table_name":"files","fileID":fileID})
            #刷新
            self.selectManageFilesBySomeConditions(self)
        #如果当前选择是空行,退出
        else:
            return
    
    #选中当前所选单元格所在行且激活当前选中行中所选单元格
    def activeRow_ManageFiles(self,event):
        #选中当前所选单元格所在行
        self.manage_grid_files.SelectRow(row=event.GetRow())
        #激活当前选中行中所选单元格
        self.manage_grid_files.SetGridCursor(row=event.GetRow(),col=event.GetCol())
        #获得当前选中的文件的文件ID
        fileID = self.manage_grid_files.GetCellValue(event.GetRow(),0).strip()
        if fileID != u"" :
            filepath = DoCURD.query_neirong_from_files(self.CONN, paramDict={"table_name":"files","fileID":fileID})
        else:
            return 
        #如果查到的内容(文件路径)不为无,则证明该文件有附件,可以打开
        if u'无'!=filepath[0][0] :
            self.openManage_files_button.Enable(True)
            self.OPENFILE = os.path.dirname(filepath[0][0])+"\\"
        else:
            self.openManage_files_button.Enable(False)
    
    #选中文件打开按钮    
    def openfile(self,event):
        #打开文件
        os.system("explorer.exe %s" % self.OPENFILE)
        
    def createjuanhaoByDanwei(self,event):
        #生成  卷号  下拉选择框
        temp_list = DoCURD.query_juanhao_values(self.CONN,{"table_name":"archives","quhao":self.quhao_files_comboBox.GetValue(),"guihao":self.guihao_files_comboBox.GetValue(),"hehao":self.hehao_files_comboBox.GetValue()})
        self.temp_juanhao_list = []
        for row in temp_list:
            self.temp_juanhao_list.append(row[0])
        self.temp_juanhao_list.sort()        
        if len(self.temp_juanhao_list) == 0:
            self.temp_juanhao_list = [u""] + self.temp_juanhao_list
        self.juanhao_files_comboBox.Set(self.temp_juanhao_list)
        self.juanhao_files_comboBox.SetSelection(0)
#$$$$$$$$$$被绑定的事件区结束$$$$$$$$$$

#表格类，用于接收数据库查询结果
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
    
def initTables(self):
    #生成档案表
    DoCURD.create_archives_table(CONNECTION,paramDict={"table_name":"archives"})
    #生成已销毁档案表
    DoCURD.create_destroy_archives_table(CONNECTION,paramDict={"table_name":"destroyedArchives"})
    #生成文件表
    DoCURD.create_files_table(CONNECTION, paramDict={"table_name":"files"})
    #生成已销毁文件表
    DoCURD.create_destory_files_table(CONNECTION, paramDict={"table_name":"destroyedFiles"})  
    #生成用户表
    DoCURD.create_users_table(CONNECTION)
    #如果用户表是空表则添加admin账号
    userCountList = DoCURD.query_user_count(CONNECTION)
    if len(userCountList)==0:
        DoCURD.add_user(CONNECTION,paramDict={"username":'admin',"password":'123'})

class MyApp(wx.App):
    def __init__(self):
        wx.App.__init__(self)
    def OnInit(self):
        initTables(self)        
        #生成主frame
        self.frame = MyFrame() 
        re = ''
        if not os.path.exists(u'./re.py') :
            Regist.RegistPanel(self.frame)
        else:
            with open(u'./re.py') as r :
                r.readline().strip()
                sc = r.readline().strip()
                r.readline().strip()
                d_w = r.readline().strip()
            mc_param = Regist.HardDisk.getMachineCode()+'@'+Regist.HardDisk.encodeBase64(d_w.encode('utf-8'))
            sc_param = Regist.HardDisk.createSN(mc_param)
            if sc != sc_param:
                #序列号验证
                Regist.RegistPanel(self.frame)
        
        if os.path.exists(u'./re.py') :
            with open(u'./re.py') as r :
                r.readline().strip()
                r.readline().strip()
                r.readline().strip()
                self.dw_value = r.readline().strip()
                n = UtilData.DANWEI_List.index(self.dw_value)
                self.frame.danwei_archives_comboBox.SetSelection(n)
        else:
            #设置第0个选项为默认值
            self.danwei_archives_comboBox.SetSelection(n=0)        
        #未登录则冻结所有页面
        for i in range(1,int(self.frame.main_Notebook.GetPageCount())):
            self.frame.main_Notebook.GetPage(i).Freeze()   
        #初始化时更新卷号、案卷号
        self.frame.createChoicesForJuanHao(event="")
        #self.frame.createChoicesForAnJuanHao(event="")
        #更新状态栏
        self.frame.statusBar.SetStatusText(u"档案管理")
        #打印状态栏信息
        #print self.frame.statusBar.GetStatusText()
        #设置frame为主框架
        self.SetTopWindow(self.frame)
        return True

if __name__=="__main__":
    
    app = MyApp()
    app.MainLoop()    