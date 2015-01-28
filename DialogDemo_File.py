# -*- coding:utf-8 -*-
import wx
import os
import shutil
import UtilData
import DoCURD
import datetime

temp_list_file_frame = []
CONN = DoCURD.connect_db('./tbdata.db')
class TestDialog(wx.Dialog):
    def __init__(
            self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition, 
            style=wx.DEFAULT_DIALOG_STYLE):
        pre = wx.PreDialog()
        #pre.SetExtraStyle(wx.DIALOG_EX_CONTEXTHELP)
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)
        #生成GridBagSizer布局管理器
        modify_files_gridBagSizer = wx.GridBagSizer(vgap=10,hgap=5)
        #生成文本标签
        modify_files_quhaoText = wx.StaticText(parent=self,label=u"区        号:",style=wx.ALIGN_CENTER_VERTICAL)
        modify_files_guihaoText = wx.StaticText(parent=self,label=u"柜        号:",style=wx.ALIGN_CENTER_VERTICAL)
        modify_files_hehaoText = wx.StaticText(parent=self,label=u"盒        号:",style=wx.ALIGN_CENTER_VERTICAL)
        modify_files_juanhaoText = wx.StaticText(parent=self,label=u"卷        号:",style=wx.ALIGN_CENTER_VERTICAL)         
        modify_files_wenjiantimuText = wx.StaticText(parent=self,label=u"文件题目:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_wenjianbianhaoText = wx.StaticText(parent=self,label=u"文件编号:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_fawendanweiText = wx.StaticText(parent=self,label=u"发文单位:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_xingchengriqiText = wx.StaticText(parent=self,label=u"形成日期:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_yeshuText = wx.StaticText(parent=self,label=u"页        数:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_mijiText = wx.StaticText(parent=self,label=u"密        级:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_neirongText = wx.StaticText(parent=self,label=u"附        件:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_beizhuText = wx.StaticText(parent=self,label=u"备        注:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_inputterText = wx.StaticText(parent=self,label=u"录  入  人:",style=wx.wx.ALIGN_CENTER_VERTICAL)
        modify_files_font = wx.Font(pointSize=9,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        modify_files_quhaoText.SetFont(modify_files_font)
        modify_files_guihaoText.SetFont(modify_files_font)
        modify_files_hehaoText.SetFont(modify_files_font)
        modify_files_juanhaoText.SetFont(modify_files_font)
        modify_files_wenjiantimuText.SetFont(modify_files_font)
        modify_files_wenjianbianhaoText.SetFont(modify_files_font)
        modify_files_fawendanweiText.SetFont(modify_files_font)
        modify_files_xingchengriqiText.SetFont(modify_files_font)
        modify_files_yeshuText.SetFont(modify_files_font)
        modify_files_mijiText.SetFont(modify_files_font)
        modify_files_neirongText.SetFont(modify_files_font)
        modify_files_beizhuText.SetFont(modify_files_font)
        modify_files_inputterText.SetFont(modify_files_font)   
        #将文本标签纳入modify_files_gridBagSizer布局管理器中
        modify_files_gridBagSizer.Add(modify_files_quhaoText,pos=(0,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        modify_files_gridBagSizer.Add(modify_files_guihaoText,pos=(1,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_hehaoText,pos=(2,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_juanhaoText,pos=(3,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_wenjianbianhaoText,pos=(4,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_wenjiantimuText,pos=(5,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)        
        modify_files_gridBagSizer.Add(modify_files_neirongText,pos=(6,0),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_fawendanweiText,pos=(0,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT|wx.TOP,border=15)
        modify_files_gridBagSizer.Add(modify_files_xingchengriqiText,pos=(1,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_yeshuText,pos=(2,2),flag=wx.wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_mijiText,pos=(3,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_beizhuText,pos=(4,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT)
        modify_files_gridBagSizer.Add(modify_files_inputterText,pos=(5,2),flag=wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT) 
        ARCHIVES_SIZE = (120,25)
        #从档案表中查找到的门类、归档、单位、卷号列表,temp_list_file_frame[1]为要查找的档案ID号
        self.mgdj_List = DoCURD.query_mgdj_from_archives(CONN,paramDict={"table_name":"archives","archiveID":temp_list_file_frame[1]})
        #生成  区号  下拉选择框
        self.modify_quhao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=self.mgdj_List[0][0],size=ARCHIVES_SIZE,choices=UtilData.QUHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_quhao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_quhao_files_comboBox.SetPopupMaxHeight(self.modify_quhao_files_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件,目的是根据选择不同的门类提供不同的归档列表值
        self.modify_quhao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFileFrame)
        #生成  柜号  下拉选择框
        self.modify_guihao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=self.mgdj_List[0][1],size=ARCHIVES_SIZE,choices=UtilData.GUIHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_guihao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_guihao_files_comboBox.SetPopupMaxHeight(self.modify_guihao_files_comboBox.GetCharHeight()*10)
        self.modify_guihao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFileFrame)
        #生成  盒号  下拉选择框
        self.modify_hehao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=self.mgdj_List[0][2],size=ARCHIVES_SIZE,choices=UtilData.HEHAO_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_hehao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_hehao_files_comboBox.SetPopupMaxHeight(self.modify_hehao_files_comboBox.GetCharHeight()*10)
        #为下拉选择框绑定事件,目的是根据选择不同的门类提供不同的归档列表值
        self.modify_hehao_files_comboBox.Bind(wx.EVT_COMBOBOX, self.createChoicesForJuanHaoFileFrame)        
        
        #取得初始化时区号、柜号、盒号的值,设置案卷号控件为None方便初始化数据库查询
        self.modify_initQuhao_files_comboBox_value = self.modify_quhao_files_comboBox.GetValue()
        self.modify_initGuihao_files_comboBox_value = self.modify_guihao_files_comboBox.GetValue()
        self.modify_initHehao_files_comboBox_value = self.modify_hehao_files_comboBox.GetValue()
        self.modify_juanhao_files_comboBox = None
        #生成 卷号 初始化列表
        self.temp_accfjhf = self.createChoicesForJuanHaoFileFrame(event='')
        self.temp_list_juanhao =self.mgdj_List[0][3].split() + self.temp_accfjhf 
        #生成  卷号  下拉选择框
        self.modify_juanhao_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=self.mgdj_List[0][3],size=ARCHIVES_SIZE,choices=self.temp_list_juanhao,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_juanhao_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_juanhao_files_comboBox.SetPopupMaxHeight(self.modify_juanhao_files_comboBox.GetCharHeight()*10)
        
        #生成 文件编号 输入框
        self.modify_wenjianbianhao_files_ctrl = wx.TextCtrl(parent=self,value=temp_list_file_frame[3],size=ARCHIVES_SIZE)
        self.modify_initWenjianbianhao_files_ctrl_value = self.modify_wenjianbianhao_files_ctrl.Value.strip()
        #生成 文件题目 输入框
        self.modify_wenjiantimu_files_ctrl = wx.TextCtrl(parent=self,value=temp_list_file_frame[2],size=ARCHIVES_SIZE)
        
        #生成文件内容上传按钮
        filepath = DoCURD.query_neirong_from_files(CONN,paramDict={"table_name":"files","fileID":temp_list_file_frame[0]})
        self.modify_neirong_file_FileBrowseButton = wx.lib.filebrowsebutton.FileBrowseButton(
                                                                                     parent=self,
                                                                                     size=(310,60), 
                                                                                     labelText='', 
                                                                                     buttonText=u"上传", 
                                                                                     toolTip=u"请点击导入按钮,并选择文件的电子稿", 
                                                                                     dialogTitle=u"请选择要导入的电子稿..."
                                                                                     )
        if u'无'!=filepath[0][0]:
            self.modify_neirong_file_FileBrowseButton.SetValue(value=filepath[0][0])
            #取得内容初始化值
            self.modify_initNeirong_file_FileBrowseButton_value = filepath[0][0]
        #生成 发文单位 输入框
        self.modify_fawendanwei_files_ctrl = wx.TextCtrl(parent=self,value=temp_list_file_frame[4],size=ARCHIVES_SIZE)
        #生成 形成日期 下拉选择框
        self.modify_xingchengriqi_files_pop = UtilData.DateControl(self, -1,value=temp_list_file_frame[5], pos = (30,30),size=ARCHIVES_SIZE)
        today = datetime.datetime.now().strftime('%Y%m%d')
        #设置 形成日期 默认值
        self.modify_xingchengriqi_files_pop.SetValue(unicode(today[:4]+"年"+today[4:6]+"月"+today[6:]+"日"))
        #生成 页数 下拉选择框
        self.modify_yeshu_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,value=temp_list_file_frame[6],size=ARCHIVES_SIZE,choices=[unicode(i) for i in range(1,1001)],style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_yeshu_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_yeshu_files_comboBox.SetPopupMaxHeight(self.GetCharHeight()*10)
        #生成 密级 下拉选择框
        self.modify_miji_files_comboBox = wx.combo.OwnerDrawnComboBox(parent=self,value=temp_list_file_frame[7],id=-1,size=ARCHIVES_SIZE,choices=UtilData.MIJI_List,style=wx.CB_READONLY)
        #设置第0个选项为默认值
        self.modify_miji_files_comboBox.SetSelection(n=0)
        #设置下拉列表一次容纳10个元素
        self.modify_miji_files_comboBox.SetPopupMaxHeight(self.GetCharHeight()*10)
        #生成 备注 输入框
        self.modify_beizhu_files_ctrl = wx.TextCtrl(parent=self,value=temp_list_file_frame[8],size=ARCHIVES_SIZE)     
        #生成 录入人 输入框
        self.modify_inputter_files_ctrl = wx.TextCtrl(parent=self,value=temp_list_file_frame[9],size=ARCHIVES_SIZE) 
        #生成 新增文件 按钮并进行绑定
        self.modify_newFile_button = wx.Button(parent=self,label=u'修改文件')
        self.modify_newFile_button.Bind(wx.EVT_BUTTON, self.modifyOldFile)
        #将各选择与输入控件加入布局管理器
        modify_files_gridBagSizer.Add(self.modify_quhao_files_comboBox,pos=(0,1),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        modify_files_gridBagSizer.Add(self.modify_guihao_files_comboBox,pos=(1,1),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_hehao_files_comboBox,pos=(2,1),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_juanhao_files_comboBox,pos=(3,1),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_wenjianbianhao_files_ctrl,pos=(4,1),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_wenjiantimu_files_ctrl,pos=(5,1),flag=wx.ALIGN_LEFT) 
        #此处放置 内容 pos=(6,1)
        modify_files_gridBagSizer.Add(self.modify_neirong_file_FileBrowseButton,pos=(6,1),span=(1,3),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_fawendanwei_files_ctrl,pos=(0,3),flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        modify_files_gridBagSizer.Add(self.modify_xingchengriqi_files_pop,pos=(1,3),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_yeshu_files_comboBox,pos=(2,3),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_miji_files_comboBox,pos=(3,3),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_beizhu_files_ctrl,pos=(4,3),flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        modify_files_gridBagSizer.Add(self.modify_inputter_files_ctrl,pos=(5,3),flag=wx.ALIGN_LEFT)
        modify_files_gridBagSizer.Add(self.modify_newFile_button,pos=(7,0),span=(1,4),flag=wx.ALIGN_CENTER)    
        #各列尽可能扩展
        modify_files_gridBagSizer.AddGrowableCol(0)
        modify_files_gridBagSizer.AddGrowableCol(3)
        self.SetSizer(modify_files_gridBagSizer)
        self.Show()
        self.Center()
    ##根据修改frame的门类的选择动态提供归档类型
    #def choiceFileGuiDang(self,event):
        #TEMP = self.modify_menlei_files_comboBox.GetValue().split(u"_")[:1]
        #TEMP = "".join(TEMP)
        ##print TEMP,type(TEMP)
        #GUIDANG_MENLEI_List = []
        #if TEMP in u'MENLEI_W_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_W_List
        #if TEMP in u'MENLEI_J_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_J_List   
        #if TEMP in u'MENLEI_K_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_K_List
        #if TEMP in u'MENLEI_Y_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_Y_List 
        #if TEMP in u'MENLEI_S_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_S_List
        #if TEMP in u'MENLEI_D_List'.split(u"_"):
            #GUIDANG_MENLEI_List = UtilData.MENLEI_D_List 
        #self.modify_guidang_files_comboBox.Set(GUIDANG_MENLEI_List)
        ##默认第0个元素选中状态
        #self.modify_guidang_files_comboBox.SetSelection(0)
        
    def choiceEvent(self,event):
        pass
    
    def modifyOldFile(self,event):
        #如果isGuidang值为'是',则表示要将文件加入已归档的档案，提示禁止
        isGuidang = DoCURD.query_isGuidang_values_for_addfile(CONN,paramDict={"table_name":"archives",
                                                                  "quhao":self.modify_quhao_files_comboBox.GetValue().strip(),
                                                                  "guihao":self.modify_guihao_files_comboBox.GetValue().strip(),
                                                                  "hehao":self.modify_hehao_files_comboBox.GetValue().strip(),
                                                                  "juanhao":self.modify_juanhao_files_comboBox.GetValue().strip()
                                                                  }
                                                  )
        if u'是'==isGuidang[0][0] :
            errorMessage = wx.MessageDialog(parent=None,message=u"不允许往已经归档的档案中添加文件！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return          
        if self.modify_juanhao_files_comboBox.GetValue().strip() == u"" :
            self.Parent.Parent.statusBar.SetStatusText(u"该卷号下暂时没有档案,请在该单位下新建档案后再添加文件！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该卷号下暂时没有档案,请在该单位下新建档案后再添加文件！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return            
        #文件题目、文件编号、发文单位不允许为空
        if self.modify_wenjianbianhao_files_ctrl.Value.strip()=="" or self.modify_wenjiantimu_files_ctrl.Value.strip()=="" or self.modify_fawendanwei_files_ctrl.Value.strip()=="":
            errorMessage = wx.MessageDialog(parent=None,message=u"文件题目、文件编号、发文单位不允许为空！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy() 
            return
        #为了判断日期是否准确临时生成一个temp_cal变量
        temp_cal = self.modify_xingchengriqi_files_pop.GetValue().strip()
        if temp_cal[4] != u"年" or temp_cal[7] != u"月" or temp_cal[10] != u"日" :
            errorMessage = wx.MessageDialog(parent=None,message=u"形成日期格式错误(正确格式:****年**月**日,例如:2015年01月01日)！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()               
            return
        #要输入的文件编号如果已经存在则提示并返回
        temp_archiveID_for_wenjianbianhao = DoCURD.query_archiveID_values(CONN, paramDict={"table_name":"archives","quhao":self.modify_quhao_files_comboBox.GetValue(),
                                                                                                                   "guihao":self.modify_guihao_files_comboBox.GetValue(),
                                                                                                                   "hehao":self.modify_hehao_files_comboBox.GetValue(),
                                                                                                                   "juanhao":self.modify_juanhao_files_comboBox.GetValue()})
        temp_wenjianbianhao = DoCURD.query_wenjianbianhao_values(CONN,{"table_name":"files","archiveID":temp_archiveID_for_wenjianbianhao[0][0]})
        temp_wenjianbianhao_list = []
        for row in temp_wenjianbianhao :
            temp_wenjianbianhao_list.append(row[0])
        #如果输入的文件编号不等于初始化时的文件编号且已经在数据库中则不允许录入
        if self.modify_initWenjianbianhao_files_ctrl_value != self.modify_wenjianbianhao_files_ctrl.Value.strip() and \
           self.modify_wenjianbianhao_files_ctrl.Value.strip() in temp_wenjianbianhao_list :
            self.Parent.Parent.statusBar.SetStatusText(u"该文件编号已经存在,请核实后再次修改！")
            errorMessage = wx.MessageDialog(parent=None,message=u"该文件编号已经存在,请核实后再次修改！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()            
            return 
        #首先查询档案表中已经有的卷号
        paramDict={}
        paramDict['table_name'] = 'archives'
        paramDict['quhao'] = self.modify_quhao_files_comboBox.GetValue().strip()
        paramDict['guihao'] = self.modify_guihao_files_comboBox.GetValue().strip()
        paramDict['hehao'] = self.modify_hehao_files_comboBox.GetValue().strip()
        juanhaoAndarchiveID_exist_list = DoCURD.query_juanhaoArchiveID_values_for_addfile(CONN,paramDict=paramDict)
        #取出内容控件中的值作为新文件的值
        newFile = self.modify_neirong_file_FileBrowseButton.GetValue()
        #临时档案ID
        temp_archiveID = -1
        for row in juanhaoAndarchiveID_exist_list:
            for col in row[:1] :
                #如果查询到的第一个元素(已存在卷号)等于要存储的卷号
                if col == self.modify_juanhao_files_comboBox.GetValue().strip() :
                    #则将第二个元素(档案ID)赋值到临时档案ID变量中
                    temp_archiveID = row[-1]      
        #如果临时档案ID值不为-1,也就是文件修改页输入的档号已经存在,则可将文件执行修改操作
        if temp_archiveID != -1 :
            #如果内容控件中是空的,则将目标路径设为无
            if newFile.strip() == "" :
                targetPath = u'无'
                targetFile = u'无'               
            #如果内容控件中不是空的,则对内容控件的值进行判断
              #值正确，上传
              #值错误，提示错误信息
                #如果附件在数据库中没有任何指向则删除
                temp_list = DoCURD.query_filepath_from_files(CONN,{"table_name":u"files","filepath":unicode(self.modify_initNeirong_file_FileBrowseButton_value)})
                if len(temp_list) == 1  :
                    os.remove(self.modify_initNeirong_file_FileBrowseButton_value)    
                #删除空目录
                if os.path.isdir(os.path.dirname(self.modify_initNeirong_file_FileBrowseButton_value)) and len(os.listdir(os.path.dirname(self.modify_initNeirong_file_FileBrowseButton_value))) == 0:
                    os.rmdir(os.path.dirname(self.modify_initNeirong_file_FileBrowseButton_value))                
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
                        if "elec" in os.path.dirname(newFile) :
                            shutil.move(newFile,targetFile)
                        else :
                            shutil.copyfile(newFile,targetFile)
                            #如果附件在数据库中没有任何指向则删除
                            temp_list = DoCURD.query_filepath_from_files(CONN,{"table_name":u"files","filepath":unicode(newFile)})
                            if len(temp_list) == 0  :
                                os.remove(newFile)
                        #删除空目录
                        if os.path.isdir(os.path.dirname(newFile)) and len(os.listdir(os.path.dirname(newFile))) == 0:
                            os.rmdir(os.path.dirname(newFile))                                
                except Exception,e:
                    print e
                    errorMessage = wx.MessageDialog(parent=None,message=u"请输入正确的文件路径或使用上传按钮选择文件路径！",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
                    result = errorMessage.ShowModal()
                    errorMessage.Destroy()
                    return
            #文件存储位置设定
            paramDict['table_name'] = 'files'
            paramDict['fileID'] = temp_list_file_frame[0]
            paramDict['archiveID'] = temp_archiveID
            paramDict['wenjiantimu'] = unicode(self.modify_wenjianbianhao_files_ctrl.Value.strip())
            paramDict['wenjianbianhao'] = unicode(self.modify_wenjianbianhao_files_ctrl.Value.strip())
            paramDict['fawendanwei'] = unicode(self.modify_fawendanwei_files_ctrl.Value.strip())
            paramDict['xingchengriqi'] = unicode(self.modify_xingchengriqi_files_pop.GetValue().strip())
            paramDict['yeshu'] = unicode(self.modify_yeshu_files_comboBox.GetValue().strip())
            paramDict['miji'] = unicode(self.modify_miji_files_comboBox.GetValue().strip())
            paramDict['beizhu'] = unicode(self.modify_beizhu_files_ctrl.Value.strip())
            paramDict['inputter'] = unicode(self.modify_inputter_files_ctrl.Value.strip())
            paramDict['filepath'] = unicode(targetFile)
            #将连接和参数字典传入更新语句,修改档案信息
            updateState = DoCURD.update_some_values_from_filetable(CONN,paramDict)
            if updateState == u'ok' :
                #提示文件修改成功
                inforMessage = wx.MessageDialog(parent=None,message=u"所选文件修改成功！",caption=u"提示...",style=wx.ICON_INFORMATION)
                result = inforMessage.ShowModal()
                inforMessage.Destroy()   
                #调用父窗口(TestPanel)的父窗口(myFrame)下的管理页查询方法
                self.Parent.Parent.selectManageFilesBySomeConditions(self)    
                self.Parent.Parent.manage_grid_files.Refresh()
        #如果无，则提示要修改到的档案不存在，请核实后再修改文件
        else:
            errorMessage = wx.MessageDialog(parent=None,message=u"指定的档案索引: "+paramDict.get("quhao")+u"-"+paramDict.get("guihao")+u"-"+paramDict.get("hehao")+u"-"+self.modify_juanhao_files_comboBox.GetValue().strip()+u" 不存在,请核实后再修改文件!",caption=u"修改文件异常信息",style=wx.ICON_ERROR)
            result = errorMessage.ShowModal()
            errorMessage.Destroy()
            return
        #刷新文件管理页
        self.Parent.Parent.selectManageFilesBySomeConditions(self)
        self.Destroy()
        
        
    def createChoicesForJuanHaoFileFrame(self,event):
        #查询给定区号、柜号、盒号下的卷号列表
        temp_juanhao_list = DoCURD.query_juanhao_values(CONN,paramDict={"table_name":"archives",
                                                                         "quhao":self.modify_quhao_files_comboBox.GetValue() ,
                                                                         "guihao":self.modify_guihao_files_comboBox.GetValue() ,
                                                                         "hehao":self.modify_hehao_files_comboBox.GetValue()})
        temp_juanhaoExist_list = []
        for rows in temp_juanhao_list:
            temp_juanhaoExist_list.append(rows[0])  
        #删除重复元素
        if self.mgdj_List[0][3] in temp_juanhaoExist_list :
            temp_juanhaoExist_list.remove(self.mgdj_List[0][3])
        #如果元素在全列表中而不在已存在列表中则将该元素返回
        temp_juan_choices_list = temp_juanhaoExist_list
        if len(temp_juan_choices_list) == 0 :
            temp_juan_choices_list = [u" "] + temp_juan_choices_list
        if self.modify_juanhao_files_comboBox is not None :
            #如果原始单位值等于现在单位值,则将原始单位值下的案卷号列表赋值到现在案卷号列表下
            if self.modify_initQuhao_files_comboBox_value == self.modify_quhao_files_comboBox.GetValue() and \
               self.wx.modify_initGuihao_files_comboBox_value == self.modify_guihao_files_comboBox.GetValue() and \
               self.modify_initHehao_files_comboBox_value == self.modify_hehao_files_comboBox.GetValue :
                temp_juan_choices_list = self.temp_list_juanhao              
            #设置卷号控件中数据源为新生成的temp_anjuan_choices_list
            self.modify_juanhao_files_comboBox.Set(temp_juan_choices_list)
            #设置第0个元素为默认显示
            self.modify_juanhao_files_comboBox.SetSelection(0)
        else:
            return temp_juan_choices_list        
#---------------------------------------------------------------------------

class TestPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        dlg = TestDialog(self, -1, u"文件修改", size=(450,370),
                         style=wx.DEFAULT_DIALOG_STYLE)
        
        dlg.CenterOnScreen()
        val = dlg.ShowModal()    
        dlg.Destroy()
        