# -*- coding:utf-8 -*-
import wx
import wx.lib.popupctl as  pop
import wx.calendar     as  cal

#门类
MENLEI_List = [u'W_文书档案',u'J_技术与管理档案',u'K_会计档案',u'Y_音像及电子介质档案',u'S_实物档案',u'D_电子档案']
#具体各门类下的信息
MENLEI_W_List = [u'01_政策性文件',u'02_管理办法、规章制度',u'03_机构设置、人员名册等资料',u'04_会议文件、工作总结、领导讲话、会议记录、纪要、会议材料、培训、考核材料',
                 u'05_责任书、合同书',u'06_宣传报道、事迹表彰、事件处理举报及查处情况、工作简讯、简报',u'00_其它文书资料']
MENLEI_J_List = [u'01_工程规划、经营区划',u'02_年度计划',u'03_实施方案',u'04_作业设计',
                 u'05_技术标准',u'06_各种报告材料',u'07_各种统计报表、台帐',u'08_工程施工、招投标、检查验收资料、采伐证、林权证发放情况等',
                 u'09_管理形成的图、表、卡、册',u'10_管护日志、巡山记录；种苗供应材料',u'11_设施、设备、标牌建设、使用及维护情况',u'00_其它技术与管理档案']
MENLEI_K_List = [u'01_会计凭证',u'02_会计账簿',u'03_财务报告',u'04_资金使用情况统计表',
                 u'05_年度决算及相关材料',u'06_工资单、工程结算清单',u'00_其它会计资料']
MENLEI_Y_List = [u'00_音像及电子介质档案']
MENLEI_S_List = [u'00_实物档案']
MENLEI_D_List = [u'00_电子档案']

#区号
QUHAO_List = [u'01',u'02',u'03',u'04',u'05',u'06',u'07',u'08',u'30']

#柜号
GUIHAO_List = [u'11',u'12',u'13',u'14',u'15',u'21',u'22',u'23',u'24',u'25',
               u'31',u'32',u'33',u'34',u'35',u'41',u'42',u'43',u'44',u'45',
               u'51',u'52',u'53',u'54',u'55',u'61',u'62',u'63',u'64',u'65'
               ]
#盒号
HEHAO_List =[u'01',u'02',u'03',u'04',u'05',u'06',u'07',u'08',u'09'] + [ unicode(i) for i in range(10,1001)]

#卷号
JUANHAO_List = [u'01',u'02',u'03',u'04',u'05',u'06',u'07',u'08',u'09'] + [ unicode(i) for i in range(10,101)]

#级别
JIBIE_List = [u'林业局级',u'林场所级']

#期限
QIXIAN_List = [u'Y_永久',u'C_长期',u'D_短期']


#单位
DANWEI_List = [u'ZJ_张家湾农场',u'XB_西北河经营所',u'WD_五道河农场',u'WS_五四林场',u'DF_东风林场',u'WY_五一经营所',u'YJ_跃进林场',u'JX_建兴经营所',
               u'QY_七一经营所',u'YQ_义气松实验林场',u'BY_八一林场',u'BE_八二经营所',u'LG_老股流林场',u'BG_北股流林场',u'DG_东股流林场'
              ]

#位置
WEIZHI_List =  [u'ZJ_张家湾农场',u'XB_西北河经营所',u'WD_五道河农场',u'WS_五四林场',u'DF_东风林场',u'WY_五一经营所',u'YJ_跃进林场',u'JX_建兴经营所',
               u'QY_七一经营所',u'YQ_义气松实验林场',u'BY_八一林场',u'BE_八二经营所',u'LG_老股流林场',u'BG_北股流林场',u'DG_东股流林场'
              ]

#科目
KEMU_List = [u'T1_天保一期',u'T2_天保二期',u'T3_天保三期',u'T4_天保四期',u'T5_天保五期',
             u'T6_天保六期',u'T7_天保七期',u'T8_天保八期',u'T9_天保九期',u'T10_天保十期'
             ]

#密级
MIJI_List = [u'01_内部',u'02_秘密',u'03_机密',u'04_绝密']

DIQU_List = [u'牡丹江',u'哈尔滨',u'齐齐哈尔',u'佳木斯',u'大庆',u'鸡西',u'鹤岗',u'双鸭山',u'伊春',u'七台河',u'黑河',u'绥化',u'大兴安岭']

#日期类
class DateControl(pop.PopupControl):
    def __init__(self,*_args,**_kwargs):
        pop.PopupControl.__init__(self, *_args, **_kwargs)

        self.win = wx.Window(self,-1,pos = (0,0),style = 0)
        self.cal = cal.CalendarCtrl(self.win,-1,pos = (0,0))

        bz = self.cal.GetBestSize()
        self.win.SetSize(bz)

        # This method is needed to set the contents that will be displayed
        # in the popup
        self.SetPopupContent(self.win)

        # Event registration for date selection
        self.cal.Bind(cal.EVT_CALENDAR,self.OnCalSelected)


    # Method called when a day is selected in the calendar
    def OnCalSelected(self,evt):
        self.PopDown()
        date = self.cal.GetDate()

        # Format the date that was selected for the text part of the control
        self.SetValue(u'%04d年%02d月%02d日' % (date.GetYear(),
                                              date.GetMonth()+1,
                                              date.GetDay(),
                                          ))
        evt.Skip()


    # Method overridden from PopupControl
    # This method is called just before the popup is displayed
    # Use this method to format any controls in the popup
    def FormatContent(self):
        # I parse the value in the text part to resemble the correct date in
        # the calendar control
        txtValue = self.GetValue()
        dmy = txtValue.split('/')
        didSet = False

        if len(dmy) == 3:
            date = self.cal.GetDate()
            d = int(dmy[0])
            m = int(dmy[1]) - 1
            y = int(dmy[2])

            if d > 0 and d < 31:
                if m >= 0 and m < 12:
                    if y > 1000:
                        self.cal.SetDate(wx.DateTimeFromDMY(d,m,y))
                        didSet = True

        if not didSet:
            self.cal.SetDate(wx.DateTime.Today())
