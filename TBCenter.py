#-*- coding:utf-8 -*-

import wx
import TBSingle
#import datetime
'''
635*112
'''
class TBFrame(wx.Frame):
        #框架界面初始化,用户各元素外观布局^(wx.MAXIMIZE_BOX|wx.RESIZE_BORDER)
        def __init__(self):
                wx.Frame.__init__(self,parent=None,title=u"天保档案管理系统(汇总版)   v1.0",size=(640,480),style=wx.DEFAULT_FRAME_STYLE^(wx.MAXIMIZE_BOX|wx.RESIZE_BORDER))
                #frame居中
                self.Center(wx.BOTH)
                #生成panel
                self.panel = wx.Panel(parent=self) 
                dwList = [u'ZJ_张家湾农场',u'XB_西北河经营所',u'WD_五道河农场',u'WS_五四林场',u'DF_东风林场',u'WY_五一经营所',u'YJ_跃进林场',u'JX_建兴经营所',
               u'QY_七一经营所',u'YQ_义气松实验林场',u'BY_八一林场',u'BE_八二经营所',u'LG_老股流林场',u'BG_北股流林场',u'DG_东股流林场']
                hbox = wx.BoxSizer(wx.HORIZONTAL)
                logo_center = wx.StaticBitmap(parent=self.panel,bitmap=wx.Bitmap("logo_center.bmp",wx.BITMAP_TYPE_BMP),size=(635,112))
                hbox.Add(logo_center,flag=wx.ALIGN_TOP)
                vbox = wx.BoxSizer(wx.VERTICAL)
                gs = wx.GridSizer(4, 4, 20, 10)
                for i in range(15):
                        tempButton = wx.Button(parent=self.panel,label=dwList[i][3:],size=(150,50))
                        gs.Add(tempButton,flag=wx.ALIGN_CENTER)
                        tempButton.Bind(wx.EVT_BUTTON, self.showData)
                tongjiButton = wx.Button(parent=self.panel,label=u'统计',size=(150,50))
                gs.Add(tongjiButton,flag=wx.ALIGN_CENTER)
                tongjiButton.Bind(wx.EVT_BUTTON,self.tongji)
                vbox.Add(hbox,flag=wx.ALIGN_TOP)
                vbox.Add(gs,flag=wx.ALIGN_CENTER|wx.TOP,border=30)
                self.panel.SetSizer(vbox)
                self.Show()
        #数据统计
        def tongji(self,event):
                pass
        #各单位数据查看
        def showData(self,event):
                #event.GetEventObject()获取触发该事件的事件源对象
                #print datetime.datetime.now().strftime('%Y%m%d'),type(datetime.datetime.now().strftime('%Y%m%d'))
                danwei_click = event.GetEventObject().GetLabel()
                TBSingle.DANWEI = danwei_click
                TBSingle.SinglePanel(self)
                pass
class TBApp(wx.App):
        def __init__(self):
                wx.App.__init__(self)
        def OnInit(self):
                #initTables(self)        
                #生成主frame
                self.frame = TBFrame() 
                self.SetTopWindow(self.frame)
                return True
        
if __name__=="__main__":
        app = TBApp()
        app.MainLoop() 