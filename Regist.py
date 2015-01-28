# -*- coding:utf-8 -*-
import os
import wx
import wx.combo
import HardDisk
import UtilData

class RegistDialog(wx.Dialog):
    def __init__(
            self, parent, ID, title, size=wx.DefaultSize, pos=wx.DefaultPosition, 
            style=wx.DEFAULT_DIALOG_STYLE):
        pre = wx.PreDialog()
        pre.Create(parent, ID, title, pos, size, style)
        self.PostCreate(pre)
        hbox0 = wx.BoxSizer(wx.HORIZONTAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        vbox = wx.BoxSizer(wx.VERTICAL)
        font_sm = wx.Font(pointSize=12,family=wx.ROMAN,style=wx.NORMAL,weight=wx.BOLD)
        diquText = wx.StaticText(parent=self,label=u"地区:")
        diquText.SetFont(font_sm)
        self.diqucombo = wx.combo.OwnerDrawnComboBox(parent=self,id=-1,size=(100,25),choices=UtilData.DIQU_List,style=wx.CB_READONLY)
        self.diqucombo.SetSelection(0)
        self.diqucombo.SetPopupMaxHeight(self.diqucombo.GetCharHeight()*10)
        hbox0.Add(diquText,flag=wx.ALIGN_CENTER)
        hbox0.Add(self.diqucombo,flag=wx.ALIGN_CENTER)
        danweiText = wx.wx.StaticText(parent=self,label=u'单位:')
        danweiText.SetFont(font_sm)
        self.danweiCtrl = wx.TextCtrl(parent=self,size=(100,25))
        hbox0.Add(danweiText,flag=wx.ALIGN_CENTER)
        hbox0.Add(self.danweiCtrl,flag=wx.ALIGN_CENTER)
        machineCodeText = wx.StaticText(parent=self,label=u"机器码:")
        machineCodeText.SetFont(font_sm)
        self.machineCodeCopyText = wx.StaticText(parent=self,label=u'')
        self.machineCodeCopyText.SetFont(font_sm)
        hbox1.Add(machineCodeText,flag=wx.ALIGN_LEFT)
        hbox1.Add(self.machineCodeCopyText,flag=wx.ALIGN_CENTER)
        registCodeText = wx.StaticText(parent=self,label=u"序列号:")
        registCodeText.SetFont(font_sm)
        self.registCtrl = wx.TextCtrl(parent=self,size=(225,20))
        hbox2.Add(registCodeText,flag=wx.ALIGN_CENTER)
        hbox2.Add(self.registCtrl,flag=wx.ALIGN_CENTER)
        self.createButton = wx.Button(parent=self,label=u"生成机器码")
        self.createButton.Bind(wx.EVT_BUTTON, self.createMachineCode)
        self.copyButton = wx.Button(parent=self,label=u"复制机器码")
        self.copyButton.Bind(wx.EVT_BUTTON, self.copyMachineCode)
        self.registButton = wx.Button(parent=self,label=u"注册")
        self.registButton.Bind(wx.EVT_BUTTON,self.checkSecretCode)
        hbox3.Add(self.createButton,flag=wx.ALIGN_CENTER)
        hbox3.Add(self.copyButton,flag=wx.ALIGN_CENTER)
        hbox3.Add(self.registButton,flag=wx.ALIGN_CENTER)
        vbox.Add(hbox0,flag=wx.ALIGN_LEFT|wx.TOP,border=15)
        vbox.Add(hbox1,flag=wx.ALIGN_LEFT|wx.TOP,border=10)
        vbox.Add(hbox2,flag=wx.ALIGN_LEFT|wx.TOP,border=10)
        vbox.Add(hbox3,flag=wx.ALIGN_CENTER|wx.TOP,border=15)
        self.SetSizer(vbox)
        self.Show()
        self.Center()
    
    def createMachineCode(self,event):
        self.registButton.Enable(False)
        self.copyButton.Enable(False) 
        self.createButton.Enable(False)
        if self.diqucombo.GetValue().strip() == u'' or self.danweiCtrl.Value.strip() == u'' :
            wx.MessageBox(u"请填写地区和单位后再生成机器码！",u'注册提示')
            self.registButton.Enable(True)
            self.copyButton.Enable(True) 
            self.createButton.Enable(True)            
            return
        param = HardDisk.getMachineCode()+'@'+HardDisk.encodeBase64(self.danweiCtrl.Value.strip().encode('utf-8'))
        self.machineCodeCopyText.SetLabel(param)
        self.createButton.Enable(True) 
        self.registButton.Enable(True)
        self.copyButton.Enable(True) 
        
    
    def checkSecretCode(self,event):
        self.registButton.Enable(False)
        self.copyButton.Enable(False) 
        self.createButton.Enable(False) 
        if self.registCtrl.Value.strip() == u'' or self.danweiCtrl.Value.strip() == u'' or self.machineCodeCopyText.GetLabel().strip() == u'':
            wx.MessageBox(u"请先填写单位、生成机器码、粘贴序列号后再点击注册！",u'注册提示')
            self.registButton.Enable(True)
            self.copyButton.Enable(True) 
            self.createButton.Enable(True) 
            return
        param2 = HardDisk.createSN(self.machineCodeCopyText.GetLabel().strip())
        if param2 == self.registCtrl.GetValue().strip():
            try:
                with open(u"./re.py","w") as o:
                    o.write(self.registCtrl.GetValue().strip()+u"\n") 
                    o.write(self.diqucombo.GetValue().strip()+u"\n")
                    o.write(self.danweiCtrl.GetValue().strip())
                wx.MessageBox(u"序列号有效,注册成功！",u'注册成功')
            except Exception,e:
                print e.args[0]
            self.Parent.FLAG = True
            self.Destroy()
        else:
            wx.MessageBox(u"序列号无效,注册失败！",u'注册失败')
        self.registButton.Enable(True)
        self.copyButton.Enable(True) 
        self.createButton.Enable(True)
    def copyMachineCode(self,event):
        self.registButton.Enable(False)
        self.copyButton.Enable(False) 
        self.createButton.Enable(False)
        self.dataObj = wx.TextDataObject()
        self.dataObj.SetText(self.machineCodeCopyText.GetLabelText())
        if wx.TheClipboard.Open() and self.machineCodeCopyText.GetLabel().strip() != u'':
            wx.TheClipboard.SetData(self.dataObj)
            wx.TheClipboard.Flush()
            wx.MessageBox(u"已将机器码复制到剪切板,请将其发送到lgyhome@163.com中获取序列号！",u'提示')
        else:
            wx.MessageBox(u"请先生成机器码后再复制！", "提示")                
        self.registButton.Enable(True)
        self.copyButton.Enable(True) 
        self.createButton.Enable(True)
#---------------------------------------------------------------------------

class RegistPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        self.FLAG = False
        dlg = RegistDialog(self, -1, u"用户注册", size=(285,180),
                         style=wx.DEFAULT_DIALOG_STYLE)
        dlg.CenterOnScreen()
        val = dlg.ShowModal()
        dlg.Destroy()
        if self.FLAG == False:
            self.Parent.Destroy()

    
    
