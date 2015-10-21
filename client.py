#!/usr/bin/env python
# -*- coding: utf-8 -*-
import wx
import socket
import  os
import threading,time
import xlrd
import struct
import random
import  wx.lib.scrolledpanel as scrolled



JIANQU_NUM = 19
ITEM_NUM = 15

APP_ABOUT = 1
APP_UPDATE = 2

#ADDR = ('192.168.1.106', 8888)
BUFSIZE = 1024
FILEINFO_SIZE=struct.calcsize('128s32sI8s')

filename = None


jianqu = [u'一监区',
              u'二监区',
              u'三监区',
              u'四监区',
              u'五监区',
              u'六监区',
              u'七监区',
              u'八监区',
              u'九监区',
              u'十监区',
              u'十一监区',
              u'十二监区',
              u'后勤监区',u'出监教育监区',u'外籍犯监区',u'集训监区',u'病犯监区',u'高戒监区',u'合计']



datas = [u"监区",u"监区在册",u"保外就医",u"解回再审",u"监内实押",u"车间出工",
                u"监狱外就医",u"监狱内住院",u"接见人数",u"超市购物",u"禁闭人数",u"严管人数",u"监狱外就医"
                ,u"监狱内就医",u"备注项人数",u"各监区监舍",u"报表民警",u"带班民警",u"值班民警",u"值班民警",u"值班民警"
                ,u"加班民警",u"加班民警",u"加班民警",u"加班民警",u"加班民警",u"加班民警",u"加班民警"]

staticTexts = []
staticDatas = []
INIT = True

class WorkerThread(threading.Thread):
    """docstring for WorkerThread"""
    def __init__(self, window):
        super(WorkerThread, self).__init__()
        self.window = window

    def run(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.connect(ADDR)
            print ADDR
            fhead = s.recv(FILEINFO_SIZE)
            filename,temp1,filesize,temp2=struct.unpack('128s32sI8s',fhead)
            filename='new_'+filename.strip('\00')
            fp = open(filename,'wb')
            restsize = filesize
            while True:
                if restsize>BUFSIZE:
                    filedata = s.recv(BUFSIZE)
                else:
                    filedata = s.recv(restsize)
                if not filedata:break
                fp.write(filedata)
                restsize = restsize-len(filedata)
                if restsize == 0:
                    break
            fp.close()
            if INIT:
                wx.CallAfter(self.window.InitData,self.window.panel)
                print 'true'
            else:
                wx.CallAfter(self.window.updateData,self.window.panel)
                print 'false'

        except socket.error, arg:
            (errno, err_msg) = arg
            print "Connect server failed: %s, errno=%d" % (err_msg, errno)
            if errno == 10061:
                wx.MessageBox(u'服务器端没有开启，请稍候重试！',u'错误！',wx.CANCEL)
        finally:
            
            s.close()



class ClientFrame(wx.Frame):

    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(1300,840))
        panel = self.panel = scrolled.ScrolledPanel(self, -1, size=(1300, 640),style = wx.TAB_TRAVERSAL|wx.SUNKEN_BORDER, name="panel" )
        #panel = self.panel = wx.Panel(self,-1)
        #self.button = wx.Button(panel,-1,'Hello',pos=(50,20))
        #self.Bind(wx.EVT_BUTTON,self.OnClick,self.button)
        #self.button.SetDefault()
        #self.initStaticText(panel)



        #时间状态栏
        self.timer = wx.Timer(self)
        string = time.strftime('%Y-%m-%d    %H:%M:%S')
        datetext = self.datetext = self.CreateStatusBar()
        self.datetext.SetFieldsCount(3)
        self.datetext.SetStatusWidths([-3, -2, -1])
        self.datetext.SetStatusText(string, 2)
        #datetext = self.datetext = wx.StaticText(panel,-1,string,size=(200,30),style=wx.ALL)
        #timebox = self.timebox = wx.BoxSizer(wx.HORIZONTAL)
        #timebox.Add(self.datetext)
        self.Bind(wx.EVT_TIMER, self.OnTimer,self.timer)
        self.timer.Start(1000)


        self.initConfig()
        #print ADDR
        self.createPanel(panel)
        self.InitMenu()
        panel.SetupScrolling()
        #self.InitData(panel)
        self.Show(True)
        t = WorkerThread(self)
        t.setDaemon(True)
        t.start()

        

    def initConfig(self):
        try:
            configfile = open('config','r')
            addr = configfile.readline()
            if addr:
                global ADDR 
                ADDR = tuple(eval(addr))
                #print ADDR
        except Exception, e:
            wx.MessageBox(u'未找到配置文件，IP地址将无法配置！',u'错误',wx.OK)





    def createPanel(self,panel):
        
        

        #print type(string)
        #print string
        box = wx.BoxSizer(wx.HORIZONTAL)
        for data in datas:
            statictext = wx.StaticText(panel,-1,data,size=(80,30),style=wx.ALIGN_CENTER)
            box.Add(statictext)
        box2 = wx.BoxSizer(wx.VERTICAL)
        for data in jianqu:
            statictext = wx.StaticText(panel,-1,data,size=(80,30),style=wx.ALIGN_CENTER)
            box2.Add(statictext)
        self.box4 = wx.GridBagSizer(0,0)
        box3 =self.box3= wx.BoxSizer(wx.HORIZONTAL)
        box3.Add(box2)
        box3.Add(self.box4)
        vbox = self.vbox = wx.BoxSizer(wx.VERTICAL)
        #vbox.Add(self.timebox)
        vbox.Add(box)
        vbox.Add(box3)
        panel.SetSizer(vbox)

    def InitData(self,panel):

        if os.path.exists('new_1.xls'):
            try:
                
                data = xlrd.open_workbook('new_1.xls')
                table = data.sheets()[0] 
                nrows = table.nrows
                ncols = table.ncols
                cell_A1 = table.cell(1,1).value
                for i in range(1,nrows):
                    for j in range(1,ncols):
                        #print str(nrows)+'JIANQU_NUM'
                        data = table.row(i)[j].value
                        #print type(data),data
                        #if isinstance(data,unicode):
                        #    print data
                        #    print type(data)
                        #    print '@'*10
                        #    print data
                        if isinstance(data,float):
                            data = str(int(data))
                        if not data: data = ''

                        statictext = wx.StaticText(panel,-1,data,pos=(1200,800),size=(80,30),style=wx.ALIGN_CENTER)
                        staticTexts.append(statictext)
                        self.box4.Add(statictext,(i-1,j-1))
            except Exception, e:
                print e
                #wx.MessageBox(u'未找到罪犯信息！',u'错误',wx.CANCEL)
        global INIT 
        INIT = False
        self.vbox.Layout()

    def updateData(self,panel):

        if os.path.exists('new_1.xls'):
            try:
                data = xlrd.open_workbook('new_1.xls')
                table = data.sheets()[0] 
                nrows = table.nrows
                ncols = table.ncols
                cell_A1 = table.cell(1,1).value
                for i in range(1,nrows):
                    for j in range(1,ncols):
                        #print i,j
                        data = table.row(i)[j].value
                        #if isinstance(data,unicode):
                        #    print data
                        #    print type(data)
                        #    print '@'*10
                        #    print data
                        if isinstance(data,float):
                            data = str(int(data))
                        if not data: data = ''
                        staticDatas.append(data)
                for i in range(len(staticTexts)):
                        statictext = staticTexts[i]
                        data = staticDatas[i]
                        #if isinstance(data,unicode):
                        #    print data
                        #    print type(data)
                        #    print '@'*10
                        #    print data
                        if isinstance(data,float):
                            data = str(int(data))
                        if not data: data = ''

                        statictext.SetLabel(data)

                del staticDatas[:]
            

            except Exception, e:
                print e
                #wx.MessageBox(u'未找到罪犯信息！',u'错误',wx.CANCEL)
        else:
            wx.MessageBox(u'未找到罪犯信息！',u'错误',wx.OK)
        self.vbox.Layout()


    def InitMenu(self):
        menuBar = wx.MenuBar()  

        filemenu = self.filemenu = wx.Menu()  
        aboutmenu = wx.Menu() 
        popupmenu =self.popupmenu = wx.Menu()
        updatepop = popupmenu.Append(-1, u'更新')


        update = filemenu.Append(wx.ID_ANY,u"更新","update Applications")
        fitem = filemenu.Append(wx.ID_EXIT,u"退出","Quit Applications") 
        aboutitem = aboutmenu.Append(wx.ID_ANY,u"关于","about") 
        menuBar.Append(filemenu,u"&功能")
        menuBar.Append(aboutmenu,u"&关于")

        self.SetMenuBar(menuBar)  
          
        self.Bind(wx.EVT_MENU, self.OnQuit, fitem) 
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutitem)
        self.Bind(wx.EVT_MENU, self.OnUpdate, update)
        self.Bind(wx.EVT_MENU, self.OnRightDown, updatepop)
        
        self.panel.Bind(wx.EVT_CONTEXT_MENU, self.OnShowPopup)

    def OnShowPopup(self, event):
        pos = event.GetPosition()
        pos = self.panel.ScreenToClient(pos)
        self.panel.PopupMenu(self.popupmenu, pos)

    def OnQuit(self,e):  
        self.Close()

    def OnUpdate(self,e):
        t = WorkerThread(self)
        t.setDaemon(True)
        t.start()

    def OnAbout(self,e):
        wx.MessageBox(u'哈尔滨监狱罪犯信息动态管理客户端v1.0',u'made by 韩 峰',wx.OK)

    def OnRightDown(self,e):
        self.OnUpdate(e) 

    def OnTimer(self,e):
        string = time.strftime('%Y-%m-%d    %H:%M:%S')
        #print string
        self.datetext.SetStatusText(string, 2)



if __name__ == '__main__':
    app = wx.App(False)
    frame = ClientFrame(None, u'哈尔滨监狱罪犯信息动态管理客户端v1.0')
    app.MainLoop()


