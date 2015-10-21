#!/usr/bin/python
#-*- encoding:utf-8 -*-
import wx
import socket
import threading,time
import json
import collections
import struct
import xlrd
import os

ADDR = ('192.168.1.106', 8888)
BUFSIZE = 1024
FILEINFO_SIZE=struct.calcsize('128s32sI8s')
filename = '1.xls'

JIANQU_NUM = 19
ITEM_NUM = 15




class WorkerThread(threading.Thread):
    """docstring for WorkerThread"""
    def __init__(self, window):
        super(WorkerThread, self).__init__()
        self.window = window

    def run(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(ADDR)
        s.listen(50)
        while True:
        # 接受一个新连接:
            sock, addr = s.accept()
        # 创建新线程来处理TCP连接:
            t = threading.Thread(target=self.tcplink, args=(sock, addr))
            t.start()

    def tcplink(self,sock, addr):
        fhead=struct.pack('128s11I',filename,0,0,0,0,0,0,0,0,os.stat(filename).st_size,0,0)
        sock.send(fhead)
        excelfile = open(filename,'rb')
        while True:
            filedata = excelfile.read(BUFSIZE)
            if not filedata:   
                break
            sock.send(filedata)   
        excelfile.close()     
        sock.close() 
        

class ServerFrame(wx.Frame):
    """ Server Frame. """
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(400,100))
        panel = wx.Panel(self,-1)
        box = wx.BoxSizer(wx.VERTICAL)
        statictext = wx.StaticText(panel,-1,u'服务器程序正在允许，请不要关闭!',(400, 50),(400, -1), wx.ALIGN_CENTER)
        statictext.SetFont(wx.Font(15, wx.SWISS, wx.NORMAL, wx.NORMAL,False, u'Tahoma'))

        box.Add(statictext)
        
        self.initConfig()
        panel.SetSizer(box)
        #self.taskBarIcon = wx.TaskBarIcon()
        #self.Bind(wx.EVT_CLOSE, self.OnClose)
        #self.Bind(wx.EVT_ICONIZE, self.OnIconfiy)


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
        except Exception, e:
            wx.MessageBox(u'未找到配置文件，IP地址将无法配置！',u'错误',wx.OK)
        if not os.path.exists('1.xls'):
            wx.MessageBox(u'未找到罪犯信息！',u'错误',wx.OK)

    def OnIconfiy(self, event):
        self.Hide()
        event.Skip()
    def OnClose(self, event):
        self.taskBarIcon.Destroy()
        self.Destroy()





                      







        




if __name__ == '__main__':
    app = wx.App(False)
    frame = ServerFrame(None, u'哈尔滨监狱罪犯动态信息')
    app.MainLoop()