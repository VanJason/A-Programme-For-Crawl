# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import sys
from ui_mainwindow import Ui_MainWindow
from programme_for_guoyi import *
from programme_for_gdcw import *
from programme_for_gzcw import *

class MainWindow(QMainWindow, Ui_MainWindow,Programme_guoyi,Programme_gdcw,Programme_gzcw):


    # _signal6=QtCore.pyqtSignal()

    listrecord = [None,None,None,"",""]
    infolist = [None,None,None,None,None,None,None,None]
    pagemovement = [None,None,False]

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        self.Cancle.clicked.connect(self.close)
        self.leadtoOrcal.clicked.connect(self.setCreatOracl)
        self.creatExcel.clicked.connect(self.setCreatExcel)

        self.guoyiButton.toggled['bool'].connect(self.guoyiWeb)
        self.gdcwButton.toggled['bool'].connect(self.gdcw)
        self.gzcwButton.toggled['bool'].connect(self.gzcw)
        self.zhaobiaoButton.toggled['bool'].connect(self.zhaobiao)
        self.zhongbiaoButton.toggled['bool'].connect(self.zhongbiao)
        self.caigourenText.textEdited['QString'].connect(self.caigouren)
        self.biaotiText.textEdited['QString'].connect(self.biaoti)
        self.filenameText.textEdited['QString'].connect(self.filename)
        self.P_title.toggled['bool'].connect(self.gettitle)
        self.P_account.toggled['bool'].connect(self.getaccount)
        self.P_beginningtime.toggled['bool'].connect(self.getbeginningtime)
        self.P_agentCompany.toggled['bool'].connect(self.getagentcompany)
        self.P_link.toggled['bool'].connect(self.getlink)
        self.P_getaccount.toggled['bool'].connect(self.getsellaccount)
        self.P_showtime.toggled['bool'].connect(self.getshowtime)
        self.P_buyer.toggled['bool'].connect(self.getbuyer)

        self.starttime.dateChanged['QDate'].connect(self.startdate)
        self.endtime.dateChanged['QDate'].connect(self.enddate)
        self.chooseall.toggled['bool'].connect(self.chooseallmessage)

    def startdate(self,QDate):
        datetuple = QDate.getDate()
        datelist = []
        datelist.append(str(datetuple[0]))
        if datetuple[1] < 10:
            datelist.append("0" + str(datetuple[1]))
        else:
            datelist.append(str(datetuple[1]))
        if datetuple[2] < 10:
            datelist.append("0"+str(datetuple[2]))
        else:
            datelist.append(str(datetuple[2]))

        date =datelist[0]+datelist[1]+datelist[2]
        self.pagemovement[0] = int(date)
    def enddate(self,QDate):
        datetuple = QDate.getDate()
        datelist = []
        datelist.append(str(datetuple[0]))
        if datetuple[1] < 10:
            datelist.append("0" + str(datetuple[1]))
        else:
            datelist.append(str(datetuple[1]))
        if datetuple[2] < 10:
            datelist.append("0"+str(datetuple[2]))
        else:
            datelist.append(str(datetuple[2]))

        date =datelist[0]+datelist[1]+datelist[2]
        self.pagemovement[1] = int(date)
    def chooseallmessage(self,state):
        if state == True:
            self.pagemovement[2] = True
        else:
            self.pagemovement[2] = False

    """
    单项选择——招标网站选择方法
    定义传递的招标网站信号的槽
    定义信号
    """
    def guoyiWeb(self,state):
        if state is True:
            self.listrecord[0] = "guoyi"
        else:
            None
    def gdcw(self,state):
        if state is True:
            self.listrecord[0] = "gdcw"
        else:
            None
    def gzcw(self,state):
        if state is True:
            self.listrecord[0] = "gzcw"
        else:
            None


    """
    单项选择——检索类型选择方法
    定义传递的类型信号的槽
    定义信号
    """
    def zhaobiao(self,state):
        if state is True:
            self.listrecord[1] = "招标类"
        else:
            None    
    def zhongbiao(self,state):
        if state is True:
            self.listrecord[1] = "中标类"
        else:
            None

    """
    填写生成文件
    定义信号
    """
    def filename(self,strtext):
        self.strtext = strtext + ".xls"
        self.listrecord[2] = self.strtext

    """
    填写检索关键词
    定义信号
    """    
    def caigouren(self,strtext):
        if strtext is None:
            self.listrecord[3] = ""
        else:
            self.strtext = strtext
            self.listrecord[3] = self.strtext
    def biaoti(self,strtext):
        if strtext is None:
            self.listrecord[4] = ""
        else:
            self.strtext = strtext
            self.listrecord[4] = self.strtext

    """多选——选择要获取的信息"""
    def gettitle(self,state):
        self.state = state
        if state is True:
            self.infolist[0] = self.state 
        else:
            self.infolist[0] = False     
    def getaccount(self,state):
        self.state = state
        if state is True:
            self.infolist[1] = self.state 
        else:
            self.infolist[1] = False          
    def getbeginningtime(self,state):
        self.state = state
        if state is True:
            self.infolist[2] = self.state 
        else:
            self.infolist[2] = False 
    def getagentcompany(self,state):
        self.state = state
        if state is True:
            self.infolist[3] = self.state 
        else:
            self.infolist[3] = False 
    def getlink(self,state):
        self.state = state
        if state is True:
            self.infolist[4] = self.state 
        else:
            self.infolist[4] = False 
    def getsellaccount(self,state):
        self.state = state
        if state is True:
            self.infolist[5] = self.state 
        else:
            self.infolist[5] = False  
    def getshowtime(self,state):
        self.state = state
        if state is True:
            self.infolist[6] = self.state 
        else:
            self.infolist[6] = False
    def getbuyer(self,state):
        self.state = state
        if state is True:
            self.infolist[7] = self.state 
        else:
            self.infolist[7] = False 

    def setCreatOracl(self):
        print(self.listrecord,self.infolist,self.pagemovement)
    	# QtWidgets.QMessageBox.information(self.creatExcel,"标题","生成Excel成功")
    def setCreatExcel(self):

        if self.listrecord[0] =="guoyi":
            url = 'http://www.gmgit.com/'
            if self.listrecord[1] =="招标类":
                self.listrecord[1] = "招标公告"
            elif self.listrecord[1] == "中标类":
                self.listrecord[1] = "招标结果"
            else:
                QtWidgets.QMessageBox.information(self,"标题","请选择类型并重新爬虫")

            info = Programme_guoyi.get_Web_guoyi(self,url,self.listrecord[1],self.listrecord[4])
            WBall = Programme_guoyi.get_programme_guoyi(self,info,self.pagemovement[0],self.pagemovement[1],self.pagemovement[2])
            Programme_guoyi.get_detail_guoyi(self,WBall,self.listrecord[2],self.infolist[0],self.infolist[1],self.infolist[2],self.infolist[3],self.infolist[4],self.infolist[5],self.infolist[6],self.infolist[7])
            QtWidgets.QMessageBox.information(self.creatExcel,"标题","生成国义招标网爬虫成功")
        
        elif self.listrecord[0] =="gdcw":
            url = 'http://www.gdgpo.gov.cn/'
            if self.listrecord[1] == "招标类":
                self.listrecord[1] = "采购公告"
            elif self.listrecord[1] =="中标类":
                self.listrecord[1] = "中标公告"
            else:
                QtWidgets.QMessageBox.information(self,"标题","请选择类型并重新爬虫")

            info = Programme_gdcw.get_content_gdcw(self,url,self.listrecord[1],self.listrecord[4],self.listrecord[3])
            WBall = Programme_gdcw.get_Web_gdcw(self,info,self.pagemovement[0],self.pagemovement[1],self.pagemovement[2])
            Programme_gdcw.get_programme_gdcw(self,WBall,self.listrecord[2],self.infolist[0],self.infolist[1],self.infolist[2],self.infolist[3],self.infolist[4],self.infolist[5],self.infolist[6],self.infolist[7])
            QtWidgets.QMessageBox.information(self.creatExcel,"标题","生成广东招标网爬虫成功")
        
        elif self.listrecord[0] == "gzcw":
            if self.listrecord[1] == "招标类":
                url = 'http://www.gzggzy.cn/cms/wz/view/index/layout2/zfcglist.jsp?siteId=1&channelId=456'
            elif self.listrecord[1] == "中标类":
                url = 'http://www.gzggzy.cn/cms/wz/view/index/layout2/zfcglist.jsp?siteId=1&channelId=458'
            else:
                QtWidgets.QMessageBox.information(self,"标题","请选择类型并重新爬虫")
            info = Programme_gzcw.get_content2_gzcw(self,url,self.listrecord[4])
            WBall = Programme_gzcw.get_web_gzcw(self,info,self.pagemovement[0],self.pagemovement[1],self.pagemovement[2])
            Programme_gzcw.get_detail_gzcw(self,WBall,self.listrecord[2],self.infolist[0],self.infolist[1],self.infolist[2],self.infolist[3],self.infolist[4],self.infolist[5],self.infolist[6],self.infolist[7])
            QtWidgets.QMessageBox.information(self.creatExcel,"标题","生成广州公共资源交易网爬虫成功")
        else:
            QtWidgets.QMessageBox.information(self.creatExcel,"标题","请选择要爬虫的网站")

if __name__ == "__main__":   
    app = QtWidgets.QApplication(sys.argv)
    myWindow = MainWindow()
    myWindow.show()
    sys.exit(app.exec_())
