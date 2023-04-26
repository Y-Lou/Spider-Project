# Python 模块
import sys
from glob import glob
import re
import os

# Excel文件转换类
import xlwings as xw

# 界面类
from SeleCrawlerWidget import Ui_SeleCrawlerForm

# 子窗口与线程类
from MassageWindow import MassageWindow
from SeleCrawlerThread import seleCrawlerThread

# PyQt5类
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication,QFileDialog,QMessageBox,QLabel,QWidget
from PyQt5.QtGui import QMovie,QIcon
from PyQt5.QtCore import QDir

# 主窗口类(继承QWidget)
class MainWidget(QWidget):
    def __init__(self) -> None: # 构造函数
        super().__init__()
        self.ui:Ui_SeleCrawlerForm = Ui_SeleCrawlerForm()                    # 实例化界面类
        self.ui.setupUi(self) 
        self.m_MassageWindow:MassageWindow = MassageWindow()                              # 实例化爬取信息类(主要记录那些文件中的编号没有数据)
        self.m_seleCrawlerThread:seleCrawlerThread = seleCrawlerThread()     # 子线程类
        self.m_seleCrawlerThread.SetMassageClass(self.m_MassageWindow)       # 将MassageWindow类传入子线程里
        self.iFireFoxPath:str = ""                                           # 浏览器路径

        self.IsFolderChoose:bool = False                                     # 是否选择文件夹
        
        # 先隐藏相应的控件，有问题时再启用
        self.ui.DriverPathcomboBox.hide()
        self.ui.label_3.hide()
        self.ui.FireFoxPathEdit.setText('请输入火狐浏览器的路径')
        # self.ui.label_2.hide()
        self.ui.FilePathEdit.hide()
        self.ui.FileChooseButton.hide()
        # self.ui.StartButton.setEnabled(False)

        self.SignalConnect()    # 信号连接

        # 加载等待动画
        # 不显示的原因是因为主线程的资源被占据了
        self.labelLoad = QLabel("",self)
        self.labelLoad.setGeometry(134,10,124,124)
        self.movieLoad=QMovie('./resources/loading.gif')
        self.labelLoad.setMovie(self.movieLoad)
        self.movieLoad.start()
        self.labelLoad.hide()
    
    """
    信号连接
    """
    def SignalConnect(self):
        self.ui.StartButton.clicked.connect(self.MassageOpen)                 # 开始爬取、数据填入、错误信息显示
        self.ui.FileChooseButton.clicked.connect(self.OpenFile)               # 获取文件夹路径
        self.ui.SingleFileChooseButton.clicked.connect(self.OpenSingleFile)   # 获取单文件路径
        self.m_seleCrawlerThread.WarningSignal.connect(self.ThreadWarningMassage)  # 错误信息提示框
        self.m_seleCrawlerThread.FinishSignal.connect(self.MainWindowSlot)    # 线程运行结束信号
        self.ui.FoldercheckBox.clicked.connect(self.FolderChoose)             # 单文件、文件夹切换
        self.ui.FireFoxChooseButton.clicked.connect(self.OpenFireFoxPath)     # 获取火狐浏览器路径
        self.ui.AboutButton.clicked.connect(self.AboutMassage)                # About 窗口

    """
    复选框槽函数
    选中表示启用打开文件夹窗口,进行全文件操作
    """
    def FolderChoose(self):
        if self.ui.FoldercheckBox.checkState() == QtCore.Qt.CheckState.Checked:
            self.IsFolderChoose = True
            self.ui.FilePathEdit.show()
            self.ui.FileChooseButton.show()
            self.ui.SingleFilePathEdit.hide()
            self.ui.SingleFileChooseButton.hide()
        else:
            self.IsFolderChoose = False
            self.ui.FilePathEdit.hide()
            self.ui.FileChooseButton.hide()
            self.ui.SingleFilePathEdit.show()
            self.ui.SingleFileChooseButton.show()

    """
    About
    """
    def AboutMassage(self):
        QMessageBox.information(self,'操作提示',("<p><b>软件操作</b> </p>"
               "<p>1.填入(更新)电脑中的火狐浏览器位置(因为要调动浏览器不同电脑中浏览器路径不同目前需要查找填入) </p>"
               "<p>2.输入编码系统登录网址(网址会实时更改，请注意更新) </p>"
               "<p>3.文件夹复选框选中则代表进行文件夹操作，选择文件夹操作软件会批量处理(默认为单文件操作) </p>"
               "<p>4.点击文件选择(可选项有:*.xls文件、*.xlsx文件、包含前两种文件的文件夹) </p>"
               "<p>5.点击开始 </p>"
               "<p>----------------------------</p>"
               "<p>在运行结束后会弹出空信息界面，其中的信息为： </p>"
               "<p>文件夹名 -> 文件名 -> Sheet名 -> 数据为空的Item号 </p>"
               "<p>数据可保存为Excel表格 </p>"
               "<p><b>注意：</b>期间会出现一个弹窗，为浏览器调用接口请在表格处理完后关闭。</p>"))

    """
    错误提示窗口
    """
    def ThreadWarningMassage(self,iWarnValue):
        if iWarnValue == 1:
            QMessageBox.warning(self,'Warning','访问登录页面失败，请重新输入登录地址，可能地址改变!')
        elif iWarnValue == 2:
            QMessageBox.warning(self,'Warning','没有查到PDM号数据！')

    """
    爬取、写入任务开始
    """
    def MassageOpen(self):
        # 判错
        if self.ui.URLEdit.text() == "":
            QMessageBox.warning(self,'Warning','请输入登录登录网址')
            return 
        self.labelLoad.show()                                                 # 等待动画加载
        self.m_seleCrawlerThread.SetLoginUrl(self.ui.URLEdit.text(),self.ui.FireFoxPathEdit.text())          # 将登录网址传入线程类
        self.m_seleCrawlerThread.start()                                      # 线程开始

    """
    单文件操作
    """
    def OpenSingleFile(self):
        # 初始化
        self.m_seleCrawlerThread.Initial()
        self.m_MassageWindow.Initial()
        self.iFilePath,ok = QFileDialog.getOpenFileName(self,'请打开推荐表文件','./',"*.xlsx;;*.xls")
        if ok:
            pattern=re.compile("/")                              
            result = pattern.split(self.iFilePath)
            self.m_MassageWindow.SetFolderName(result[-2])  # 将文件夹名传入错误信息类，用于根节点名称
            self.ui.SingleFilePathEdit.setText(self.iFilePath)
            # 是否为文件夹选择
            self.m_seleCrawlerThread.IsChooseFolder(self.IsFolderChoose)
            if self.iFilePath[-1] == 's':
                self.Xls2Xlsx(self.iFilePath)
                os.remove(self.iFilePath)
                self.m_seleCrawlerThread.SetFilePath(self.iFilePath+'x')
            else:
                # 文件路径传输
                self.m_seleCrawlerThread.SetFilePath(self.iFilePath)

    """
    获取线程结束信号,显示错误信息
    """
    def MainWindowSlot(self,iBoolValue:bool):
        if iBoolValue == True:
            self.labelLoad.hide()                                             # 等待动画结束
            # self.m_MassageWindow = self.m_seleCrawlerThread.GetMassageClass() # 获取子窗口指针
            self.m_MassageWindow.ShowMassage()                                # 数据显示
            self.m_MassageWindow.show()                                       # 错误信息界面显示
            self.m_seleCrawlerThread.quit()                                   # 线程结束
            self.m_seleCrawlerThread.wait()
    
    """
    XLS2XLSX
    """
    def Xls2Xlsx(self,iFilePath):
        try:
            app = xw.App(visible=False,add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(iFilePath)
            wb.save(iFilePath+"x")
            wb.close()
            app.quit()
            return
        except:
            QMessageBox.warning(self,'Warning','文件损坏或文件异常请手动转换\n'+iFilePath)

    """
    火狐浏览器选择
    """
    def OpenFireFoxPath(self):
        self.iFireFoxPath,ok = QFileDialog.getOpenFileName(self,'请打开推荐表文件','./',"*.exe")
        if ok:
            self.ui.FireFoxPathEdit.setText(self.iFireFoxPath)

    """
    打开文件夹,提取该文件夹下所有文件的地址
    """
    def OpenFile(self):
        # 初始化
        self.m_seleCrawlerThread.Initial()
        self.m_MassageWindow.Initial()
        # 打开文件夹
        cur_dir = QDir.currentPath()                                                # 文件夹地址类
        self.iFilePath = QFileDialog.getExistingDirectory(self,'打开文件夹',cur_dir)  # 获取文件夹路径窗口
        # 正则表达式获取文件夹名称
        pattern=re.compile("/")                              
        result = pattern.split(self.iFilePath)
        self.m_MassageWindow.SetFolderName(result[-1])  # 将文件夹名传入错误信息类，用于根节点名称
        # self.m_MassageWindow # 需要压入文件夹名
        if self.iFilePath != None:
            self.ui.FilePathEdit.setText(self.iFilePath)
            # 是否为文件夹选择
            self.m_seleCrawlerThread.IsChooseFolder(self.IsFolderChoose)
            XlsFile=glob(self.iFilePath+"/*xls")            # 提取后缀为xls的文件地址
            XlsxFile = glob(self.iFilePath+"/*xlsx")        # 提取后缀为xlsx的文件地址
            oExcelPathList = []                              # 临时变量:保存文件地址
            for file in XlsFile:
                self.Xls2Xlsx(file)
                os.remove(file)
                oExcelPathList.append(file.replace("\\","/"))
            XlsNum = len(oExcelPathList)
            self.m_seleCrawlerThread.SetXlsNum(XlsNum)       # 将Xls文件的数量传给线程类(有两种格式的文件需要两种操作方式)
            for xfile in XlsxFile:
                oExcelPathList.append(xfile.replace("\\","/"))
            self.m_seleCrawlerThread.SetFilePathList(oExcelPathList)  # 将文件地址传输至线程类
            print(oExcelPathList)

"""
主函数运行
"""
if __name__ =='__main__':
    # 用于适应高分辨率
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling) 
    app = QApplication(sys.argv)
    MainWindow = MainWidget() # 实例化
    MainWindow.setWindowIcon(QIcon('./resources/robot.ico')) # 设置图标
    MainWindow.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, True) # 设置最小化按钮
    MainWindow.show() # 界面展示
    sys.exit(app.exec_())
