'''

文件对话框：QFileDialog

'''

import sys,os
import xlrd
from docx import Document
from docxtpl import DocxTemplate

from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import pyqtSignal


class mainWindow(QWidget):
    def __init__(self):
        super(mainWindow, self).__init__()
        self.docxPath = ''
        self.xlsxPath = ''
        self.variableList = []#变量列表
        self.sourceDataList = []#来源数据列表
        self.loadSheetNamesList = []#加载Sheet名称列表
        self.VarAssInf = {}
        self.setupUi(self)

        self.show()
    def setupUi(self, groupBox_3):
        if not groupBox_3.objectName():
            groupBox_3.setObjectName(u"groupBox_3")
        groupBox_3.resize(445, 294)
        self.groupBox = QGroupBox(groupBox_3)
        self.groupBox.setObjectName(u"groupBox")
        self.groupBox.setGeometry(QRect(10, 10, 421, 91))
        self.label_4 = QLabel(self.groupBox)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(10, 20, 101, 31))
        self.label_4.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.lineEdit_3 = QLineEdit(self.groupBox)
        self.lineEdit_3.setObjectName(u"label_3")
        self.lineEdit_3.setGeometry(QRect(120, 20, 221, 31))
        self.lineEdit_3.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.pushButton_2 = QPushButton(self.groupBox)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(340, 20, 75, 31))
        self.pushButton_3 = QPushButton(self.groupBox)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setGeometry(QRect(250, 60, 161, 23))
        self.groupBox_2 = QGroupBox(groupBox_3)
        self.groupBox_2.setObjectName(u"groupBox_2")
        self.groupBox_2.setGeometry(QRect(10, 110, 421, 131))
        self.label = QLabel(self.groupBox_2)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 20, 101, 31))
        self.label.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.lineEdit_2 = QLineEdit(self.groupBox_2)
        self.lineEdit_2.setObjectName(u"label_2")
        self.lineEdit_2.setGeometry(QRect(120, 20, 221, 31))
        self.lineEdit_2.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.pushButton = QPushButton(self.groupBox_2)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(340, 20, 71, 31))
        self.pushButton_4 = QPushButton(self.groupBox_2)
        self.pushButton_4.setObjectName(u"pushButton_4")
        self.pushButton_4.setGeometry(QRect(230, 60, 181, 23))
        #关联变量与引用的数据按钮
        self.pushButton_5 = QPushButton(self.groupBox_2)
        self.pushButton_5.setObjectName(u"pushButton_5")
        self.pushButton_5.setGeometry(QRect(10, 90, 160, 31))
        #保存的配置文件
        self.saveINIButton = QPushButton(self.groupBox_2)
        self.saveINIButton.setGeometry(QRect(180, 90, 116, 31))
        self.openINIButton = QPushButton(self.groupBox_2)
        self.openINIButton.setGeometry(QRect(296, 90, 116, 31))
        #测试生成与开始
        self.pushButton_6 = QPushButton(groupBox_3)
        self.pushButton_6.setObjectName(u"pushButton_6")
        self.pushButton_6.setGeometry(QRect(160, 250, 271, 31))
        self.pushButton_8 = QPushButton(groupBox_3)
        self.pushButton_8.setObjectName(u"pushButton_8")
        self.pushButton_8.setGeometry(QRect(10, 250, 141, 31))
        QWidget.setTabOrder(self.pushButton_2, self.pushButton_3)
        QWidget.setTabOrder(self.pushButton_3, self.pushButton)
        QWidget.setTabOrder(self.pushButton, self.pushButton_4)
        QWidget.setTabOrder(self.pushButton_4, self.pushButton_5)
        QWidget.setTabOrder(self.pushButton_5, self.pushButton_6)

        self.retranslateUi(groupBox_3)

        QMetaObject.connectSlotsByName(groupBox_3)

        self.setWindowTitle('签证批量处理工具V1.0.2')
        self.pushButton_2.clicked.connect(self.loadDocx)
        self.pushButton.clicked.connect(self.loadXlsx)
        self.pushButton_5.clicked.connect(self.associatedVariableClick)
        self.pushButton_8.clicked.connect(self.testGenerationClick)
        self.pushButton_6.clicked.connect(self.batchProcessingClick)
        self.saveINIButton.clicked.connect(self.saveConfigure)
        self.openINIButton.clicked.connect(self.openConfigure)
    # setupUi

    def retranslateUi(self, groupBox_3):
        groupBox_3.setWindowTitle(QCoreApplication.translate("groupBox_3", u"GroupBox", None))
        groupBox_3.setWindowTitle("")
        self.groupBox.setTitle(QCoreApplication.translate("groupBox_3", u"\u6a21\u677f", None))
        self.label_4.setText(QCoreApplication.translate("groupBox_3", u"\u6a21\u677f\uff08*.docx\uff09", None))
        self.lineEdit_3.setText(QCoreApplication.translate("groupBox_3", u"\u9009\u62e9\u6587\u4ef6", None))
        self.pushButton_2.setText(QCoreApplication.translate("groupBox_3", u"\u9009\u62e9\u6587\u4ef6", None))
        self.pushButton_3.setText(QCoreApplication.translate("groupBox_3", u"\u70b9\u6b64\u67e5\u770b\u76f8\u5173\u6a21\u677f\u7f16\u8f91\u8bf4\u660e", None))
        self.groupBox_2.setTitle(QCoreApplication.translate("groupBox_3", u"\u6570\u636e\u6e90", None))
        self.label.setText(QCoreApplication.translate("groupBox_3", u"\u6570\u636e\u6e90\uff08*.xlsx\uff09", None))
        self.lineEdit_2.setText(QCoreApplication.translate("groupBox_3", u"\u9009\u62e9\u6587\u4ef6", None))
        self.pushButton.setText(QCoreApplication.translate("groupBox_3", u"\u9009\u62e9\u6587\u4ef6", None))
        self.pushButton_4.setText(QCoreApplication.translate("groupBox_3", u"\u70b9\u6b64\u67e5\u770b\u76f8\u5173\u6570\u636e\u6e90\u7f16\u8f91\u8bf4\u660e", None))
        self.pushButton_5.setText(QCoreApplication.translate("groupBox_3", u"关联变量与引用的数据", None))
        self.saveINIButton.setText('保存配置')
        self.openINIButton.setText('打开配置')
        self.pushButton_6.setText(QCoreApplication.translate("groupBox_3", u"\u5f00\u59cb\u6279\u91cf\u5904\u7406\u7b7e\u8bc1\u6587\u4ef6", None))
        self.pushButton_8.setText('测试生成十页签证单')

    def loadDocx(self):#加载docx
        fname,_ = QFileDialog.getOpenFileName(self,'选择文件','.','docx模板(*.docx);;所有文件(*)')

        if fname == '':
            self.lineEdit_3.setText('选择文件')
        else:
            self.lineEdit_3.setText(fname)
            self.docxPath = fname

    def loadXlsx(self):#加载xlsx
        fname, _ = QFileDialog.getOpenFileName(self, '选择文件', '.', '数据源(*.xlsx);;所有文件(*)')
        if fname == '':
            self.lineEdit_2.setText('选择文件')
        else:
            try:
                self.lineEdit_2.setText(fname)
                #print('fname:',fname)
                self.xlsxPath = fname
                data = xlrd.open_workbook(fname)
                #print('所有工作表的名称：',data.sheet_names())
                self.SheetChoice = SheetChoice(self,data.sheet_names())
                self.SheetChoice.returnSignal.connect(self.getSheets)
                self.SheetChoice.show()
                self.SheetChoice.exec_()
            except Exception as e:
                import traceback
                print('traceback.format_exc():\n%s' % traceback.format_exc())
                print(e)
            #print(data.sheet_names())

    #关联变量与引用数据被点击事件
    def associatedVariableClick(self):
        try:
            if self.docxPath != '' and self.loadSheetNamesList != []:
                docx = DocxTemplate(self.docxPath)
                self.variableList = docx.getVariable()#获取变量列表
                print('变量列表',self.variableList)

                #获取来源数据名称列表
                xlsx = xlrd.open_workbook(self.xlsxPath)
                xlsxSheet = xlsx.sheet_by_name(self.loadSheetNamesList[0])
                while True:
                    if xlsx.sheet_loaded(self.loadSheetNamesList[0]):
                        break
                    else:
                        pass
                #effectivenRows = xlsxSheet.nrows #有效行数
                effectivenNcols = xlsxSheet.ncols   #有效列数
                print('有效列数：',effectivenNcols)
                self.sourceDataList = []
                i = 0
                while i < effectivenNcols:
                    #print(i)
                    self.sourceDataList.append(xlsxSheet.cell_value(0,i))
                    #print('列名：',xlsxSheet.cell_value(0,i))
                    i = i + 1

                self.sourceDataList.append('**不关联任何数据**')#增加一个选项，选择即跳过关联

                print('表格数据名称',self.sourceDataList)

                #开始关联数据UI

                self.associatedVariable = associatedVariable(self,self.variableList,self.sourceDataList)
                self.associatedVariable.returnSignal.connect(self.getVarAssInf)
                self.associatedVariable.show()
                self.associatedVariable.exec_()
            else:
                if self.docxPath == '':
                    msg_box = QMessageBox(QMessageBox.Warning, '错误', 'docx文件目录不能为空！')
                    msg_box.exec_()
                elif self.loadSheetNamesList == []:
                    msg_box = QMessageBox(QMessageBox.Warning, '错误', 'xlsx文件目录或需要加载Sheet不能为空！')
                    msg_box.exec_()
        except:
            import traceback
            print(traceback.format_exc())

    #loadXlsx事件回调，将sheet列表保存到self.loadSheetNamesList
    def getSheets(self,SheetList):
        print('选择加载的表有：',SheetList)
        self.loadSheetNamesList = SheetList

    #associatedVariableClick事件完成回调事件将关联字典保存到self.VarAssInf
    def getVarAssInf(self, VarAssInf):#变量关联信息
        self.VarAssInf = VarAssInf
        print('关联变量信息',self.VarAssInf)

    #测试生成被点击事件
    def testGenerationClick(self):
        try:
            if self.VarAssInf != {}:

                # 创建零时目录
                from tempfile import TemporaryDirectory
                self.temporaryDirectory = TemporaryDirectory()
                print(self.temporaryDirectory.name)
                    #self.temporaryDirectory = dirname
                    #print(self.temporaryDirectory)

                # 获取来源数据名称列表
                xlsx = xlrd.open_workbook(self.xlsxPath)
                xlsxSheet = xlsx.sheet_by_name(self.loadSheetNamesList[0])
                while True:
                    if xlsx.sheet_loaded(self.loadSheetNamesList[0]):
                        break
                    else:
                        pass
                effectivenRows = xlsxSheet.nrows  # 有效行数
                # effectivenNcols = xlsxSheet.ncols  # 有效列数
                self.sourceDataList = []
                num = 0

                for i in range(0, 10):
                    if num < effectivenRows:
                        context = {}
                        for variableName in self.variableList:
                            if self.VarAssInf[variableName][1] == '文本':
                                context[variableName] = str(
                                    xlsxSheet.cell_value(i + 1, self.VarAssInf[variableName][0]))
                            elif self.VarAssInf[variableName][1] == '整数':
                                context[variableName] = int(
                                    xlsxSheet.cell_value(i + 1, self.VarAssInf[variableName][0]))
                            elif self.VarAssInf[variableName][1] == '浮点数':
                                context[variableName] = float(
                                    xlsxSheet.cell_value(i + 1, self.VarAssInf[variableName][0]))
                        docx = DocxTemplate(self.docxPath)
                        docx.render(context)
                        docx.save(self.temporaryDirectory.name + "\签证单%s.docx"%num)  # 保存
                    num = num + 1
                curpath = os.path.realpath(__file__)
                print('当前文件夹路径：',os.path.dirname(curpath))
                merge('签证单',range(0, 10),self.temporaryDirectory.name,os.path.dirname(curpath))

                self.temporaryDirectory.cleanup()

                msg_box = QMessageBox(QMessageBox.Warning, '完成', '测试文件生成完毕，储存位置为程序目录！')
                msg_box.exec_()

            else:
                msg_box = QMessageBox(QMessageBox.Warning, '警告', '模板中的变量相关信息配置为空，我们将不会为其传入任何值（数据表中的任何信息在这种情况下不会有作用），但我们任然会执行下一步操作，这是为了执行您在docx模板中书写的'\
                                                                 'Jinja2语句（如果有的话）。')
                msg_box.exec_()
        except:
            import traceback
            print(traceback.format_exc())

    #批量处理签证文件被点击事件
    def batchProcessingClick(self):
        #批量处理
        try:
            if self.VarAssInf != {}:

                # 获取来源数据名称列表
                xlsx = xlrd.open_workbook(self.xlsxPath)
                xlsxSheet = xlsx.sheet_by_name(self.loadSheetNamesList[0])
                while True:
                    if xlsx.sheet_loaded(self.loadSheetNamesList[0]):
                        break
                    else:
                        pass
                effectivenRows = xlsxSheet.nrows  # 有效行数
                effectivenNcols = xlsxSheet.ncols  # 有效列数
                self.saveDirectorySelectionDialog = saveDirectorySelectionDialog(self,effectivenRows,effectivenNcols)
                self.saveDirectorySelectionDialog.returnSignal.connect(self.taskPageClick)
                self.saveDirectorySelectionDialog.show()
                self.saveDirectorySelectionDialog.exec_()
            else:
                msg_box = QMessageBox(QMessageBox.Warning, '警告',
                                      '模板中的变量相关信息配置为空，我们将不会为其传入任何值（数据表中的任何信息在这种情况下不会有作用），但我们任然会执行下一步操作，这是为了执行您在docx模板中书写的' \
                                      'Jinja2语句（如果有的话）。')
                msg_box.exec_()
                self.saveDirectorySelectionDialog = saveDirectorySelectionDialog(self, '数据表的数据不会起作用', '数据表的数据不会起作用')
                self.saveDirectorySelectionDialog.returnSignal.connect(self.taskPageClick)
                self.saveDirectorySelectionDialog.show()
                self.saveDirectorySelectionDialog.exec_()
        except Exception as e:
            import traceback
            print('traceback.format_exc():\n%s' % traceback.format_exc())
            print(e)

    #开始任务被点击事件
    def taskPageClick(self,pathDict):
        try:
            self.pathDict = pathDict
            self.taskPage = taskPageUI(self)
            self.taskPage.show()
            self.thread = thread(self)
            self.thread.start()
            self.taskPage.exec_()

        except:
            import traceback
            print(traceback.format_exc())

    #保存配置
    def saveConfigure(self):
        try:
            if self.docxPath == '':
                msg_box = QMessageBox(QMessageBox.Warning, '错误', '请选择docx模板文件路径！')
                msg_box.exec_()
            elif self.xlsxPath == '':
                msg_box = QMessageBox(QMessageBox.Warning, '错误', '请选择xlsx数据源文件路径！')
                msg_box.exec_()
            elif self.loadSheetNamesList == []:
                msg_box = QMessageBox(QMessageBox.Warning, '错误', '请检查当前xlsx文件及sheet配置！')
                msg_box.exec_()
            elif self.VarAssInf == {}:
                msg_box = QMessageBox(QMessageBox.Warning, '警告', '变量与数据源信息关联的配置为空！这可能是您在docx模板中使用set语句设置了所有变量的值，导致不需要为他们关联任何值，也可能是您没有点击关联按钮进行关联，之后程序将\
                                                                 继续执行生成配置文件的命令但如果您为上述第二种情况，则会导致程序崩溃！')
                msg_box.exec_()
                fname, _ = QFileDialog.getSaveFileName(self, "选取文件夹", "./", "所有支持文件 (*.visaData);;所有文件 (*)")
                if fname == '':
                    msg_box = QMessageBox(QMessageBox.Warning, '错误', '您必须选择保存位置！')
                    msg_box.exec_()
                else:
                    print(fname)
                    docx = open(self.docxPath, 'rb')
                    docxdata = docx.read()
                    docx.close()
                    xlsx = open(self.xlsxPath, 'rb')
                    xlsxdata = xlsx.read()
                    xlsx.close()
                    # do_config = HandleConfig(fname)
                    datas = {
                        'setting': {
                            'docxPath': self.docxPath,
                            'xlsxPath': self.xlsxPath,
                            'loadSheetNamesList': self.loadSheetNamesList,
                            'VarAssInf': self.VarAssInf
                        }
                    }
                    result = HandleConfig.write_config(datas, fname)
                    if result:
                        msg_box = QMessageBox(QMessageBox.Warning, '成功', '配置文件保存成功\n下一次您可以通过打开配置文件快速完成前期设置！')
                        msg_box.exec_()
            else:
                fname, _ = QFileDialog.getSaveFileName(self, "选取文件夹", "./", "所有支持文件 (*.visaData);;所有文件 (*)")
                if fname == '':
                    msg_box = QMessageBox(QMessageBox.Warning, '错误', '您必须选择保存位置！')
                    msg_box.exec_()
                else:
                    print(fname)
                    docx = open(self.docxPath,'rb')
                    docxdata = docx.read()
                    docx.close()
                    xlsx = open(self.xlsxPath, 'rb')
                    xlsxdata = xlsx.read()
                    xlsx.close()
                    #do_config = HandleConfig(fname)
                    datas = {
                        'setting' : {
                            'docxPath':self.docxPath,
                            'xlsxPath':self.xlsxPath,
                            'loadSheetNamesList':self.loadSheetNamesList,
                            'VarAssInf':self.VarAssInf
                        }
                    }
                    result = HandleConfig.write_config(datas,fname)
                    if result:
                        msg_box = QMessageBox(QMessageBox.Warning, '成功', '配置文件保存成功\n下一次您可以通过打开配置文件快速完成前期设置！')
                        msg_box.exec_()
        except:
            import traceback
            print(traceback.format_exc())

    #读取配置
    def openConfigure(self):
        try:
            fname, _ = QFileDialog.getOpenFileName(self, "选取配置文件", "./", "所有支持文件 (*.visaData);;所有文件 (*)")
            if fname == '':
                msg_box = QMessageBox(QMessageBox.Warning, '错误', '您没有选择配置文件！')
                msg_box.exec_()
            else:
                print(fname)
                separatorNumList = []
                num = 0
                for i in fname:
                    if i == '/':
                        separatorNumList.append(num)
                    num = num + 1
                import zipfile
                r = zipfile.is_zipfile(fname)
                if r:
                    fz = zipfile.ZipFile(fname, 'r')
                    unziping = True
                    for file in fz.namelist():
                        print(fname[0:separatorNumList[-1]]+'/'+file)
                        fz.extract(file,fname[0:separatorNumList[-1]])
                        unziping = False

                    while unziping:
                        pass

                    confirmBox = QMessageBox(QMessageBox.Question, '加载方式', '您是否希望接下来可以更改此配置？')  # 创建一个确认框
                    yes = confirmBox.addButton(' 是，我然后要更改 ', QMessageBox.YesRole)
                    no = confirmBox.addButton(' 否，我不会尝试更改 ', QMessageBox.NoRole)
                    confirmBox.exec_()
                    if confirmBox.clickedButton() == yes:

                        setting = HandleConfig(fname[0:separatorNumList[-1]] + '/value.ini')  # 读取那个文件
                        print(fname[0:separatorNumList[-1]] + '/value.ini')
                        self.loadSheetNamesList = setting.get_eval_data('setting', "loadSheetNamesList")  # 读取什么内容

                        self.VarAssInf = setting.get_value('setting', "VarAssInf")  # 读取什么内容

                        self.lineEdit_2.setText(fname[0:separatorNumList[-1]] + '/data.xlsx')
                        self.lineEdit_3.setText(fname[0:separatorNumList[-1]] + '/Template.docx')
                        self.docxPath = fname[0:separatorNumList[-1]] + '/Template.docx'
                        self.xlsxPath = fname[0:separatorNumList[-1]] + '/data.xlsx'

                        os.remove(fname[0:separatorNumList[-1]] + '/value.ini')

                        self.profileModification = [True,fz.namelist(),'%s'%fname[0:separatorNumList[-1]]] #配置文件修改

                        msg_box = QMessageBox(QMessageBox.Warning, '成功','加载配置文件成功\n\r加载方式:读写\n\r注：您本次做出的更改会保存为%s和%s，您可重新将其打包为新配置！'\
                                              %(self.docxPath,self.xlsxPath))
                        msg_box.exec_()

                    elif confirmBox.clickedButton() == no:


                        setting = HandleConfig(fname[0:separatorNumList[-1]] + '/value.ini')  # 读取那个文件
                        self.loadSheetNamesList = setting.get_eval_data('setting', "loadSheetNamesList")  # 读取什么内容
                        self.VarAssInf = setting.get_value('setting', "VarAssInf")  # 读取什么内容

                        self.docxPath = fname[0:separatorNumList[-1]] + '/Template.docx'
                        self.xlsxPath = fname[0:separatorNumList[-1]] + '/data.xlsx'

                        os.remove(fname[0:separatorNumList[-1]] + '/value.ini')

                        self.lineEdit_2.setText('正在使用配置文件（不可更改）：' + fname)
                        self.lineEdit_3.setText('正在使用配置文件（不可更改）：' + fname)

                        self.profileModification = [False,fz.namelist(),'%s'%fname[0:separatorNumList[-1]]]  # 配置文件修改

                        msg_box = QMessageBox(QMessageBox.Warning, '成功', '加载配置文件成功\n\r加载方式:只读\n\r注：您本次做出的更改在此次程序结束后会自动清除')
                        msg_box.exec_()

                else:
                    msg_box = QMessageBox(QMessageBox.Warning, '错误', '配置文件损坏！')
                    msg_box.exec_()
        except:
            import traceback
            print(traceback.format_exc())

    def closeEvent(self, event):
        try:
            if self.profileModification[0] == True:
                pass
            else:
                # 删除文件
                complete = False
                for file in self.profileModification[1]:
                    os.remove(self.profileModification[2] + '/' + file)
                    print(self.profileModification[2] + '/' + file)
                    complete = True
            while True:
                if complete:
                    break
        except:
            pass
        self.close()

class thread(QThread):  # 子线程执行写入
    mysignal = pyqtSignal(tuple)  # 创建一个自定义信号，元组参数

    def __init__(self,parent=None):
        super(thread, self).__init__(parent)
        self.parent = parent

    def run(self):
        try:
            #创建临时目录
            from tempfile import TemporaryDirectory
            self.parent.temporaryDirectory = TemporaryDirectory()
            print(self.parent.temporaryDirectory.name)

            # 获取来源数据名称列表
            num = 0
            xlsx = xlrd.open_workbook(self.parent.xlsxPath)

            for sheet in self.parent.loadSheetNamesList:
                xlsxSheet = xlsx.sheet_by_name(sheet)
                self.parent.taskPage.textBrowser.append('开始处理Sheet %s' % sheet)  # 文本框逐条添加数据
                self.parent.taskPage.textBrowser.moveCursor(self.parent.taskPage.textBrowser.textCursor().End)  # 文本框显示到底部
                count = 0
                print('重置count为0')
                while True:
                    if xlsx.sheet_loaded(sheet):
                        break
                    else:
                        pass
                effectivenRows = xlsxSheet.nrows  # 有效行数
                while True:
                    #print('effectivenRows:', effectivenRows)
                    if count + 1 < effectivenRows:
                        context = {}
                        for variableName in self.parent.variableList:
                            print('count+1', count + 1)
                            print('num:',num)
                            try:
                                if self.parent.VarAssInf[variableName][1] == '文本':
                                    context[variableName] = str(
                                        xlsxSheet.cell_value(count + 1, self.parent.VarAssInf[variableName][0]))
                                elif self.parent.VarAssInf[variableName][1] == '整数':
                                    context[variableName] = int(
                                        xlsxSheet.cell_value(count + 1, self.parent.VarAssInf[variableName][0]))
                                elif self.parent.VarAssInf[variableName][1] == '浮点数':
                                    context[variableName] = float(
                                        xlsxSheet.cell_value(count + 1, self.parent.VarAssInf[variableName][0]))
                            except:
                                continue


                        docx = DocxTemplate(self.parent.docxPath)
                        print('docx.render(context)最终替换数据：context为',context)

                        import jinja2
                        try:
                            docx.render(context)
                        except jinja2.exceptions.UndefinedError:
                            import traceback
                            msg_box = QMessageBox(QMessageBox.Warning, '错误', traceback.format_exc()+'\n可能原因：\n若提示错误类似于jinja2.exceptions.UndefinedError: number（代指您用的某个变量名） is undefined，那么可能是您没有为其传值尝试在docx\
                                                         模板中使用{% set 变量名 = 值 %}或关联变量是不要选“**不关联任何数据**”项' )
                            msg_box.exec_()
                            break


                        self.parent.taskPage.textBrowser.append('替换数据 %s' % context)  # 文本框逐条添加数据
                        self.parent.taskPage.textBrowser.moveCursor(
                            self.parent.taskPage.textBrowser.textCursor().End)  # 文本框显示到底部
                        docx.save(self.parent.temporaryDirectory.name + "\%s%s.docx" % (
                        self.parent.pathDict['fileName'], num))  # 保存
                        self.parent.taskPage.textBrowser.append('保存文件 %s' % (self.parent.temporaryDirectory.name + "\%s%s.docx" % (self.parent.pathDict['fileName'], num)))  # 文本框逐条添加数据
                        self.parent.taskPage.textBrowser.moveCursor(self.parent.taskPage.textBrowser.textCursor().End)  # 文本框显示到底部

                        count = count + 1
                        num = num + 1
                    else:
                        break
            merge(self.parent.pathDict['fileName'],range(0,num),self.parent.temporaryDirectory.name,self.parent.pathDict['storagePath'])
            msg_box = QMessageBox(QMessageBox.Warning, '完成', '%s/%s生成完成'%(self.parent.pathDict['storagePath'],self.parent.pathDict['fileName']))
            msg_box.exec_()
            #self.parent.temporaryDirectory.cleanup()

        except:
            import traceback
            print(traceback.format_exc())
class saveDirectorySelectionDialog(QDialog):#保存目录对话框
    #关联模板中的变量与引用的数据
    returnSignal = pyqtSignal(dict)
    def __init__(self, parent=None,effectivenRows=0,effectivenNcols=0):
        super(saveDirectorySelectionDialog, self).__init__(parent)
        # 设置标题与初始窗口大小
        self.setWindowTitle('关联模板中的变量与引用的数据')
        self.resize(400, 220)
        self.effectivenRows = effectivenRows
        self.effectivenNcols = effectivenNcols
        self.returnDict = {}#保存路径

        # 阻塞父类窗口不能点击
        self.setWindowModality(Qt.ApplicationModal)

        self.UI()

    def UI(self):
        try:
            self.saveDirectorySelectionUI = saveDirectorySelectionUI(self)
            self.saveDirectorySelectionUI.label_2.setText(str(self.effectivenRows))
            self.saveDirectorySelectionUI.label_4.setText(str(self.effectivenNcols))
            self.saveDirectorySelectionUI.pushButton_2.clicked.connect(self.selectDirectory)
            self.saveDirectorySelectionUI.pushButton.clicked.connect(self.starMission)
        except Exception as e:
            import traceback
            print('traceback.format_exc():\n%s' % traceback.format_exc())
            print(e)

    def selectDirectory(self):
        try:
            fname, _ = QFileDialog.getSaveFileName(self, '保存文件', '.', 'docx模板(*.docx);;所有文件(*)')
            self.saveDirectorySelectionUI.lineEdit.setText(fname)

            dividerList = []
            num = 0
            for i in fname:
                if i == '/':
                    dividerList.append(num)
                num = num + 1
            #print(fname[0:dividerList[-1]])
            self.returnDict['storagePath'] = fname[0:dividerList[-1]]
            self.returnDict['fileName'] = fname[dividerList[-1]+1:]
        except:
            import traceback
            print(traceback.format_exc())
    def starMission(self):
        try:
            if self.returnDict != {}:
                self.returnSignal.emit(self.returnDict)
                self.close()
            else:
                msg_box = QMessageBox(QMessageBox.Warning, '错误', '保存目录不能为空！')
                msg_box.exec_()
        except:
            import traceback
            print(traceback.format_exc())

class associatedVariable(QDialog):#关联变量对话框
    #关联模板中的变量与引用的数据
    returnSignal = pyqtSignal(dict)
    def __init__(self, parent=None,variableList=[],sourceDataList=[]):
        super(associatedVariable, self).__init__(parent)
        # 设置标题与初始窗口大小
        self.setWindowTitle('关联模板中的变量与引用的数据')
        self.resize(440, 275)

        # 阻塞父类窗口不能点击
        self.setWindowModality(Qt.ApplicationModal)

        self.variableList = variableList#变量列表
        self.sourceDataList = sourceDataList#来源数据名称列表
        print(self.variableList)

        self.UI()

    def UI(self):
        try:
            self.topFiller = QWidget(self)
            self.topFiller.setMinimumSize(420,len(self.variableList)*55)  #######设置滚动条的尺寸

            num = 0
            for i in self.variableList:
                self.addItem(num,i)
                num = num + 1
            pass

            self.scroll = QScrollArea()
            self.scroll.setWidget(self.topFiller)
            self.vbox = QVBoxLayout()
            self.vbox.setContentsMargins(0, 0, 0, 0)

            self.confirmButton = QPushButton(self)
            self.confirmButton.setText('确认关联对应变量与数据')
            self.confirmButton.clicked.connect(self.confirm)
            self.vbox.addWidget(self.confirmButton)

            self.vbox.addWidget(self.scroll)
            self.setLayout(self.vbox)
            #self.a = Ui_associatedVariableItem(self.topFiller)
        except Exception as e:
            import traceback
            print('traceback.format_exc():\n%s' % traceback.format_exc())
            print(e)

    def quit(self):
        self.close()
    def addItem(self,nameNum,text):
        exec('self.Item%s = associatedVariableItem(self.topFiller)'%(nameNum))
        exec('self.Item = self.Item%s'%nameNum)
        self.Item.move(0, 56 * nameNum+1)
        self.Item.label_1.setText(text)
        self.Item.comboBox.addItems(self.sourceDataList)

    def confirm(self):
        try:
            #print('confirm')
            num = 0
            self.returnDict = {}
            for i in self.variableList:
                #print(i)
                exec('self.Item = self.Item%s' % num)
                print('数据表变量名称：',self.Item.comboBox.currentIndex(),self.Item.comboBox.currentText(),self.Item.comboBox_2.currentText())

                # 去除选择 **不关联任何数据** 选项的变量
                if self.Item.comboBox.currentText() == '**不关联任何数据**':
                    print('跳过'+i)
                    continue

                self.returnDict[i] = [self.Item.comboBox.currentIndex(),self.Item.comboBox_2.currentText()]

                print('最终返回的变量列表与其关联信息',self.returnDict)

                num = num + 1
            print(self.returnDict)
            self.returnSignal.emit(self.returnDict)
            self.close()
        except:
            import traceback
            print(traceback.format_exc())

class taskPageUI(QDialog):#任务开始UI
    def __init__(self, parent=None,variableList=[],sourceDataList=[]):
        super(taskPageUI, self).__init__(parent)
        # 设置标题与初始窗口大小
        self.setWindowTitle('关联模板中的变量与引用的数据')
        self.resize(400, 300)

        # 阻塞父类窗口不能点击
        self.setWindowModality(Qt.ApplicationModal)

        self.setupUi()

    def setupUi(self):
        try:
            GroupBox = QGroupBox(self)
            GroupBox.resize(400, 300)
            GroupBox.move(0, 0)
            self.textBrowser = QTextBrowser(GroupBox)
            self.textBrowser.setObjectName(u"textBrowser")
            self.textBrowser.setGeometry(QRect(0, 30, 401, 121))
            self.label_1 = QLabel(GroupBox)
            self.label_1.setObjectName(u"label_1")
            self.label_1.setGeometry(QRect(0, 0, 401, 31))
            self.label_1.setStyleSheet(u"background-color: rgb(177, 177, 177)")
            self.tabWidget = QTabWidget(GroupBox)
            self.tabWidget.setObjectName(u"tabWidget")
            self.tabWidget.setGeometry(QRect(0, 150, 401, 150))
            self.tab = QWidget()
            self.tab.setObjectName('tab')
            self.textBrowser_2 = QTextBrowser(self.tab)
            self.textBrowser_2.setObjectName(u"textBrowser_2")
            self.textBrowser_2.setGeometry(QRect(0, 0, 401, 130))
            self.tabWidget.addTab(self.tab, "")
            self.tab_2 = QWidget()
            self.tab_2.setObjectName(u"tab_2")
            self.tabWidget.addTab(self.tab_2, "")

            self.retranslateUi(GroupBox)

            QMetaObject.connectSlotsByName(GroupBox)
        except:
            import traceback
            print(traceback.format_exc())
    # setupUi

    def retranslateUi(self, GroupBox):
        GroupBox.setWindowTitle(QCoreApplication.translate("GroupBox", u"GroupBox", None))
        self.label_1.setText(QCoreApplication.translate("GroupBox", u"\u8be6\u7ec6\u4fe1\u606f\uff1a", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), QCoreApplication.translate("GroupBox",'值为空', None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), QCoreApplication.translate("GroupBox", u"Tab 2", None))
    # retranslateUi
class associatedVariableItem(QGroupBox):#关联变量对话框单项UI
    def __init__(self, parent=None):
        super(associatedVariableItem, self).__init__(parent)
        self.setupUi(self)

    def setupUi(self, GroupBox):
        GroupBox.resize(410, 55)
        # 变量名标签
        self.label_1 = QLabel(GroupBox)
        self.label_1.setObjectName(u"label_1")
        self.label_1.setGeometry(QRect(10, 10, 81, 31))
        self.label_1.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.comboBox = QComboBox(GroupBox)
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(QRect(150, 10, 140, 31))
        self.comboBox_2 = QComboBox(GroupBox)
        self.comboBox_2.addItem("")
        self.comboBox_2.addItem("")
        self.comboBox_2.addItem("")
        self.comboBox_2.setObjectName(u"comboBox_2")
        self.comboBox_2.setGeometry(QRect(340, 10, 60, 31))
        self.label_3 = QLabel(GroupBox)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(290, 10, 51, 31))
        self.label_3.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.label_2 = QLabel(GroupBox)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(100, 10, 51, 31))
        self.label_2.setStyleSheet(u"background-color: rgb(177, 177, 177)")

        self.retranslateUi(GroupBox)

        QMetaObject.connectSlotsByName(GroupBox)
    # setupUi

    def retranslateUi(self, GroupBox):
        GroupBox.setWindowTitle(QCoreApplication.translate("GroupBox", u"GroupBox", None))
        self.label_1.setText("")
        self.comboBox_2.setItemText(0, QCoreApplication.translate("GroupBox", u"\u6587\u672c", None))
        self.comboBox_2.setItemText(1, QCoreApplication.translate("GroupBox", u"\u6574\u6570", None))
        self.comboBox_2.setItemText(2, QCoreApplication.translate("GroupBox", u"\u6d6e\u70b9\u6570", None))

        self.label_3.setText(QCoreApplication.translate("GroupBox", u"\u6570\u636e\u7c7b\u578b", None))
        self.label_2.setText(QCoreApplication.translate("GroupBox", u"\u5339\u914d\u6570\u636e", None))
    # retranslateUi


class SheetChoice(QDialog):#Sheet选择对话框
    # 加载选中Sheet
    returnSignal = pyqtSignal(list)
    def __init__(self, parent=None,SheetNameList=[]):
        super(SheetChoice, self).__init__(parent)
        # 设置标题与初始窗口大小
        self.setWindowTitle('选择加载的Sheet')
        self.resize(320, 275)

        self.SheetNameList = SheetNameList

        # 阻塞父类窗口不能点击
        self.setWindowModality(Qt.ApplicationModal)

        self.UI()

    def UI(self):
        try:
            self.topFiller = QWidget(self)
            self.topFiller.setMinimumSize(300, len(self.SheetNameList)*21)  #######设置滚动条的尺寸

            num = 0
            for i in self.SheetNameList:
                exec('self.checkBox%s = QCheckBox(self.topFiller)' % num)
                exec('self.checkBox = self.checkBox%s' % num)
                self.checkBox.resize(240,20)
                self.checkBox.move(0,num*21)
                self.checkBox.setText(i)

                if  num == 0:
                    self.checkBox.setChecked(True)

                num = num + 1
            pass

            self.scroll = QScrollArea()
            self.scroll.setWidget(self.topFiller)
            self.vbox = QVBoxLayout()
            self.vbox.setContentsMargins(3, 3, 3, 3)

            self.confirmButton = QPushButton(self)
            self.confirmButton.setText('加载选中Sheet（排名靠前的Sheet会优先被加载）')
            self.confirmButton.clicked.connect(self.confirm)
            self.vbox.addWidget(self.confirmButton)
            self.vbox.addWidget(self.scroll)
            self.setLayout(self.vbox)

        except Exception as e:
            import traceback
            print('traceback.format_exc():\n%s' % traceback.format_exc())
            print(e)
    def confirm(self):
        self.returnList = []

        num = 0
        for i in self.SheetNameList:
            exec('self.checkBox = self.checkBox%s' % num)
            if self.checkBox.isChecked() == True:
                #print(i)
                self.returnList.append(i)
            num = num + 1
        self.returnSignal.emit(self.returnList)

        self.close()

class saveDirectorySelectionUI(QGroupBox):#选择保存目录的UI设计
    #保存目录对话框
    def __init__(self, parent=None):
        super(saveDirectorySelectionUI, self).__init__(parent)
        self.setupUi(self)
    def setupUi(self, GroupBox):
        if not GroupBox.objectName():
            GroupBox.setObjectName(u"GroupBox")
        GroupBox.resize(400, 220)
        self.label_1 = QLabel(GroupBox)
        self.label_1.setObjectName(u"label_1")
        self.label_1.setGeometry(QRect(0, 0, 101, 31))
        self.label_1.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.label_3 = QLabel(GroupBox)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(0, 40, 101, 31))
        self.label_3.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.label_2 = QLabel(GroupBox)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(110, 0, 291, 31))
        self.label_2.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.label_4 = QLabel(GroupBox)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(110, 40, 291, 31))
        self.label_4.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.label = QLabel(GroupBox)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(60, 70, 341, 31))
        self.label.setTextFormat(Qt.AutoText)
        self.label_5 = QLabel(GroupBox)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(0, 120, 101, 31))
        self.label_5.setStyleSheet(u"background-color: rgb(177, 177, 177)")
        self.lineEdit = QLineEdit(GroupBox)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(100, 119, 221, 31))
        self.pushButton_2 = QPushButton(GroupBox)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(320, 120, 81, 31))
        self.pushButton = QPushButton(GroupBox)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(0, 170, 401, 31))

        self.retranslateUi(GroupBox)

        QMetaObject.connectSlotsByName(GroupBox)
    # setupUi

    def retranslateUi(self, GroupBox):
        GroupBox.setWindowTitle(QCoreApplication.translate("GroupBox", u"GroupBox", None))
        self.label_1.setText(QCoreApplication.translate("GroupBox", u"\u6709\u6548\u884c\uff1a\uff08\u884c\uff09", None))
        self.label_3.setText(QCoreApplication.translate("GroupBox", u"\u6709\u6548\u5217\uff1a\uff08\u5217\uff09", None))
        self.label_2.setText("")
        self.label_4.setText("")
        self.label.setText(QCoreApplication.translate("GroupBox", u"\u8bf7\u6838\u5bf9\u6709\u6548\u884c\u548c\u6709\u6548\u5217\u6570\uff0c\u5c3d\u91cf\u786e\u4fdd\u6ca1\u6709\u8bfb\u53d6\u591a\u4f59\u7a7a\u767d\u884c\u4e0e\u5217\uff01", None))
        self.label_5.setText(QCoreApplication.translate("GroupBox", u"\u4fdd\u5b58\u76ee\u5f55\uff1a", None))
        self.lineEdit.setText(QCoreApplication.translate("GroupBox", u"\u9009\u62e9\u4fdd\u5b58\u76ee\u5f55", None))
        self.pushButton_2.setText(QCoreApplication.translate("GroupBox", u"\u9009\u62e9\u76ee\u5f55", None))
        self.pushButton.setText(QCoreApplication.translate("GroupBox", u"\u5f00\u59cb\u4efb\u52a1", None))
    # retranslateUi
class merge:#合并docx文档
    def __init__(self,fileName,fileNum,filePath,storagePath):
        # 合并文档的列表
        files = []
        for i in fileNum:
            if os.path.isfile(r'%s\%s%s.docx' % (filePath,fileName,i)):
                files.append(r'%s\%s%s.docx' % (filePath,fileName,i))
            else:
                print('没有文件：', r'E:\word\result\test%s.docx' % i)
        # 合并操作
        print(files)

        self.combine_word_documents(files,fileName,storagePath)

    def combine_word_documents(self,files,fileName,storagePath):
        try:
            merged_document = Document()

            for index, file in enumerate(files):
                sub_doc = Document(file)

                # Don't add a page break if you've reached the last file.
                if index < len(files) - 1:
                    sub_doc.add_page_break()

                for element in sub_doc.element.body:
                    merged_document.element.body.append(element)

            merged_document.save('%s\%s' % (storagePath, fileName))
            print('完毕')
        except:
            pass


# 封装读取保存配置文件功能
from configparser import ConfigParser
class HandleConfig:
    """
    配置文件读写数据的封装
    """
    def __init__(self, filename):
        """
        :param filename: 配置文件名
        """
        self.filename = filename
        self.config = ConfigParser()  # 读取配置文件1.创建配置解析器
        self.config.read(self.filename, encoding="utf-8")  # 读取配置文件2.指定读取的配置文件
        ''' self.filename = filename
        with open(self.filename,'r',encoding="utf-8") as f:
            self.text = f.read()'''

    # get_value获取所有的字符串，section区域名, option选项名
    def get_value(self, section, option):
        def get_value(self, section, option):
            return self.config.get(section, option)
        '''import pickle
        data = pickle.load(self.text)
        return '''

    # get_int获取整型，section区域名, option选项名
    def get_int(self, section, option):
        return self.config.getint(section, option)

    # get_float获取浮点数类型，section区域名, option选项名
    def get_float(self, section, option):
        return self.config.getfloat(section, option)

    # get_boolean（译：比例恩）获取布尔类型，section区域名, option选项名
    def get_boolean(self, section, option):
        return self.config.getboolean(section, option)

    # get_eval_data 获取列表，section区域名, option选项名
    def get_eval_data(self, section, option):
        return eval(self.config.get(section, option))  # get 获取后为字符串，再用 eval 转换为列表

    @staticmethod
    def write_config(datas, filename):
        """
        写入配置操作
        :param datas: 需要传入写入的数据
        :param filename: 指定文件名
        :return:
        """
        # 做校验，为嵌套字典的字典才可以（意思.隐私.谈.ce)
        if isinstance(datas, dict):  # 遍历，在外层判断是否为字典
            # 再来判断内层的 values 是否为字典
            for value in datas.values():    # 先取出value
                if not isinstance(value, dict):     # 在判断
                    return False

            config = ConfigParser()
            for key in datas:
                config[key] = datas[key]

            os.mkdir(filename[0:-9]+'visaData')
            with open(filename[0:-9]+'visaData/value.ini', "w",encoding="utf-8") as file:
                config.write(file)

            from shutil import copyfile
            copyfile(datas['setting']['docxPath'],filename[0:-9]+'visaData/Template.docx')
            copyfile(datas['setting']['xlsxPath'], filename[0:-9] + 'visaData/data.xlsx')

            import zipfile
            visaDataZip = zipfile.ZipFile(filename, 'w')
            for file in os.listdir(filename[0:-9]+'visaData'):
                print('file',file)
                print("filename[0:-9]+'visaData'",filename[0:-9]+'visaData')
                print("filename[0:-9]+'visaData/' + file",filename[0:-9]+'visaData/' + file)
                visaDataZip.write(filename[0:-9]+'visaData/' + file,file)
            visaDataZip.close()
            #清除目录
            import shutil
            shutil.rmtree(filename[0:-9]+'visaData')
            return True


# do_config = HandleConfig('testcase.conf')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = mainWindow()
    sys.exit(app.exec_())