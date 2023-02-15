from PySide2.QtWidgets import QApplication, QTextBrowser, QMessageBox
from PySide2.QtGui import QIcon
from PySide2.QtUiTools import QUiLoader
from PySide2 import QtWidgets
import time
import os
import threading
from PySide2.QtCore import QThread, Signal, QObject
from threading import Lock
from xlrd import open_workbook, xldate_as_tuple
import warnings
import datetime


warnings.filterwarnings("ignore", "(?s).*MATPLOTLIBDATA.*", category=UserWarning)
class MySignals(QObject):

    text_print = Signal(QTextBrowser, str)
    text_print2 = Signal(QTextBrowser, str)
    text_clear = Signal(QTextBrowser)
    reset_bar = Signal(QTextBrowser)
    ui_show = Signal(QTextBrowser, int)
    ui_show2 = Signal(QTextBrowser, int)
    
        
class WebdriverTest():

    def __init__(self):
        self.ui = QUiLoader().load('test.ui')
               
        self.ui.buttonGroup.buttonClicked.connect(self.react)        
        self.ui.pushButton.clicked.connect(self.work)

        ms.text_print.connect(self.printf)
        ms.text_print2.connect(self.printff)
        ms.text_clear.connect(self.textClear)
        ms.reset_bar.connect(self.reset)
        ms.ui_show.connect(self.setProg)
        ms.ui_show2.connect(self.update)

        self.ui.MessageBox = QMessageBox()


    def printf(self, fb, mystr):
        fb.append("<font color='blue'>"+str(mystr)+"</font>")
        fb.ensureCursorVisible()


    def printff(self, fb, mystr):
        fb.append("<font color='red'>"+str(mystr)+"</font>")
        fb.ensureCursorVisible()


    def textClear(self, fb):
        fb.clear()


    def reset(self, bar):
        bar.reset()


    def setProg(self, bar, num):
        bar.setRange(0, num)


    def update(self, bar, num):
        bar.setValue(num)


    def react(self):
        if self.ui.radioButton_1.isChecked():
            return self.ui.radioButton_1.text()
        elif self.ui.radioButton_2.isChecked():
            return self.ui.radioButton_2.text()
        elif self.ui.radioButton_3.isChecked():
            return self.ui.radioButton_3.text()
        elif self.ui.radioButton_4.isChecked():
            return self.ui.radioButton_4.text()
        elif self.ui.radioButton_5.isChecked():
            return self.ui.radioButton_5.text()
        elif self.ui.radioButton_6.isChecked():
            return self.ui.radioButton_6.text()
        else:
            return ''

    
    def work(self):
        signal = self.react()
        print(signal)
        if '1' in signal:
            self.thread = Operation_1()
            self.thread.start()
        elif '2' in signal:
            self.thread = Operation_2()
            self.thread.start()
        elif '3' in signal:
            self.thread = Operation_3()
            self.thread.start()
        elif '4' in signal:
            self.thread = Operation_4()
            self.thread.start()
        elif '5' in signal:
            self.thread = Operation_5()
            self.thread.start()
        elif '6' in signal:
            self.thread = Operation_6()
            self.thread.start()
        else:
            self.ui.MessageBox.critical(self.ui, "错误", "请选择一种操作")
        

class RunThread(QThread):

    trigger = Signal()

    def __init__(self, parent=None):
        super(RunThread, self).__init__()
        
    def __del__(self):
        self.wait()


class Operation_1(RunThread):

    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.reset_bar.emit(tool.ui.progressBar)
        ms.text_clear.emit(tool.ui.output_1)
        ms.text_print.emit(tool.ui.output_1, "{:=^20}".format("开始导出信息"))
        from first.download import FM
        username = crm_account
        try:
            password = str(int(crm_password))
        except:
            password = crm_password
        teacher = token.split('_')[1]
        f = FM(username, password, teacher, token, ms, tool)
        f.get_info()
        ms.text_print.emit(tool.ui.output_1, "{:=^20}".format("导出信息完成"))
        ms.text_print.emit(tool.ui.output_1, "程序退出")
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()
        

class Operation_2(RunThread):

    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.reset_bar.emit(tool.ui.progressBar)
        ms.text_clear.emit(tool.ui.output_1)
        from first.send import FM
        username = mail_account
        password = mail_password
        subject = mail_subject
        fm = FM(username, password, subject, ms, tool)
        fm.run()
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()


class Operation_3(RunThread):

    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.reset_bar.emit(tool.ui.progressBar)
        ms.text_clear.emit(tool.ui.output_1)
        from first.write import CRM
        username = crm_account
        try:
            password = str(int(crm_password))
        except:
            password = crm_password
        workbook = open_workbook('代注册.xls')
        sheet = workbook.sheets()[0]
        n = sheet.nrows
        ms.ui_show.emit(tool.ui.progressBar, n)
        result = []
        for i in range(n):
            name = sheet.cell(i, 0).value
            email = sheet.cell(i,1).value
            result.append([name, email])
        crm = CRM(username, password, n, result, token, ms, tool)
        crm.run()
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()
        

class Operation_4(RunThread):

    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.text_clear.emit(tool.ui.output_2)
        from second.process import ProcessEmail
        lst = ProcessEmail.excel_data("待审核.xls")
        ms.reset_bar.emit(tool.ui.progressBar_2)
        ms.ui_show.emit(tool.ui.progressBar_2, len(lst))
        username = mail_account
        password = mail_password
        pe = ProcessEmail(lst, username, password, ms, tool)
        pe.run()
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()

        
class Operation_5(RunThread):

    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.text_clear.emit(tool.ui.output_3)
        from third.success_email import SendMail
        lst = SendMail.excel_data('审核通过.xls') # 读取表格
        ms.reset_bar.emit(tool.ui.progressBar_3)
        ms.ui_show.emit(tool.ui.progressBar_3, len(lst))
        password = mail_password
        username = mail_account
        sender = username
        mail = SendMail(lst, username, sender, password, ms, tool)
        mail.run()
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()

        
class Operation_6(RunThread):
    def run(self):
        tool.ui.pushButton.setEnabled(False)
        ms.text_clear.emit(tool.ui.output_3)
        user = crm_account
        try:
            pw = str(int(crm_password))
        except:
            pw = crm_password
        wb = open_workbook("审核通过.xls")
        sheet = wb.sheets()[0]
        n = sheet.nrows
        # 重置进度条
        ms.reset_bar.emit(tool.ui.progressBar_3)
        if n >= 2:
            ms.ui_show.emit(tool.ui.progressBar_3, n-1)
        result = []
        for i in range(1, n):
            name = sheet.cell(i, 0).value
            try:
                mobile = str(int(sheet.cell(i, 5).value))
            except:
                mobile = sheet.cell(i, 5).value
            email = sheet.cell(i, 6).value
            shenfen = sheet.cell(i, 1).value
            account = sheet.cell(i, 2).value
            regi_email = sheet.cell(i, 4).value
            date = sheet.cell(i, 7).value
            date2 = datetime.datetime(*xldate_as_tuple(date, 0))
            dob = date2.strftime('%Y-%m-%d')
            regi_addr = sheet.cell(i, 8).value
            date = sheet.cell(i, 10).value
            date2 = datetime.datetime(*xldate_as_tuple(date, 0))
            regi_date = date2.strftime('%Y-%m-%d')
            result.append([name,mobile,email,shenfen,account,regi_email,dob,regi_addr,regi_date])
        from third.success_crm import WriteCRM
        crm = WriteCRM(user, pw, n, result, tool, ms, token)       
        crm.run()
        tool.ui.pushButton.setEnabled(True)
        self.trigger.emit()

    
if __name__ == '__main__':
    info = open_workbook("hidden/账户信息.xls")
    mail = info.sheets()[0]
    crm = info.sheets()[1]
    mail_account = mail.cell(1,0).value
    mail_password = mail.cell(1,1).value
    mail_subject = mail.cell(1,2).value
    crm_account = crm.cell(1,0).value
    crm_password = crm.cell(1,1).value
    token = crm.cell(1,2).value
    

    ms = MySignals()
    app_main = QApplication([])
    app_main.setWindowIcon(QIcon('logo.jpg'))
    tool = WebdriverTest()
    tool.ui.show()    
    app_main.exec_()
