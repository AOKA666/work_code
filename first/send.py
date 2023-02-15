import smtplib
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import time



class FM:
    def __init__(self, user, password, subject, ms, tool):
        self.user = user
        self.password = password
        self.subject = subject
        self.ms = ms
        self.tool = tool
        self.target = 0
        
        
    def send(self, msg, email):
        smtpserver = 'smtp.exmail.qq.com'
        
        server = smtplib.SMTP(host=smtpserver)
        server.login(self.user, self.password)
        server.sendmail(self.user, email, msg.as_string())
        server.quit()


    def mail(self):
        # 附件：注册信息确认表
        excel = 'first/新ACCA学员注册信息确认表 (高顿财经).xls'
        excelApart = MIMEApplication(open(excel, 'rb').read())
        excelApart.add_header('Content-Disposition', 'attachment', filename='新ACCA学员注册信息确认表 (高顿财经).xls')
        # 邮件正文读取
        html = open("first/FM.txt", encoding='utf-8').read()
        text = MIMEText(html, 'html', 'utf-8')
        # 根据表格发邮件
        wb = open_workbook('代注册.xls', formatting_info=True)
        sheet = wb.sheets()[0]
        # 为之后填写crm做准备
        wb2 = copy(wb)
        sheet2 = wb2.get_sheet(0)
        num = sheet.nrows
        self.ms.ui_show.emit(self.tool.ui.progressBar, num)
        for i in range(num):
            email = sheet.cell(i, 1).value
            name = sheet.cell(i, 0).value
            subject = "{} {}".format(self.subject, name)
            if email == "":
                self.ms.text_print.emit(self.tool.ui.output_1, "{}的邮箱不存在!".format(name))
                continue
            msg = MIMEMultipart()
            # 添加正文和附件
            msg.attach(text)
            msg.attach(excelApart)
            # 添加头部
            msg['Subject'] = Header(subject, 'utf-8').encode()
            msg['From'] = Header(self.user)
            msg['To'] = Header(email)
            try:
                self.send(msg, email)
                self.ms.text_print.emit(self.tool.ui.output_1, "{}-->成功".format(name))
                # 填写excel
                style = self.excel_write(4) 
                sheet2.write(i, 5, "邮件发送成功", style)
                self.ms.ui_show2.emit(self.tool.ui.progressBar, i+1)
                wb2.save("代注册.xls")
            except Exception as e:
                self.ms.text_print2.emit(self.tool.ui.output_1, "{}-->邮箱格式错误".format(name))
                style = self.excel_write(2)
                sheet2.write(i, 5, "邮箱格式错误", style)
                self.ms.ui_show2.emit(self.tool.ui.progressBar, i+1)
                wb2.save("代注册.xls")


    def excel_write(self,index):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.colour_index = index
        style.font = font
        return style


    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("开始发送代注册邮件"))
        start = time.perf_counter()
        self.mail()
        end = time.perf_counter()
        self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("代注册邮件发送完成"))
        self.ms.text_print.emit(self.tool.ui.output_1, "共用时{:.2f}秒".format(end-start))
        self.ms.text_print.emit(self.tool.ui.output_1, "程序退出")

    
