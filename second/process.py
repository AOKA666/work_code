import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import base64
from bs4 import BeautifulSoup
import xlrd
import re
import time
import traceback


class ProcessEmail():
    def __init__(self, lst, username, password, ms, tool):
        self.lst = lst
        self.username = username
        self.password = password
        self.ms = ms
        self.tool = tool
        self.smtpserver = 'smtp.exmail.qq.com'

        
    def Content(self, new_account,new_email):
        
        f = open('second/Process.txt').read()
        soup = BeautifulSoup(f,'html.parser')
        tag = soup.find_all('strong')
        s = ''
        for i in tag:
            s = s + i.text
        account = ''.join(re.findall(r'\d{7}',s))
        email_text = tag[1].text
        email = ''.join(re.findall(r'[a-zA-Z0-9z_-]+@[a-zA-Z0-9z_-]+\.[com,cn,net]{1,3}',s))
        first = f.replace(account,new_account)
        final = first.replace(email,new_email)
        new_soup = BeautifulSoup(final,'html.parser')
        with open('second/Process_copy.txt','w')as fp:
            fp.write(new_soup.prettify())
        f_new = open('second/Process_copy.txt').read()
        text = f_new
        return text


    @classmethod
    def excel_data(cls, file):
        lst = []
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheets()[0]
        nrows = sheet.nrows
        for i in range(1,nrows):
            table = sheet.row_values(i)
            lst.append(table)
        return lst


    def send_mail(self, msg, subject, sender, receiver, username, password):
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = receiver

        smtp = smtplib.SMTP_SSL(host=self.smtpserver)
        smtp.connect(host=self.smtpserver, port=465)
        try:
            smtp.login(username, password)
        except:
            self.ms.text_print2.emit(self.tool.ui.output_2, "邮箱登录失败")
            return ''
        smtp.sendmail(sender, receiver,msg.as_string())
        smtp.quit()


    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_2, "{:=^20}".format("开始发送邮件"))
        self.ms.text_print.emit(self.tool.ui.output_2, '共有{}个'.format(len(self.lst)))
        start_time = time.perf_counter()        
        for i in range(len(self.lst)):
            username = self.username #发送邮箱
            password = self.password #邮箱密码
            sender = username
            receiver = self.lst[i][6]  #收件人
            subject = 'ACCA学员代注册等待审核通知-' + self.lst[i][0]
            try:
                msg = MIMEMultipart('related')
                html = MIMEText(self.Content(str(int(self.lst[i][2])),self.lst[i][4]),'html','utf-8')
                msg.attach(html)
                self.send_mail(msg, subject, sender, receiver, username, password)
                self.ms.text_print.emit(self.tool.ui.output_2, '{}-->完成'.format(self.lst[i][0]))
                self.ms.ui_show2.emit(self.tool.ui.progressBar_2, i+1)
            except Exception as e:
                self.ms.text_print2.emit(self.tool.ui.output_2, '啊！出现未知错误')
                self.ms.ui_show2.emit(self.tool.ui.progressBar_2, i+1)
        end_time = time.perf_counter()
        self.ms.text_print.emit(self.tool.ui.output_2, "{:=^20}".format("邮件发送完成")) 
        self.ms.text_print.emit(self.tool.ui.output_2, '本次共用时{:.2f}秒！'.format(end_time-start_time))
        self.ms.text_print.emit(self.tool.ui.output_2, '程序退出') 

    
