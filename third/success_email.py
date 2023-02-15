import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
import base64
from bs4 import BeautifulSoup
import xlrd
import re
import time
import traceback
from threading import Thread, Lock
import random
from concurrent.futures import ThreadPoolExecutor


class SendMail():
    def __init__(self, lst, username, sender, password, ms, tool):
        self.smtpserver = "smtp.exmail.qq.com"
        self.lst = lst
        self.username = username
        self.sender = sender
        self.password = password
        self.ms = ms
        self.tool = tool
        self.lock = Lock()
        self.target = 0
        
        
    def ContentFIA(self, m, new_account, new_email):

        #附件
        pdfFile = 'third/高顿教育ACCA学员专用：My ACCA Account 操作指南（2020年更新）.pdf'
        pdfApart = MIMEApplication(open(pdfFile,'rb').read())
        pdfApart.add_header('Content-Disposition','attachment',filename='ACCA账户操作指南最新版（高顿教育ACCA学员专用）.pdf')
        m.attach(pdfApart)
        #添加正文
        try:
            self.lock.acquire()
            with open('third/FIAHtml.txt') as f:
                Maintext = f.read()
            soup = BeautifulSoup(Maintext,'html.parser')

            #替换账号和邮箱
            tag = soup.find_all('strong')
            s = ''
            for i in tag:
                s = s + i.text
            account = ''.join(re.findall(r'\d{7}',s))
            email = ''.join(re.findall(r'[a-zA-Z0-9z_-]+@[a-zA-Z0-9z_-]+\.[com,cn,net]{1,3}',s))
            first = Maintext.replace(account,new_account)
            final = first.replace(email,new_email)
            
            new_soup = BeautifulSoup(final,'html.parser')
            with open('third/FIAHtml_copy.txt','w')as fp:
                fp.write(new_soup.prettify())
            with open('third/FIAHtml_copy.txt') as fp:
                f_new = fp.read()            
        finally:
            self.lock.release()
        return f_new

    def ContentACCA(self, m, new_account, new_email, new_exe):

        #附件
        pdfFile = 'third/高顿教育ACCA学员专用：My ACCA Account 操作指南（2020年更新）.pdf'
        pdfApart = MIMEApplication(open(pdfFile,'rb').read())
        pdfApart.add_header('Content-Disposition','attachment',filename='ACCA账户操作指南最新版（高顿教育ACCA学员专用）.pdf')
        m.attach(pdfApart)

        #添加正文
        try:
            self.lock.acquire()
            with open('third/ACCAHtml.txt') as f:
                Maintext = f.read()
            soup = BeautifulSoup(Maintext,'html.parser')

            #替换免试数
            info = soup.find('span',style='color:#D250DA')
            amend = ''.join(re.findall(r'[fF].*\d',info.text))
            if 'F' in new_exe:
                exe = ''.join(re.findall(r'[fF].*\d',new_exe))
                Maintext = Maintext.replace(amend,exe)
            else:
                Maintext = Maintext.replace(info.text,'无免试。')

            #替换账号和邮箱
            tag = soup.find_all('strong')
            s = ''
            for i in tag:
                s = s + i.text
            account = ''.join(re.findall(r'\d{7}',s))
            email = ''.join(re.findall(r'[a-zA-Z0-9z_-]+@[a-zA-Z0-9z_-]+\.[com,cn,net]{1,3}',s))
            first = Maintext.replace(account,new_account)
            final = first.replace(email,new_email)

            new_soup = BeautifulSoup(final,'html.parser')
            with open('third/ACCAHtml_copy.txt','w')as fp:
                fp.write(new_soup.prettify())
            with open('third/ACCAHtml_copy.txt') as fp:
                f_new = fp.read()
        finally:
            self.lock.release()
        return f_new

    @classmethod
    def excel_data(cls, file):
        lst = []
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheets()[0]
        sheet2 = workbook.sheets()[1]
        nrows = sheet.nrows
        for i in range(1,nrows):
            table = sheet.row_values(i)
            lst.append(table)
        return lst 

    def send_mail(self, m, receiver, subject, account, email):    
        m['Subject'] = subject
        m['From'] = self.sender
        m['To'] = receiver
        smtp = smtplib.SMTP_SSL(host=self.smtpserver)
        smtp.connect(host=self.smtpserver, port=465)
        try:
            smtp.login(self.username, self.password)
        except:
            self.ms.text_print2.emit(self.tool.ui.output_3, "邮箱登录失败")
            return ''
        smtp.sendmail(self.sender, receiver, m.as_string())
        smtp.quit()


    def main(self, data):
        receiver = data[6] #收件人        
        if 'FIA' in data[1]:                
            subject = 'ACCA注册账号审核通过-FIA-' + data[0]
            m = MIMEMultipart('mixed')
            htmlFIA = MIMEText(self.ContentFIA(m,str(int(data[2])),data[4]),'html','utf-8')
            m.attach(htmlFIA)
            try:
                self.send_mail(m, receiver, subject, str(int(data[2])), data[4])
            except:
                self.ms.text_print2.emit(self.tool.ui.output_3, "异常错误,可能邮箱{}不存在".format(data[6]))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                return ''
            self.ms.text_print.emit(self.tool.ui.output_3, '{}----->>>发送完成'.format(data[0]))
            self.target += 1
            self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
        elif 'ACCA' in data[1]:
            subject = 'ACCA注册账号审核通过-ACCA-' + data[0]
            m = MIMEMultipart('mixed')
            htmlACCA = MIMEText(self.ContentACCA(m,str(int(data[2])),data[4],str(data[3])),'html','utf-8')
            m.attach(htmlACCA)
            try:
                self.send_mail(m, receiver, subject, str(int(data[2])), data[4])
            except:
                self.ms.text_print2.emit(self.tool.ui.output_3, "异常错误,可能邮箱{}不存在".format(data[6]))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                return ''
            self.ms.text_print.emit(self.tool.ui.output_3, '{}----->>>发送完成,耶！'.format(data[0]))
            self.target += 1
            self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
            

    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("邮件开始发送"))
        self.ms.text_print.emit(self.tool.ui.output_3, "共有{}个".format(len(self.lst)))
        start_time = time.perf_counter()
        with ThreadPoolExecutor(max_workers=4) as t:
            t.map(self.main, [self.lst[i] for i in range(len(self.lst))])
        end_time = time.perf_counter()
        self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("邮件发送完成"))
        self.ms.text_print.emit(self.tool.ui.output_3, '共用时{:.2f}秒！'.format(end_time-start_time))
        self.ms.text_print.emit(self.tool.ui.output_3, "程序已退出")


        
