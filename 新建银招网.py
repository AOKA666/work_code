import requests
from xlrd import open_workbook
import os


def main(uname, email):
    url = "http://mail.yinzhaowang.com:8010/Users/edit"
    data = {
        'email': email,
        'forwardemail': 'acca-rz02@yinzhaowang.com',
        'uname': uname,
        'tel': '',
        'active': 1,
        'password': 'Gaodun123',
        'password2': 'Gaodun123',
        '_method': 'put',
        '_forward': '%2FUsers',
        }
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36',
        'Cookie': cookie
        }
    r = requests.post(url, headers=headers, data=data)
    if r.status_code == 200:
        if "邮箱地址已存在" in r.text:
            print("{}   邮箱地址已存在".format(email))
        else:
            print("{}   新建成功".format(email))


if __name__ == '__main__':
    workbook = open_workbook("待建银招网.xls")
    sheet = workbook.sheets()[0]
    n = sheet.nrows
    print("共有{}个".format(n-1))
    cookie = input('请输入cookie:')
    for i in range(1, n):
        uname = sheet.cell(i,0).value
        email = sheet.cell(i,4).value
        main(uname, email)
    print("全部新建成功！")
    os.system("pause")
