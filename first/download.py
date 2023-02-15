import requests
import json
import xlwt
import datetime

class FM:
    def __init__(self, username, password, teacher, token, ms, tool):
        self.username = username
        self.password = password
        self.teacher = teacher
        self.token = token
        self.ms = ms
        self.tool = tool
        
        
    def get_info(self):
        login_session = requests.Session()
        headers = {
        'referer': 'https://eds.gaodun.com/',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) \
            Chrome/84.0.4147.105 Safari/537.36'
        }
        kv = {
        'userName': self.username,
        'password': self.password,
        'GDSID': 'R0tPWVZISEFvS1ZiWVdGZW1DdmZXRUVWR2NQVnZSa28=',
        }
        url = "https://ssm.gaodun.com/api/member_login/vLogin"
        # 第一次爬取，Session()获得缓存Cookies等

        r = login_session.get(url, params=kv, headers=headers)
        result = str(r.content, 'utf-8')
        result2 = json.loads(result)
        # 获取token值
        token = result2['data']['session']
        print(token)

        # 获取日期
        today = datetime.date.today()
        before = today - datetime.timedelta(days=30)
        page = 1
        search_url = 'https://ssm.gaodun.com/api/Authentication/index'
        search_kv = {
            'status': '2012959',
            'project': '1000053',
            'enterTime[]': (before, today),
            'teacherID[]': self.teacher,
            'status_stage': '100118200',
            'access_token': token,
            'pageSize': '30',
            'p': page,
            'CourseType': '1000053',
        }
        # 按页码搜索
        ls = []
        while page:
            # 第二次爬取，获取所有学员信息
            fmlist = login_session.get(search_url, params=search_kv, headers=headers)
            if fmlist == '':
                break
            fmdata = str(fmlist.content, encoding='utf-8')
            try:
                json_loads_file = json.loads(fmdata)['data']['lists']
            except:
                self.ms.text_print2.emit(self.tool.ui.output_1, "爬取失败，可能未正常登录")
                self.ms.text_print2.emit(self.tool.ui.output_1, "请重试")
                return ''
            for i in json_loads_file:
                ls.append(i['member_id'])
            # 判断单页数量是否到达上限30    
            if len(json_loads_file) == 30:  
                # 超过的话页码加一，再搜索一次
                page += 1
                search_kv['p'] = page
            # 没有超过就结束循环    
            else:                           
                break
        if ls == '':
            return ''
        # 第三次爬取，根据member_id搜索学员的姓名和邮箱
        lt = []
        for i in ls:
            email_url = 'https://ssm.gaodun.com/api/student/studentPersonsalData'
            email_kv = {
                'access_token': token,
                'member_id': i
            }
            emList = login_session.get(email_url, params=email_kv, headers=headers)
            emData = str(emList.content, encoding='utf-8')
            email = json.loads(emData)['data']['email']
            name = json.loads(emData)['data']['real_name']
            lt.append([name, email])        # lt储存所有学员的姓名和邮箱
        # 打开一个表格写入数据
        wb = xlwt.Workbook()
        sheet = wb.add_sheet("学员信息")
        if len(lt) >= 1:
            self.ms.ui_show.emit(self.tool.ui.progressBar, len(lt))
        for i in range(len(lt)):
            sheet.write(i, 0, lt[i][0])
            sheet.write(i, 1, lt[i][1])
            self.ms.ui_show2.emit(self.tool.ui.progressBar, i+1)        
        sheet.col(1).width = 256*30        
        wb.save("代注册.xls")
    
