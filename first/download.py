import requests
import json
import xlwt
import datetime

class FM:
    def __init__(self, username, password, ms, tool):
        self.username = username
        self.password = password
        self.ms = ms
        self.tool = tool
        
        
    def get_info(self):        
        # 登录，主要获得access_token
        login_data = {
        # "GDSID": "Xgnnu-czWXl-zNL_ln7nYv8_nAog_m0lLfL1Vabc",
        "appid": 210666,
        "password": self.password,
        "user": self.username,
        }
        url = 'https://apigateway.gaodun.com/api/v4/vigo/login'
        
        # 获取access token
        login = requests.post(url, data=login_data).json()
        user_id = login['result']['user_id']
        access_token = login['accessToken']
        
        # 将 access_token 写入 headers
        headers = {
        "Authentication": "Basic "+access_token,
        "Content-Type": "application/json;charset=UTF-8",
        "Host": "apigateway.gaodun.com",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        }

        # 获取日期
        today = datetime.date.today()
        before = today - datetime.timedelta(days=30)
        page = 1
        search_url = 'https://apigateway.gaodun.com/olaf/api/v1/auth/list'
        data = {
            "projectId": 1000053,
            "statusStage": 100118200,
            "status": 2012959,
            "teacherIdList": [user_id],
            "enterTimeStart": str(before),
            "enterTimeEnd": str(today),
            "studentName": "",
            "pageNum": page,
            "pageSize": 100
            }
        # 按页码搜索
        ls = []
        while page:
            # 第二次爬取，获取所有学员信息
            # data 需要是json 格式，因为 contentType 是 json
            fmlist = requests.post(search_url, data=json.dumps(data), headers=headers)
            if fmlist == '':
                break
            fmdata = str(fmlist.content, encoding='utf-8')
            try:
                json_loads_file = json.loads(fmdata)['result']['list']
            except:
                self.ms.text_print2.emit(self.tool.ui.output_1, "爬取失败，可能未正常登录")
                self.ms.text_print2.emit(self.tool.ui.output_1, "请重试")
                return ''
            for i in json_loads_file:
                ls.append(i['guid'])
            # 判断单页数量是否到达上限30    
            if len(json_loads_file) == 100:  
                # 超过的话页码加一，再搜索一次
                page += 1
                data['pageNum'] = page
            # 没有超过就结束循环    
            else:                           
                break
        if ls == '':
            return ''
        # 第三次爬取，根据 guid 搜索学员的姓名和邮箱
        lt = []
        for i in ls:
            email_url = 'https://apigateway.gaodun.com/olaf/api/v1/member/guid/'

            emList = requests.get(email_url+i, headers=headers).json()
            name = emList['result']['realName']
            email = emList['result']['email']
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
    
