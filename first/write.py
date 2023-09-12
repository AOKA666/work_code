import requests
import json
import time
import datetime


class CRM():

    def __init__(self, user, pw, n, result, ms, tool):
        self.user = user
        self.pw = pw
        self.count = n
        self.tool = tool
        self.ms = ms
        self.result = result
        self.target = 0

        
    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("开始填写"))
        self.ms.text_print.emit(self.tool.ui.output_1, "共有{}个".format(self.count))
        start = time.perf_counter()
        session = requests.Session()
        
        url = 'https://apigateway.gaodun.com/api/v4/vigo/login'
        login_data = {
            "appid": 210666,
            "password": self.pw,
            "user": self.user,
            }              
        # 获取access token
        login = requests.post(url, data=login_data).json()
        access_token = login['accessToken']
        user_id = login['result']['user_id']
        
        # 将 access_token 写入 headers
        headers = {
            "Authentication": "Basic "+access_token,
            "Content-Type": "application/json;charset=UTF-8",
            "Host": "apigateway.gaodun.com",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
            }
        
        # 根据邮箱查询学员信息
        for i in range(self.count):
            url = 'https://apigateway.gaodun.com/olaf/api/v1/auth/list'
            data = {
                "email": self.result[i][1],
                "pageNum":"1",
                "pageSize":10,
                }
            # 第二次爬取，获得 member_id
            details = requests.post(url, data=json.dumps(data), headers=headers).json()
            # 解析获取member_id，用于后续爬取
            try:
                m_id = details['result']['list'][0]['memberId']
                # 如果邮箱查不到，则退出
            except:
                self.ms.text_print2.emit(self.tool.ui.output_1, "无法获取学员信息！姓名：{}".format(self.result[i][0]))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
                continue
                
            # 第三次爬取，用member_id获取学员student_id，用于下一步
            url = 'https://apigateway.gaodun.com/olaf/api/v1/auth/save'
            data = {
                "authStatus": 2012599,
                "authId": 0,
                "id": 264197,
                "projectId": 1000053,
                "memberId": m_id,
                "teacherId": user_id
                }  
            
            response =  requests.post(url, data=json.dumps(data), headers=headers).json()
            if response['message'] == '请求成功':
                self.ms.text_print.emit(self.tool.ui.output_1, "{:<6}:{:<10}".format(self.result[i][0],"状态更改成功"))
            else:
                print(response)
                self.ms.text_print2.emit(self.tool.ui.output_1, "{:<6}:{:<10}".format(self.result[i][0],"状态更改失败"))
            
            # 第四次爬取， 填写回访
            url = 'https://apigateway.gaodun.com/hela/api/v1/manager-task/node-status/recall-other-type'
            data = {
                "memberId": m_id,
                "completeType": 2000333,
                "message": "注册指导邮件已发送",
                "projectId": 1000053,
                }
            response = requests.post(url, data=json.dumps(data), headers=headers).json()
            if response['message'] == "请求成功":
                self.ms.text_print.emit(self.tool.ui.output_1, "{:<6}:{:<10}".format(self.result[i][0],"回访填写成功"))
            self.target += 1
            self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
        end = time.perf_counter()
        self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("填写完成"))
        self.ms.text_print.emit(self.tool.ui.output_1, "共用时{:.2f}秒".format(end-start))
        self.ms.text_print.emit(self.tool.ui.output_1, "程序退出")
