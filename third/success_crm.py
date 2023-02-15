import requests
import json
import time
import datetime
from .addr import Addr


class WriteCRM():

    def __init__(self, user, pw, n, result, tool, ms, token):
        self.user = user
        self.pw = pw
        self.count = n
        self.tool = tool
        self.ms = ms
        self.x = 2
        self.result = result
        self.info = []
        self.target = 0
        
    def get_info(self, session, url, data):
        headers = {
            'referer': 'https://eds.gaodun.com/',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) \
                Chrome/84.0.4147.105 Safari/537.36'
            }
        data = data
        result = session.get(url, params=data, headers=headers)
        result = str(result.content, encoding='utf-8')
        info = json.loads(result)
        try:
            token = info['data']['session']
        except:
            token = ''
        return info,token

    def post_info(self, session, url, data):
        headers = {
            'referer': 'https://eds.gaodun.com/',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) \
                Chrome/84.0.4147.105 Safari/537.36'
            }
        result = session.post(url, headers=headers, data=data)
        result2 = str(result.content, encoding='utf-8')
        info = json.loads(result2)
        return info
        
    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("开始填写"))
        self.ms.text_print.emit(self.tool.ui.output_3, "共有{}个".format(self.count-1))
        start = time.perf_counter()
        session = requests.Session()
        url = 'https://ssm.gaodun.com/api/member_login/vLogin'
        data = {
            'userName': self.user,
            'password': self.pw,
            'GDSID': 'R0tPWVZISEFvS1ZiWVdGZW1DdmZXRUVWR2NQVnZSa28='
            }
        # 第一次获取登录信息
        info, token = self.get_info(session, url, data)
        for i in range(self.count-1):
            url = 'https://ssm.gaodun.com/api/Authentication/index'
            data = {
                'mobile': str(int(self.result[i][1])),
                'access_token': token,
                'pageSize': '30',
                'p': '1',
                'CourseTypes': '1000053'
                }
            # 第二次爬取，获得member_id
            response, useless = self.get_info(session, url, data)
            # 解析获取member_id，用于后续爬取
            try:
                m_id = response['data']['lists'][0]['member_id']
                # 如果手机查不到，换邮箱再次搜索
            except IndexError:
                data = {
                'email': self.result[i][2],
                'access_token': token,
                'pageSize': '30',
                'p': '1',
                'CourseTypes': '1000053'
                }
                response,useless = self.get_info(session, url, data)
            try:
                m_id = response['data']['lists'][0]['member_id']
            except:
                self.ms.text_print2.emit(self.tool.ui.output_3, "无法获取学员信息")
                continue
            # 第三次爬取，用member_id获取学员student_id，用于下一步
            url = 'https://ssm.gaodun.com/api/Exam/studentExamInfo'
            data = {
                'access_token': token,
                'memberId': m_id,        
                'CourseTypes': '1000053'
                }
            response, useless = self.get_info(session, url, data)
            stu_id = response['data']['ExamInfo']['1000053']['Id']
            # 第四次爬取，提交表单
            url = 'https://ssm.gaodun.com/api/Exam/editExamProject'
            if self.result[i][3] == 'FIA':
                qualification = '1'
            if self.result[i][3] == 'ACCA':
                qualification = '2'
            # 构造数据
            data = {
                'Id': stu_id,
                'account': str(int(self.result[i][4])),
                'password': 'Gaodun123',
                'email': self.result[i][5],
                'email_pass': 'Gaodun123',
                'birthday': str(self.result[i][6]),
                'register_time': str(self.result[i][8]),
                'register_address': Addr.reg_addr.get(self.result[i][7]),
                'register_identity': qualification,
                'approval_date': str(datetime.date.today()),
                'memberId': m_id,
                'projectId': '1000053',
                'authStatus': '2012601',                
                'access_token': token
                }           
            response =  self.post_info(session, url, data)
            if response['info'] == '获取成功':
                self.ms.text_print.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(self.result[i][0],"状态更改成功"))
            else:
                self.ms.text_print2.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(self.result[i][0],"状态更改失败"))
            # 第五次爬取， 填写回访
            url = 'https://ssm.gaodun.com/api/obCallLog/addObCallLogFast'
            data = {
                'memberId': m_id,
                'message': '审核通过，邮件已发送。',
                'type': '2000333',
                'attachment': [],
                'access_token': token
                }
            headers = {
                'referer': 'https://eds.gaodun.com/',
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) \
                Chrome/84.0.4147.105 Safari/537.36'
                }
            message = session.get(url, params=data, headers=headers)
            if message.status_code == 200:
                self.ms.text_print.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(self.result[i][0],"回访填写成功"))
            self.target += 1
            self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
        end = time.perf_counter()
        self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("填写完成"))
        self.ms.text_print.emit(self.tool.ui.output_3, "共用时{:.2f}秒".format(end-start))
        self.ms.text_print.emit(self.tool.ui.output_3, "程序退出")
