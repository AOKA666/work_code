import  requests
from xlrd import open_workbook
import os
from threading import Lock
from concurrent.futures import ThreadPoolExecutor


def send(body):
    max_time = 5    
    def main(max_time):
        global success
        global failure
        num = body["username"]
        url = "https://login.iam.accaglobal.com/accaglobalsso/json/realms/users/users?_action=forgotPassword"
        headers = {
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1",
            "Content-Type": "application/json;charset=UTF-8"
            }                          
        r = requests.post(url,headers=headers,json=body)
        if r.status_code == 200:
            print("{} 发送重置密码链接成功！".format(num))  
            success += 1
        else:
            max_time -= 1
            print("{} 发送失败，还有{}次！".format(num,max_time))
            if max_time:
                main(max_time)
            else:
                print("{} 发送失败".format(num))
                failure += 1
                result.append(num)

    main(max_time)  

if __name__ == '__main__':
    workbook = open_workbook("待审核.xls")
    sheet = workbook.sheets()[0]
    n = sheet.nrows
    print("共有{}个".format(n-1))
    success = 0
    failure = 0
    result = []
    lock = Lock()
    with ThreadPoolExecutor(max_workers=4) as t:
        t.map(send, [{"username": str(int(sheet.row_values(i)[2]))} for i in range(1,n)])
    print("全部发送完成！")
    print("成功数:{},失败数:{},失败账号:{}".format(success,failure,result))
    os.system("pause")
