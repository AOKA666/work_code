import selenium
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import datetime
from xlwings import App
import traceback
from threading import Thread, Lock
import threading


class WriteCRM():
    
    def __init__(self, user, pw, app, n, result, tool, ms, teacher):
        self.user = user
        self.pw = pw
        self.app = app
        self.count = n
        self.tool = tool
        self.ms = ms
        self.x = 2
        self.result = result
        self.info = []
        self.lock = Lock()
        self.target = 0
        self.teacher = teacher
        
        
    def main(self, user, pw):
        option = webdriver.ChromeOptions()
        option.add_argument("headless")
        option.add_argument("disable-gpu")
        option.add_argument("--log-level=3")
        try:
            driver = webdriver.Chrome(options=option)
        except:
            self.ms.text_print2.emit(self.tool.ui.output_3, "webdriver版本不一致")
            return ''
        driver.implicitly_wait(5)

        url = 'https://eds.gaodun.com/#/login'
        driver.get(url)
        driver.maximize_window()
        try:
            dirver = self.login(driver, user, pw)
            self.write(driver)
            driver.quit()
        except Exception as e:
            traceback.print_exc()
            self.ms.text_print2.emit(self.tool.ui.output_3, "运行出错")
            driver.quit()


    def login(self, driver, user, pw):
        # 浏览器登录，易出错
        driver.find_element_by_css_selector("input[type='text']").send_keys(self.user)
        driver.find_element_by_css_selector("input[type='password']").send_keys(self.pw)
        time.sleep(1)
        try:
            target = driver.find_element_by_css_selector("button[class='el-button login-btn el-button--primary']")
            driver.execute_script("arguments[0].click();", target)
        except:
            self.ms.text_print2.emit(self.tool.ui.output_3, "账号密码有误，登录失败")
            return ''
        path = "li[role='menuitem']"
        locator = (By.CSS_SELECTOR, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_elements_by_css_selector(path)[5].click()
        self.ms.text_print.emit(self.tool.ui.output_3, "登录成功！")
        self.ms.text_print.emit(self.tool.ui.output_3, "选择认证列表完成")
        # 选择好ACCA
        locator = (By.CSS_SELECTOR, "div[class='el-select']")
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_elements_by_css_selector("div[class='el-select']")[0].click()
        time.sleep(2)
        path = "//div[@x-placement='bottom-start']/div[1]/div[1]/ul/li"
        driver.find_element_by_xpath(path).click()
        # 删除认证老师
        path = "input[title='高顿财经-ACCA认证-{}']".format(self.teacher)
        driver.find_element_by_css_selector(path).click()
        path = "i[class='tree-circle-close el-icon-circle-close']"
        locator = (By.CSS_SELECTOR, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_element_by_css_selector(path).click()
        # 选择 官方账号
        path = "div[class='searchListStyle']"
        locator = (By.CSS_SELECTOR, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_element_by_css_selector(path).click()
        time.sleep(1)
        path = "//span[contains(text(), '官方账号')]/.."
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        target = driver.find_element_by_xpath(path)
        driver.execute_script("arguments[0].click();", target)
        self.ms.text_print.emit(self.tool.ui.output_3, "准备填写...")
        return driver
        
        
    def write(self, driver):
        # 用while不用for，因为这样每次都会对比x和n,lock才有用
        while self.x < self.count+1:
            try:
                self.lock.acquire()
                num = self.result[self.x-2][0]
                name = self.result[self.x-2][1]
                self.x += 1
            finally:
                self.lock.release()
            path = "//div[@class='searchListStyle']/div[1]/input"
            # 输入账号
            driver.find_element_by_xpath(path).send_keys(num)
            driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
            try:
                # 点击这个学员信息进入详情
                target = driver.find_element_by_xpath('//*[@id="pane-0"]/div[1]/div[5]/div[2]/table/tbody/tr/td[12]/div/a')
                driver.execute_script("arguments[0].click();", target)
                time.sleep(1)
            except:
                self.ms.text_print2.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(name,"账号查询不到"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                # excel标记一下错误
                self.info.append([name, num, '查询不到'])
                time.sleep(1)
                path = "//div[@class='searchListStyle']/div[1]/input"
                driver.find_element_by_xpath(path).clear()
                driver.find_element_by_css_selector("div[class='aside-menu']").click()
                time.sleep(1)
                continue
            try: 
                self.try_write(driver)
                self.ms.text_print.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(name,"填写成功"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                self.info.append([name, num, 'OK'])
            except selenium.common.exceptions.ElementClickInterceptedException as e:
                self.ms.text_print2.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(name,"填写出错"))
                self.ms.text_print2.emit(self.tool.ui.output_3, "ElementClickInterceptedException")
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                # excel标记一下错误
                self.info.append([name, num, '填写出错'])
                target = driver.find_elements_by_css_selector("i[class='el-dialog__close el-icon el-icon-close']")[-2]
                driver.execute_script("arguments[0].click();", target)
                time.sleep(0.5)
                target = driver.find_elements_by_css_selector("span[class='el-icon-close']")[1]
                driver.execute_script("arguments[0].click();", target)
                time.sleep(1)
                path = "//div[@class='searchListStyle']/div[1]/input"
                driver.find_element_by_xpath(path).clear()
                driver.find_element_by_css_selector("div[class='aside-menu']").click()
                time.sleep(1)
            except selenium.common.exceptions.NoSuchElementException as e:
                self.ms.text_print2.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(name,"填写出错"))
                self.ms.text_print2.emit(self.tool.ui.output_3, "NoSuchElementException")
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                # excel标记一下错误
                self.info.append([name, num, '填写出错'])
                target = driver.find_elements_by_css_selector("span[class='el-icon-close']")[1]
                driver.execute_script("arguments[0].click();", target)
                time.sleep(1)
                path = "//div[@class='searchListStyle']/div[1]/input"
                driver.find_element_by_xpath(path).clear()
                driver.find_element_by_css_selector("div[class='aside-menu']").click()
                time.sleep(1)
            except selenium.common.exceptions.ElementNotInteractableException as e:
                self.ms.text_print2.emit(self.tool.ui.output_3, "{:<6}:{:<10}".format(name,"填写出错"))
                self.ms.text_print2.emit(self.tool.ui.output_3, "ElementNotInteractableException")
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar_3, self.target)
                # excel标记一下错误
                self.info.append([name, num, '填写出错'])
                target = driver.find_elements_by_css_selector("span[class='el-icon-close']")[1]
                driver.execute_script("arguments[0].click();", target)
                time.sleep(1)
                path = "//div[@class='searchListStyle']/div[1]/input"
                driver.find_element_by_xpath(path).clear()
                driver.find_element_by_css_selector("div[class='aside-menu']").click()
                time.sleep(1)


    def try_write(self, driver):
        # 开始填写crm
        time.sleep(1)
        path = "//div[text()='认证']"
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator)) 
        target = driver.find_element_by_xpath(path)
        driver.execute_script("arguments[0].click();", target)
        time.sleep(1)
        # 更改状态
        driver.find_elements_by_css_selector("div[class='el-input el-input--suffix']")[0].click()
        time.sleep(0.5)
        driver.find_element_by_xpath("//div[@x-placement]/div[1]/div[1]/ul/li[4]").click()
        # 填写日期
        try:
            driver.find_elements_by_css_selector("input[class='el-input__inner']")[-3].send_keys(str(datetime.date.today()))
        except:
            driver.execute_script("window.scrollBy(0, 50);")
            driver.find_elements_by_css_selector("input[class='el-input__inner']")[-3].send_keys(str(datetime.date.today()))
        driver.find_element_by_css_selector("div[class='aside-menu']").click()
        # 保存
        target = driver.find_element_by_xpath('//*[@id="pane-0"]/form/div[4]/button')
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", target)
        time.sleep(1)
      
        # 填写回访
        path = "button[class='el-button el-button--primary el-button--small']"
        target = driver.find_element_by_css_selector(path)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", target)
        # 快速记录  
        path = "//label[@for='sourceType']/following-sibling::div[1]/div[1]/div[1]/input"
        target = driver.find_element_by_xpath(path)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", target)
        # 报备
        path = "//div[@x-placement='bottom-start']/div[1]/div[1]/ul/li[3]"
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        time.sleep(0.5)
        driver.find_element_by_xpath(path).click()
        # 回访    
        driver.find_element_by_css_selector("textarea[placeholder='...请填写详情']").send_keys("审核通过，邮件已发")
        path = '//*[@id="pane-0"]/div[1]/div[2]/div/div[2]/form/div[4]/button[2]'
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_element_by_xpath(path).click()
        # 关闭
        target = driver.find_elements_by_css_selector("span[class='el-icon-close']")[1]
        driver.execute_script("arguments[0].click();", target)
        time.sleep(1)
        path = "//div[@class='searchListStyle']/div[1]/input"
        driver.find_element_by_xpath(path).clear()
        driver.find_element_by_css_selector("div[class='aside-menu']").click()
        time.sleep(1)
        

    def output(self):
        try:
            wb_new = self.app.books.add()
            wb_new.sheets['sheet1'].range('A1').value = '姓名'
            wb_new.sheets['sheet1'].range('B1').value = '账号'
            wb_new.sheets['sheet1'].range('C1').value = '状态'
            n = 2
            for i in self.info:
                if i[2] == '查询不到':
                    wb_new.sheets['sheet1'].range('A%d' % n).value = i[0]
                    wb_new.sheets['sheet1'].range('B%d' % n).value = i[1]
                    wb_new.sheets['sheet1'].range('C%d' % n).value = "查询不到"
                    wb_new.sheets['sheet1'].range('C%d' % n).api.Font.Color = 0x0000ff
                elif i[2] == 'OK':
                    wb_new.sheets['sheet1'].range('A%d' % n).value = i[0]
                    wb_new.sheets['sheet1'].range('B%d' % n).value = i[1]
                    wb_new.sheets['sheet1'].range('C%d' % n).value = "OK"
                    wb_new.sheets['sheet1'].range('C%d' % n).api.Font.Color = 0xff0000
                elif i[2] == '填写出错':
                    wb_new.sheets['sheet1'].range('A%d' % n).value = i[0]
                    wb_new.sheets['sheet1'].range('B%d' % n).value = i[1]
                    wb_new.sheets['sheet1'].range('C%d' % n).value = "填写出错"
                    wb_new.sheets['sheet1'].range('C%d' % n).api.Font.Color = 0x0000ff
                n += 1
            wb_new.save("hidden/填写结果.xls")
            wb_new.close()            
        except Exception as e:
            self.ms.text_print2.emit(self.tool.ui.output_3, "填写结果输出错误")


    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_3, "准备测试中...")
        try:
            option = webdriver.ChromeOptions()
            option.add_argument("headless")
            option.add_argument("disable-gpu")
            option.add_argument("--log-level=3")
            driver = webdriver.Chrome(options=option)
            url = 'https://eds.gaodun.com/#/login'
            driver.get(url)
            driver.quit()
        except:
            self.ms.text_print2.emit(self.tool.ui.output_3, "webdriver无法启动，可能版本不一致...")
            self.ms.text_print.emit(self.tool.ui.output_3, "程序退出")
        else:
            self.ms.text_print.emit(self.tool.ui.output_3, "测试完成，准备填写...")                      
            # 读取完成，开始填写
            self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("开始填写"))
            start = time.perf_counter()
            self.ms.text_print.emit(self.tool.ui.output_3, "共有%s个" % str(self.count-1))            
            try:
                threads = []
                threads_count = 2 if self.count < 10 else 3
                for i in range(threads_count):
                    t = threading.Thread(target=self.main, args=(self.user, self.pw))
                    threads.append(t)
                for t in threads:
                    t.start()
                for t in threads:
                    t.join()
                end = time.perf_counter()
                self.ms.text_print.emit(self.tool.ui.output_3, "{:=^20}".format("填写完成"))
                self.ms.text_print.emit(self.tool.ui.output_3, "共用时{:.2f}秒".format(end-start))
            except Exception as e:
                self.ms.text_print.emit(self.tool.ui.output_3, "程序运行出错，准备退出")
            finally:
                self.output()
                self.ms.text_print.emit(self.tool.ui.output_3, "填写结果输出完毕")
                self.app.quit()
                self.ms.text_print.emit(self.tool.ui.output_3, "程序退出")
