from xlrd import open_workbook
import datetime
import selenium
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import datetime
import time
from xlutils.copy import copy
import xlwt
import traceback


class CRM:
    def __init__(self, user, password, ms, tool):
        self.user = user
        self.password = password
        self.ms = ms
        self.tool = tool
        self.target = 0
        

    # 填写回访的函数
    def crm_write(self, driver):
        wb = open_workbook("代注册.xls", formatting_info=True)
        sheet = wb.sheets()[0]
        num = sheet.nrows
        self.ms.text_print.emit(self.tool.ui.output_1, "共有{}个".format(num))
        if num >= 1:
            self.ms.ui_show.emit(self.tool.ui.progressBar, num)
        wb2 = copy(wb)
        sheet2 = wb2.get_sheet(0)
        for i in range(num):
            name = sheet.cell(i,0).value
            # 每次总点开位于第一个位置的信息操作
            try:
                target = driver.find_element_by_xpath('//*[@id="pane-0"]/div[1]/div[5]/div[2]/table/tbody/tr[1]/td[12]/div/a')
            except:
                self.ms.text_print.emit(self.tool.ui.output_1,"已查询不到新的学员...")
                return ''
            driver.execute_script("arguments[0].click();", target)
            time.sleep(1)
            try:
                self.react_write(driver)
                self.ms.text_print.emit(self.tool.ui.output_1, "{:<6}{:<4}".format(name,"成功"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
                style = self.excel_write(4)
                sheet2.write(i, 10, "CRM:OK", style)
                sheet2.col(10).width = 256*10
                wb2.save("代注册.xls")
            except selenium.common.exceptions.ElementClickInterceptedException as e:
                print(e)
                style = self.excel_write(2)
                sheet2.write(i, 10, "CRM填写出错", style)
                sheet2.col(10).width = 256*10
                wb2.save("代注册.xls")
                self.ms.text_print2.emit(self.tool.ui.output_1, "{:<6}{:<4}".format(name,"失败"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
                self.ms.text_print2.emit(self.tool.ui.output_1, "ElementClickInterceptedException")
                driver.find_elements_by_css_selector("i[class='el-dialog__close el-icon el-icon-close']")[-2].click()
                time.sleep(0.5)
                driver.find_elements_by_css_selector("span[class='el-icon-close']")[1].click()
                time.sleep(0.5)
                driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
                time.sleep(1)
                continue
            except selenium.common.exceptions.NoSuchElementException as e:
                print(e)
                style = self.excel_write(2)
                sheet2.write(i, 10, "CRM填写出错", style)
                sheet2.col(10).width = 256*10
                wb2.save("代注册.xls")
                self.ms.text_print2.emit(self.tool.ui.output_1, "{:<6}{:<4}".format(name,"失败"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
                self.ms.text_print2.emit(self.tool.ui.output_1, "NoSuchElementException")
                driver.find_elements_by_css_selector("span[class='el-icon-close']")[1].click()
                time.sleep(0.5)
                driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
                time.sleep(1)
                continue
            except selenium.common.exceptions.ElementNotInteractableException as e:
                print(e)
                style = self.excel_write(2)
                sheet2.write(i, 10, "CRM填写出错", style)
                sheet2.col(10).width = 256*10
                wb2.save("代注册.xls")
                self.ms.text_print2.emit(self.tool.ui.output_1, "{:<6}{:<4}".format(name,"失败"))
                self.target += 1
                self.ms.ui_show2.emit(self.tool.ui.progressBar, self.target)
                self.ms.text_print2.emit(self.tool.ui.output_1, "ElementNotInteractableException")
                driver.find_elements_by_css_selector("span[class='el-icon-close']")[1].click()
                time.sleep(0.5)
                driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
                time.sleep(1)
                continue
            
          
    def login_select(self, driver):
        # 填写当天的日期
        end = datetime.date.today()
        begin = str(end - datetime.timedelta(days=20))
        end = str(end)
        
        url = 'https://eds.gaodun.com/#/login'
        driver.get(url)
        driver.find_element_by_css_selector("input[type='text']").send_keys(self.user)
        driver.find_element_by_css_selector("input[type='password']").send_keys(self.password)
        try:
            driver.find_element_by_css_selector("button[class='el-button login-btn el-button--primary']").click()
        except:
            self.ms.text_print2.emit(self.tool.ui.output_1, "账号密码有误，登录失败")
            return ''
        self.ms.text_print.emit(self.tool.ui.output_1,"登录成功！")
        locator = (By.CSS_SELECTOR,"li[role='menuitem']")
        WebDriverWait(driver, 5, 0.5).until(EC.presence_of_all_elements_located(locator))
        time.sleep(1)
        driver.find_elements_by_css_selector("li[role='menuitem']")[5].click()
        self.ms.text_print.emit(self.tool.ui.output_1,'选择认证列表...')

        # 选择ACCA
        driver.find_elements_by_css_selector("div[class='el-select']")[0].click()
        path = "//div[@x-placement='bottom-start']/div[1]/div[1]/ul/li"
        time.sleep(2)
        driver.find_element_by_xpath(path).click()
        self.ms.text_print.emit(self.tool.ui.output_1,'项目选择完成')

        # 选择所属阶段
        driver.find_elements_by_css_selector("div[class='el-select']")[1].click()
        path = "//div[@x-placement='bottom-start']/div[1]/div[1]/ul/li"
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        time.sleep(0.5)
        driver.find_element_by_xpath(path).click()

        # 选择阶段状态 易出错！
        driver.find_elements_by_css_selector("div[class='el-select']")[2].click()
        path = "//div[@x-placement='bottom-start']/div[1]/div[1]/ul/li[8]"
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        time.sleep(0.5)
        driver.find_element_by_xpath(path).click()
        self.ms.text_print.emit(self.tool.ui.output_1,'阶段状态选择完成')

        # 选择日期
        driver.find_element_by_css_selector("input[placeholder='开始时间']").send_keys(begin)
        driver.find_element_by_css_selector("input[placeholder='结束时间']").send_keys(end)
        driver.find_element_by_css_selector("div[class='aside-menu']").click()

        # 搜索
        driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
        self.ms.text_print.emit(self.tool.ui.output_1,'准备填写...')
        return driver
    
    
    def react_write(self, driver):
        driver.find_element_by_xpath("//div[text()='认证']").click()
        time.sleep(1)
        # 更改状态
        driver.find_elements_by_css_selector("div[class='el-input el-input--suffix']")[0].click()
        time.sleep(0.5)
        driver.find_element_by_xpath("//div[@x-placement]/div[1]/div[1]/ul/li[2]").click()
        time.sleep(0.5)
        # 保存
        target = driver.find_element_by_xpath('//*[@id="pane-0"]/form/div[4]/button')
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", target)
        # 填写回访
        path = "button[class='el-button el-button--primary el-button--small']"
        target = driver.find_element_by_css_selector(path)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", target)
        # 点击快速记录
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
        driver.find_element_by_css_selector("textarea[placeholder='...请填写详情']").send_keys("代注册邮件已发送")
        path = '//*[@id="pane-0"]/div[1]/div[2]/div/div[2]/form/div[4]/button[2]'
        locator = (By.XPATH, path)
        WebDriverWait(driver, 10, 0.5).until(EC.presence_of_element_located(locator))
        driver.find_element_by_xpath(path).click()
        time.sleep(1)
        # 点击关闭
        path = "span[class='el-icon-close']"
        locator = (By.CSS_SELECTOR, path)
        WebDriverWait(driver, 10, 0.5).until(EC.element_to_be_clickable(locator))
        driver.find_elements_by_css_selector(path)[1].click()
        time.sleep(1)
        
        # 点击搜索
        driver.find_elements_by_css_selector("button[class='el-button el-button--primary']")[0].click()
        time.sleep(1)
    

    # 登录函数
    def do_login(self):
        try:
            option = webdriver.ChromeOptions()
            option.add_argument("headless")
            option.add_argument("disable-gpu")
            option.add_argument("log-level=3")
            driver = webdriver.Chrome(options=option)
        except:
            self.ms.text_print2.emit(self.tool.ui.output_1,"webdriver版本不一致")
            return ''
        else:
            driver.maximize_window()
            driver.implicitly_wait(5)
            return self.login_select(driver)

    
    def excel_write(self,index):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.colour_index = index
        style.font = font
        return style


    def run(self):
        self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("开始填写CRM"))
        start = time.perf_counter()
        driver = self.do_login()
        if driver:
            self.crm_write(driver)
            driver.quit()
            end = time.perf_counter()
            self.ms.text_print.emit(self.tool.ui.output_1, "{:=^20}".format("CRM填写完成"))
            self.ms.text_print.emit(self.tool.ui.output_1, "用时{:.2f}秒".format(end-start))
            self.ms.text_print.emit(self.tool.ui.output_1, "程序退出")
        else:
            self.ms.text_print2.emit(self.tool.ui.output_1,"无法启动程序，可能版本不一致")
            self.ms.text_print2.emit(self.tool.ui.output_1,"填写失败")
            self.ms.text_print.emit(self.tool.ui.output_1,"程序退出")


