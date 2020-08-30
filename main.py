#!/usr/bin/env python3
# -*- coding: utf-8 -*
import time
import unittest
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from PIL import Image
from aip import AipOcr
import re
from selenium.webdriver.support.select import Select
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.header import Header
from selenium.webdriver.chrome.options import Options


def get_code(driver):
    png = driver.find_element_by_id('code-box')
    png.screenshot('code.png')
    img = Image.open('code.png')
    img = img.convert('L')  # P模式转换为L模式(灰度模式默认阈值127)
    count = 185  # 设定阈值
    table = []
    for i in range(256):
        if i < count:
            table.append(0)
        else:
            table.append(1)
    img = img.point(table, '1')
    img.save('code_process.png')  # 保存处理后的验证码


def code_rec():
    APP_ID = ''
    API_KEY = ''
    SECRET_KEY = ''
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    def get_file_content(file_path):
        with open(file_path, 'rb') as f:
            return f.read()

    image = get_file_content(file_path='code_process.png')
    options = {'language_type': 'ENG', }  # 识别语言类型，默认为'CHN_ENG'中英文混合

    #  调用通用文字识别
    result = client.basicGeneral(image, options)  # 高精度接口 basicAccurate
    for word in result['words_result']:
        captcha = (word['words'])
        print('验证码识别结果：' + captcha)
        captcha_list = re.findall('[a-zA-Z]', captcha, re.S)[:4]
        captcha_2 = ''.join(captcha_list)
        print('验证码去杂结果：' + captcha_2)
        return captcha_2


def login(driver, captcha, stu_dic):
    try:
        # driver = webdriver.Chrome()
        stu_id = driver.find_element_by_name('txtUid')
        stu_pw = driver.find_element_by_name('txtPwd')
        code = driver.find_element_by_name('code')
        stu_id.send_keys(stu_dic['id'])
        stu_pw.send_keys(stu_dic['pw'])
        code.send_keys(captcha)
        driver.find_element_by_id('Submit').click()
        print("登陆表单提交成功！")
        return True
    except BaseException:
        print("登陆表单提交失败！")
        return False


# 方法必须要继承自unittest.TestCase
class VisitSogouByIE(unittest.TestCase):
    def setUp(self):
        # 启动谷歌浏览器
        self.driver = webdriver.Chrome()

    def isElementPresent(self, by, value, driver):
        try:
            driver.find_element(by=by, value=value)
        except NoSuchElementException:
            print("没有找到指定元素！")
            return False
        else:
            return True


def sub_form(driver):
    try:
        driver.find_element_by_xpath("//div[@id='platfrom2']/a").click()
        select_1 = driver.find_element_by_name('FaProvince').get_attribute('value')
        select1 = Select(driver.find_element_by_name('Province'))
        select1.select_by_value(select_1)
        print("行政区代码" + select_1)
        select_2 = driver.find_element_by_name('FaCity').get_attribute('value')
        select2 = Select(driver.find_element_by_name('City'))
        select2.select_by_value(select_2)
        select_3 = driver.find_element_by_name('FaCounty').get_attribute('value')
        select3 = Select(driver.find_element_by_name('County'))
        select3.select_by_value(select_3)
        ck_cls = driver.find_element_by_xpath('//*[@id="form1"]/div[1]/div[5]/div[3]/div/div/label')
        js = "var q=document.documentElement.scrollTop=10000"
        driver.execute_script(js)
        # ActionChains(driver).send_keys(Keys.END).perform()
        ck_cls.click()
        save_form = driver.find_element_by_class_name('save_form')
        save_form.click()
        print("信息表单提交成功！")
        driver.quit()
        return True
    except BaseException as erro:
        time.sleep(5)
        print(erro)
        already_text = driver.find_element_by_xpath(
            "//*[@class='layui-m-layercont']").text
        if already_text == '当前采集日期已登记！':
            print(already_text)
            driver.quit()
            return already_text
        else:
            print("提交失败！")
            print(already_text)
            return False


def send_email(text, stu_dic):
    try:
        email_from = ''
        email_to = stu_dic['email_to']  # 接收邮箱
        hostname = 'smtp.qq.com'  # QQ邮箱的smtp服务器地址
        login = ''  # 发送邮箱的用户名
        password = ''  # 开启smtp服务得到的授权码。
        subject = '打卡状态'  # 邮件主题

        smtp = SMTP_SSL(hostname)  # SMTP_SSL默认使用465端口
        smtp.login(login, password)

        msg = MIMEText(text, 'plain', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['from'] = email_from
        msg['to'] = email_to

        smtp.sendmail(email_from, email_to, msg.as_string())
        smtp.quit()
        print("发送成功！")
        return True
    except BaseException:
        print("发送失败！")
        return False


if __name__ == "__main__":
    try:
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--headless')
        stu = [
            {'id': 'xxxx', 'pw': 'xxxx', 'email_to': 'xxxx'},
            {'id': 'xxxx', 'pw': 'xxxx', 'email_to': 'xxxx'},
            {'id': 'xxxx', 'pw': 'xxxx', 'email_to': 'xxxx'},
        ]
        for stu_dic in stu:
            test = VisitSogouByIE()  # 实例化对象
            # driver = webdriver.Chrome(options=chrome_options,executable_path='/usr/local/bin/chromedriver')  # 实例化对象
            driver = webdriver.Chrome(options=chrome_options)
            driver.get('http://xsc.sicau.edu.cn/SPCP/Web/')
            driver.implicitly_wait(20)
            get_code(driver)
            captcha = code_rec()
            while True:
                try:
                    if len(captcha) < 4:
                        driver.refresh()
                        time.sleep(2)
                        get_code(driver)
                        captcha = code_rec()
                    else:
                        login(driver, captcha, stu_dic)
                        time.sleep(5)
                        is_exist = test.isElementPresent(
                            driver=driver, by='id', value='platfrom2')
                        if not is_exist:
                            if not test.isElementPresent(driver=driver, by='id', value='layui-m-layer0'):
                                break
                            driver.back()
                            time.sleep(2)
                            driver.refresh()
                            time.sleep(2)
                            get_code(driver)
                            captcha = code_rec()
                            print("登陆失败！")
                        elif is_exist:
                            print("登陆成功！")
                            break
                except Exception as erro:
                    print(erro)
                    captcha = 'xxx'
                    continue
            time.sleep(2)
            is_sub = sub_form(driver)
            if is_sub:
                text = "恭喜您！打卡成功！"
                if is_sub == '当前采集日期已登记！':
                    text = is_sub
                send_email(text, stu_dic)
            elif not is_sub:
                print("打卡失败！稍后重试！")
                text = "打卡失败！稍后重试！"
                send_email(text, stu_dic)
        text = "程序运行结束，所有STU均已操作一遍!"
        send_email(text, stu[0])
    except BaseException as E:
        me = {'id': '201800180', 'pw': '303312', 'email_to': 'zty.zhou@foxmail.com'}
        print(E)
        print("程序意外终止！")
        text = "程序意外终止！"
        send_email(text,me)
