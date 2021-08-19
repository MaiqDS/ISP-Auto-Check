# 说明：
# 该程序删去了部分关键信息，不能直接运行。

import os
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import re
import time
import linecache
import win32api
import win32con
from selenium.webdriver.common.by import By

# NO_ALART, LOGINED, INFO, BROWSER_LIST, HELP
NO_ALART = True
LOGINED = False
INFO = {'USER': "", 'PSW': "", 'BROWSER': 0}
BROWSER_LIST = ["Chrome", "Firefox", "Microsoft Edge"]

def Anounce():
    print("*  是否同意以上说明？(Y/N) \n")
    ca = input()
    if ca.upper() == "Y":
        return
    else:
        print("将退出程序！\n")
        os.system("pause")
        sys.exit()

def MainMenu():
    cm = -1
    os.system("cls")
    print("\n\n***** 主菜单 *****\n")
    print("0. 帮助（首次运行必看）\n")
    print("1. 填入打卡信息\n")
    print("*  请选择(0/1)： ")
    while(cm == -1):
        cm = int(input())
        if cm in range(0, 3):
            break
        else:
            print("请输入正确的选项！\n")
            cm = -1
    if cm == 0:
        Help()
    elif cm == 1:
        UserInfo()

def Help():
    print(HELP_SHORT)
    print("按任意键返回上级菜单...")
    input()
    MainMenu()

def UserInfo():
    info_is_right = False

    print("\n\n***** 开始记录打卡信息。 *****")

    while info_is_right == False:
        print("\n*  请输入学号：")
        INFO['USER'] = input()

        print("\n*  请输入密码：")
        INFO['PSW'] = input()

        cu = -1
        while(cu == -1):
            print("\n*  打卡使用的浏览器：")
            print("1. Chrome")
            print("2. Firefox")
            print("3. Microsoft Edge")
            print("4. 其他浏览器")
            print("请选择：")
            cu = int(input())
            if cu in range(1, 4):
                INFO['BROWSER'] = cu
                break
            else:
                if cu == 4:
                    cu = -1
                else:
                    print("请输入正确的选项！\n")
                    cu = -1
        print("\n*  你的信息如下：\n")
        print("学号： "+INFO['USER']+"\n"
              "密码： "+INFO['PSW']+"\n"
              "使用浏览器： "+BROWSER_LIST[INFO['BROWSER']-1]+"\n")
        print("*  核对正确？(Y/N)")
        confirm = input()
        if confirm.upper() == "Y":
            info_is_right = True
            CreateFile()
            print("\n*  设置成功！是否需要预览使用效果？(Y/N)")
            to_preview = input()
            if to_preview.upper() == "Y":
                MainProgram(INFO['USER'], INFO['PSW'], INFO['BROWSER'])
        else:
            print("\n请重新输入：\n")
            continue

def MainProgram(USER, PSW, BROWSER):
    global NO_ALART
    user = INFO['USER']
    psw = INFO['PSW']
    browser = INFO['BROWSER']-1
    try:
        if browser == 0:
            option = webdriver.ChromeOptions()
            option.add_experimental_option(
                "excludeSwitches", ['enable-automation', 'enable-logging'])
            driver = webdriver.Chrome(options=option)
        elif browser == 1:
            option = webdriver.FirefoxOptions()
            option.add_experimental_option(
                "excludeSwitches", ['enable-automation', 'enable-logging'])
            driver = webdriver.Firefox(options=option)
        elif browser == 2:
            option = webdriver.EdgeOptions
            option.add_experimental_option(
                "excludeSwitches", ['enable-automation', 'enable-logging'])
            driver = webdriver.Edge(options=option)
    except Exception as ex:
        print(ex)
        print("\nwebdriver发生错误，请检查帮助文档。\n")

    # 登录账户
    driver.get('somewebsite')
    try:
        for i in user:
            driver.find_element(By.ID, "username").send_keys(i)
        for j in psw:
            driver.find_element(By.ID, "userpwd").send_keys(j)
        verti_code = driver.find_element(
            By.XPATH, "//form[@name='form']/div[3]")
        if verti_code:
            text = verti_code.get_attribute('textContent')
            nums = re.findall(r"\d", text)
            for q in nums:
                driver.find_element(By.ID, "code").send_keys(q)
            driver.find_element(By.ID, "code").send_keys(Keys.ENTER)
    except Exception as ex:
        print(ex)
    # time.sleep(5)
    win = driver.window_handles
    driver.switch_to.window(win[1])
    try:
        alert = driver.switch_to.alert
        print("\n来自网页的警告：")
        print(alert.text)
        NO_ALART = False
    except:
        pass
    if NO_ALART == True:
        try:
            win = driver.window_handles
            driver.switch_to.window(win[1])
            driver.switch_to.frame("leftFrame")
            driver.find_element(By.XPATH, "//font[text()='xxx']").click()
            time.sleep(3)
            driver.switch_to.parent_frame()
            driver.switch_to.frame("main")
            driver.find_element(
                By.XPATH, "//input[contains(@value,'xxx')]").click()
        except Exception as ex:
            print(ex)
    time.sleep(5)
    driver.quit()

def Auto():
    ac = input()
    if ac.upper() == "Y":
        name = 'Auto_Check'
        path = sys.argv[0]
        # 注册表项名
        KeyName = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        # 异常处理
        try:
            key = win32api.RegOpenKey(
                win32con.HKEY_CURRENT_USER,  KeyName, 0,  win32con.KEY_ALL_ACCESS)
            win32api.RegSetValueEx(key, name, 0, win32con.REG_SZ, path)
            win32api.RegCloseKey(key)
        except Exception as ex:
            print(ex)
        print('\n* 添加启动项成功！')
        os.system("pause")
    else:
        print("\n* 已取消设置开机自启！")

def CreateFile():
    list = []
    with open('lsp.txt', 'w', encoding="utf-8") as f:
        f.write(INFO['USER'])
        f.write("\n")
        f.write(INFO['PSW'])
        f.write("\n")
        f.write(str(INFO['BROWSER']))
        f.write("\n")
        f.write(HELP_COMPLETE)

def ReadFile():
    file_path = r'lsp.txt'
    INFO['USER'] = linecache.getline(file_path, 1).strip()
    INFO['PSW'] = linecache.getline(file_path, 2).strip()
    INFO['BROWSER'] = int(linecache.getline(file_path, 3).strip())

def main():
    # ReadFile()
    # print(INFO.values())
    try:
        ReadFile()
        global LOGINED, NO_ALART
        LOGINED = True
        NO_ALART = True
        MainProgram(INFO['USER'], INFO['PSW'], INFO['BROWSER'])
        if NO_ALART == False:
            print("\n*  你的学号或密码设置有误，请重新配置。")
            os.remove('lsp.txt')
    except Exception as ex:
        # print("出现如下异常%s" % ex)
        pass
    if LOGINED == False:
        Anounce()
        MainMenu()
        if NO_ALART == True:
            print("\n*  打卡信息已设置完成。是否需要创建自动打卡任务？(Y/N)")
            to_auto = input()
            if to_auto.upper() == "Y":
                Auto()
        else:
            print("\n*  你的学号或密码设置有误，请重新配置。")
            os.remove('lsp.txt')
    sys.exit()

if __name__ == "__main__":
    main()
