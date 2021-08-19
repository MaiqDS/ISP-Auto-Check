# 说明：
# 本源码中包含了上个版本删除的代码。虽然写得乱七八糟，很惭愧，但还是厚着脸皮放了全部。
# 欢迎交流指正。

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
import inspect
import types

# NO_ALART, LOGINED, INFO, BROWSER_LIST, HELP
NO_ALART = True
LOGINED = False
INFO = {'USER': "", 'PSW': "", 'BROWSER': 0}
BROWSER_LIST = ["Chrome", "Firefox", "Microsoft Edge"]
HELP_COMPLETE = '''\n帮助信息：
* 0. 说在最前面：
    本着一劳永逸的心态做出来的东西，多少在配置阶段还是需要一点动手能力（也承认没耐心没能力做成傻瓜式一条龙服务...）。
    配置做起来还是很简单的，只是我写得比较复杂...这个是一开始写的帮助，一不小心写上头了，就写得很长，感觉也很随便，
    所以又搞了个简版的，这个就当彩蛋了吧...\n
* 1. 简介：
    本程序基于python3.8与selenium、win32api库开发，旨在方便使用Windows的成大学子不断打卡、按时返校。没了。\n
* 2. 关于WebDriver配置：
    selenium需要WebDriver以调用Selenium Python bindings API，也就是有了这玩意程序才能跑，
    我在支持的浏览器中选了三种主流的，即Chrome, Firefox和Microsoft Edge，
    没有这三种浏览器的朋友请去下一个。（官方也支持IE和Opera，但懒得测试了）
    简单来说，配置WebDriver一共三步：选择合适的浏览器；到网上下载对应版本的WebDriver；将WebDriver路径写入系统变量PATH。\n
* 2.1 第一步很简单，三个浏览器里选一种\n
* 2.2 第二步需要先知道浏览器版本（火狐不用），这一步请自行百度，
    下面分别贴上三种浏览器WebDriver的下载地址：
    Chrome：https://npm.taobao.org/mirrors/chromedriver/（找到对应版本的文件夹，点进去下载chromedriver_win32.zip）
    Firefox：https://github.com/mozilla/geckodriver/releases/（不用看版本，下win32或win64，一般x64）
    Edge: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/（看了版本再下，一般x64）\n
* 2.3 下载好之后解压，在C盘新建文件夹"webDriver"，把exe文件放进去，右键属性，把他的**位置**复制下来；
    然后右击`此电脑`，点击`属性`，选择右侧`高级系统设置`，在弹出菜单的右下角选择`环境变量`，
    选中下方**系统变量**中的`Path`，点击`编辑`，选择右侧`新建`，将刚刚复制的地址粘贴进去，一路点`确定`退出来就完成了。\n
* 3. 定时打卡功能的说明
    本程序可设置开机启动，但不适用于一天开关2次以上电脑的用户，同样不适用于一直不关机的用户，比如我。
    本程序的启动项名称为'ISP_Auto_Check'，不需要时，可在启动项管理中禁止本程序开机启动。
    因此这里介绍另一种自动打卡的方式，即手动创建Windows计划任务。
    首先还是右击`此电脑`，选择`管理`，点击左侧`任务计划程序`，选择右侧`创建基本任务`，跟着设置向导走就行了。如果还有不会的，百度一下你就知道。
    说两点要注意的，一是本程序机制为只要配置文件存在就不会要求重新设置，所以尽量不要使主程序和txt文件分开，以免造成不必要的麻烦；
    二是尽量把打卡时间设置在没用电脑的时候，因为本程序执行自动操作时将唤起浏览器窗口，模拟键盘鼠标输入，如果执行时同时在操作其他窗口，
    可能引起奇怪的bug。\n
* 4. 碎碎念：
    本来只是自用小程序，突(wei)发(le)奇(rui)想(rui)扩写成这么一个玩意，虽然功能依旧很简单，但打包之后大小达到了惊人的8M，还有众多潜伏bug
    等待修复...自知python基础比较薄弱，代码结构也很臃肿，还停留在面向过程，可能哪天来劲了再大修一遍吧...完整代码放我博客了，还请各路大佬不吝赐教：
    https://harenouta.com/2021/08/13/selenium%e8%87%aa%e5%8a%a8%e6%89%93%e5%8d%a1%e7%a8%8b%e5%ba%8f%e6%ba%90%e7%a0%81%e5%ad%98%e6%a1%a3/
    特别感谢协助我调试的Rarity。
    bug反馈或想和我交流，可发我邮箱maiqds@163.com。感谢使用，祝你生活愉快。\n
    '''

HELP_SHORT = '''\n帮助信息：\n
* 1. 简介：
    本程序基于python3.8与selenium、win32api库开发，旨在方便使用Windows的成大学子不断打卡、按时返校。\n
* 2. 关于WebDriver配置（重要！务必完成以下步骤后再运行主程序）：
    目前支持Chrome, Firefox和Microsoft Edge三种浏览器，
    配置WebDriver共三步：选择合适的浏览器；到网上下载对应版本的WebDriver；将WebDriver路径写入系统变量PATH。\n
* 2.1 略\n
* 2.2 需要知道浏览器版本，不会可百度
    WebDriver下载地址：（选中后点鼠标右键可复制）
    Chrome：https://npm.taobao.org/mirrors/chromedriver/（找对应版本文件夹，点进去下chromedriver_win32.zip）
    Firefox：https://github.com/mozilla/geckodriver/releases/（不用看版本，下win32或win64，一般x64）
    Edge: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/（看了版本再下，一般x64）\n
* 2.3 下载好后解压，在C盘新建文件夹"webDriver"，把exe文件放进去，右键exe属性，把其**位置**复制下来；
    右击`此电脑`，点击`属性`，选择右侧`高级系统设置`，在弹出菜单的右下角选择`环境变量`，
    选中下方**系统变量**中的`Path`，点击`编辑`，选择右侧`新建`，将刚复制的地址粘贴进去，一路点`确定`就好了。\n
* 3. 定时打卡功能的说明
    不推荐频繁开关电脑或不关电脑的用户设置开机启动。本程序的启动项名称为'ISP_Auto_Check'，不需要时，可在启动项管理中禁止本程序开机启动。
    当然，如果觉得太麻烦，放在桌面上，每天手点也是可以的。
    介绍另一种自动打卡方式，即手动创建Windows计划任务。
    首先还是右击`此电脑`，选择`管理`，点击左侧`任务计划程序`，选择右侧`创建基本任务`，跟着设置向导走就行了。
    **注意：**
        尽量不要使主程序和txt文件分开，以免造成不必要的麻烦；
        尽量不要在程序运行时操作其他窗口，否则可能引起奇怪的bug。\n
* 4. bug反馈或想和我交流，可发我邮箱maiqds@163.com。感谢使用，祝你生活愉快。\n
'''


def Anounce():
    print("****************  lsp疫情自动打卡工具 by 某lsp  ****************\n")
    print("*  本工具适用于成都大学ISP系统疫情信息打卡，\n")
    print("*  仅为在打卡期间未离开登记所在地、未出现新冠肺炎症状的同学提供方便，\n")
    print("*  请严格遵守国家防疫规定；若不符合以上条件，请立即停止使用本工具。\n")
    print("*  不符合条件且执意使用本工具的，出现任何问题，由使用者承担所有责任。\n")
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
                    print("\n*  其他浏览器还未支持！请下载三种浏览器中一种使用('・ω・')\n")
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
            print("\n*  设置成功！是否需要预览使用效果？(Y/N)")
            to_preview = input()
            if to_preview.upper() == "Y":
                MainProgram(INFO['USER'], INFO['PSW'], INFO['BROWSER'])
        else:
            print("\n请重新输入打卡信息：\n")
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
            option = webdriver.EdgeOptions()
            option.add_experimental_option(
                "excludeSwitches", ['enable-automation', 'enable-logging'])
            driver = webdriver.Edge(options=option)
    except Exception as ex:
        print(ex)
        print("\nwebdriver发生错误，请检查帮助文档。\n")

    # 登录账户
    try:
        driver.get('https://xsswzx.cdu.edu.cn/ispstu2-1/com_user/weblogin.asp')
        # driver.get('file:./D://!Project//Python//ISP//lsp//管理系统.html')
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
            driver.find_element(By.XPATH, "//font[text()='疫情信息登记']").click()
            time.sleep(3)
            driver.switch_to.parent_frame()
            driver.switch_to.frame("main")
            driver.find_element(
                By.XPATH, "//input[contains(@value,'【一键登记：无变化】')]").click()
        except Exception as ex:
            print(ex)
    time.sleep(5)
    driver.quit()


def Auto():
    # print("建设中...")
    # print("按任意键返回上级菜单...")
    # input()
    # MainMenu()
    print("\n* 该选项将设置程序每次开机时都启动。")
    print("\n* 一天中多次开关电脑、或不关电脑的用户，不推荐使用该选项。")
    print("\n* 是否仍要继续？(Y/N)")
    ac = input()
    if ac.upper() == "Y":
        name = 'ISP_Auto_Check'
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
        print("\n\n* 本程序的启动项名称为'ISP_Auto_Check'，")
        print("\n* 不需要时，可在启动项管理中禁止本程序开机启动。\n")
        os.system("pause")
    else:
        print("\n* 已取消设置开机自启！")


class MultiMethod:
    '''
    Represents a single multimethod.
    '''

    def __init__(self, name):
        self._methods = {}
        self.__name__ = name

    def register(self, meth):
        '''
        Register a new method as a multimethod
        '''
        sig = inspect.signature(meth)

        # Build a type signature from the method's annotations
        types = []
        for name, parm in sig.parameters.items():
            if name == 'self':
                continue
            if parm.annotation is inspect.Parameter.empty:
                raise TypeError(
                    'Argument {} must be annotated with a type'.format(name)
                )
            if not isinstance(parm.annotation, type):
                raise TypeError(
                    'Argument {} annotation must be a type'.format(name)
                )
            if parm.default is not inspect.Parameter.empty:
                self._methods[tuple(types)] = meth
            types.append(parm.annotation)

        self._methods[tuple(types)] = meth

    def __call__(self, *args):
        '''
        Call a method based on type signature of the arguments
        '''
        types = tuple(type(arg) for arg in args[1:])
        meth = self._methods.get(types, None)
        if meth:
            return meth(*args)
        else:
            raise TypeError('No matching method for types {}'.format(types))

    def __get__(self, instance, cls):
        '''
        Descriptor method needed to make calls work in a class
        '''
        if instance is not None:
            return types.MethodType(self, instance)
        else:
            return self


class MultiDict(dict):
    '''
    Special dictionary to build multimethods in a metaclass
    '''

    def __setitem__(self, key, value):
        if key in self:
            # If key already exists, it must be a multimethod or callable
            current_value = self[key]
            if isinstance(current_value, MultiMethod):
                current_value.register(value)
            else:
                mvalue = MultiMethod(key)
                mvalue.register(current_value)
                mvalue.register(value)
                super().__setitem__(key, mvalue)
        else:
            super().__setitem__(key, value)


class MultipleMeta(type):
    '''
    Metaclass that allows multiple dispatch of methods
    '''
    def __new__(cls, clsname, bases, clsdict):
        return type.__new__(cls, clsname, bases, dict(clsdict))

    @classmethod
    def __prepare__(cls, clsname, bases):
        return MultiDict()


class fileOperate(metaclass=MultipleMeta):

    file_path = ''

    def __init__(self):
        self.file_path = r'./lsp.txt'

    def __init__(self, file_argv: str):
        # print(os.path.basename(sys.argv[0]))
        # print(len(os.path.basename(sys.argv[0])))
        self.file_path = sys.argv[0][:len(sys.argv[0])-len(
            os.path.basename(sys.argv[0]))]+file_argv

    def CreateFile(self):
        list = []
        with open('./lsp.txt', 'w', encoding="utf-8") as f:
            f.write(INFO['USER'])
            f.write("\n")
            f.write(INFO['PSW'])
            f.write("\n")
            f.write(str(INFO['BROWSER']))
            f.write("\n")
            f.write(HELP_COMPLETE)

    def ReadFile(self):
        try:
            with open(self.file_path, 'r', encoding="utf-8"):
                print(self.file_path)
                INFO['USER'] = linecache.getline(self.file_path, 1).strip()
                INFO['PSW'] = linecache.getline(self.file_path, 2).strip()
                INFO['BROWSER'] = int(
                    linecache.getline(self.file_path, 3).strip())
        except Exception as ex:
            raise ex


def main():
    try:
        if len(sys.argv) == 2:
            print('手动输入参数。')
            file = fileOperate(sys.argv[1])
            file.ReadFile() == True
        elif len(sys.argv) > 2:
            print('无法识别两个以上的参数。\n')
            sys.exit()
        else:
            print('自动识别路径')
            file = fileOperate()
            file.ReadFile() == True
        global LOGINED, NO_ALART
        LOGINED = True
        NO_ALART = True
        MainProgram(INFO['USER'], INFO['PSW'], INFO['BROWSER'])
        if NO_ALART == False:
            print("\n*  你的学号或密码设置有误，请重新配置。")
            os.remove('./lsp.txt')
    except Exception as ex:
        print("出现如下异常%s" % ex)
        pass
    if LOGINED == False:
        Anounce()
        MainMenu()
        if NO_ALART == True:
            file.CreateFile()
            print("\n*  打卡信息已设置完成。是否需要创建自动打卡任务？(Y/N)")
            to_auto = input()
            if to_auto.upper() == "Y":
                Auto()
            print('''\n\n\n*  你已完成全部设置。\n
*  此后运行此程序将直接打卡，不再进行以上设置。\n
*  若需要重新设置，请删除同目录下的lsp.txt，再打开本程序；\n
*  若只需要查看帮助，txt文件中附有完整帮助文档，请直接打开查看。\n
*  感谢你的使用！o(_ _)o\n
*  希望疫情早日过去。\n\n''')
            os.system("pause")
        else:
            print("\n*  你的学号或密码设置有误，请重新配置。")
    sys.exit()


if __name__ == "__main__":
    main()
