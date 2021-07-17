from selenium import webdriver
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

class stock:
    """
    学习自动化测试
    打开易仓下载相关报表

    """
    def __init__(self):
        """
        构造函数，创建对象的时候会执行
        初始化浏览器
        """
        self.driver = webdriver.Chrome()#打开浏览器
        self.driver.maximize_window()        # 将浏览器最大化显示
        sleep(2)

    def login(self):
        """
        :return:None

        """
        self.driver.get("http://beidi.eccang.com/")  # 访问网页
        # 定位元素
        self.driver.find_element_by_xpath('//*[@id="userName"]').send_keys('2881306367@qq.com')#输入账号
        self.driver.find_element_by_xpath('//*[@id="userPass"]').send_keys('1q2w3e4r5t!@')#输入密码
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="login"]').click()#点击登录
        sleep(1)


    def Input(self):
        """
        跳转到仓配系统
        点击产品销售分析
        下载国内仓库存数据

        :return: None
        """
        self.driver.get('http://beidi.eccang.com/')#跳转到仓配
        Reports  = self.driver.find_element_by_link_text("报表统计")#使用鼠标指向报告中心
        sleep(1)
        ActionChains(self.driver).double_click(Reports).perform()
        sleep(1)
        self. driver.find_element_by_link_text("产品销售分析").click()#点击产品销售分析
        sleep(1)
        self.driver.switch_to.frame("iframe-container-100")#元素定位需要切换frame框架

        Domestic_warehouse = ["SZ1 [深圳仓]","SZ16 [CD中转]","SZ3 [中转UK8]","SZ10 [中转WINIT/4PX/YKD]","SZ9 [中转至FBA]","ZJ3 [宁波中转]"]

        for i in Domestic_warehouse:
            print(i)

            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').click()
            sleep(1)
            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').send_keys(i)
            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').send_keys(Keys.ENTER)#键盘操作 回车

    def Report_Download(self):
        """
        点击下载产品销售分析报表
        点击确认
        :return: None
        """
        self.driver.find_element_by_id('exportOrderSale').click()#生成报告
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="dialog-generate-report"]/p[2]/label/input').click()#选择产品销售分析
        self.driver.find_element_by_xpath('/html/body/div[18]/div[3]/div/button[1]/span').click()#点击确认
        sleep(1)
        self.driver.refresh()
        sleep(2)



    def Input2(self):
        """
        跳转到产品销售分析
        下载海外仓库存数据

        :return: None
        """
        self.driver.switch_to.frame("iframe-container-100")#元素定位需要切换frame框架
        Overseas_warehouse = ["YKD-UK [英国]","YKD-DE [德国]","YKD-CZ [捷克]","YKD-FR [法国]","4PX-UK2 [递四方莱切斯特仓]","US-FBA [美国FBA仓]",
                              "YKD-US-WEST [美西]", "4PX-NEWYORK [4PX美国纽约仓]", "WINIT-DE [万邑通德国仓]" ,"YDY-EAST-NJ6 [易达云美东NJ6]" ,"4PX-UK [递四方英国路腾仓]",
        "DE-FBA [德国FBA仓]","4PX-CZ [4PX捷克仓]","UK-FBA [英国FBA仓]","4PX-BE [4PX比利时列日仓]","YKD-US-EAST [美东]","UK8 [SSC海外仓]","YKD-JP [日本]","CA-FBA [加拿大FBA仓]",
        "4PX-USLA [4PX美国洛杉矶仓]","YDY-CA [易达云加拿大]","YDY-EAST-CA5 [易达云加东5仓]","WYD-EAS2 [无忧达美东新泽西二号仓]","AU1 [WINIT-AU仓]","YDY-WEST-3 [易达云美西3仓]",
        "YDY-CAMIS [易达云加密4]","YDY-EAST-NJ2 [易达云美东NJ2]","YDY-EAST-NJ3 [易达云美东NJ3仓]","YKD-IT [意大利]","YKD-US-USSC [美南]","ES-FBA [西班牙-FBA]","WINIT-UK [WINIT-UK仓]",
        "YDY-UK [易达云英国]","IT-FBA [意大利FBA仓]","XXGJ-WEST [猩猩国际美西仓]","YDY-WEST-1 [易达云美西1仓]","YDY-EAST-001 [易达云美东1仓]","FR-FBA [法国-FBA]","WYD-WEST2 [无忧达美西蒙特克莱尔仓库]",
        "JP-FBA [日本FBA仓]","UK5 [SSC滞销仓]","WYD-EAS [无忧达美东]","4PX-DE [4PX德国法兰克福]","WYD-WEST [无忧达美西洛杉矶]","XXGJ-ATLANTA [猩猩国际美国亚特兰大仓]","4PX-JP [递四方日本大阪仓]","US2 [US-WINIT]",
        "CY-CA-US [成运美西仓]","4PX-CA [4PX加拿大仓]","CY-NJ-US [成运美东仓]","YDY-SOUTH-ATLANTA1 [易达云美南亚特兰大1仓]","4PX-ES [西班牙马德里仓]","YM-EAST [易脉美东中转仓]","YM-WEST [易脉美西中转仓]","LF-DE [乐丰德国中转仓]",
                              "EDIS-DE [橙联德国仓]","WYD-SOUTH [无忧达美南]","MY-ES [MY西班牙仓库]"]

        for i in Overseas_warehouse:
            print(i)

            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').click()
            sleep(1)
            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').send_keys(i)
            self.driver.find_element_by_xpath('//*[@id="searchForm"]/div[9]/div/ul/li/input').send_keys(Keys.ENTER)  # 键盘操作 回车

    def Report_Download2(self):
        """
        选择产品销售分析
        点击确认
        :return:
        """
        self.driver.find_element_by_id('exportOrderSale').click()#生成报告
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="dialog-generate-report"]/p[2]/label/input').click()#选择产品销售分析
        self.driver.find_element_by_xpath('/html/body/div[18]/div[3]/div/button[1]/span').click()#点击确认
        sleep(1)
        self.driver.refresh()
        sleep(2)












    def end(self):
        self.driver.quit()



if __name__ == '__main__':
    stock = stock()
    stock.login()#打开浏览器
    stock.Input()#输入国内仓
    stock.Report_Download()#下载国内仓数据，并刷新页面
    stock.Input2()#输入海外仓信息
    stock.Report_Download2()#下载海外仓信息
    stock.end()#关闭浏览器









#操作元素
