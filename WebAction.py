from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
import time
import logging

class action:
    #初始化
    def __init__(self,download_dir):
        self.options = webdriver.ChromeOptions()
        self.prefs = {"download.default_directory": download_dir}
        self.options.add_experimental_option("prefs", self.prefs)
        self.driver = webdriver.Chrome(options=self.options)
        # 配置日志
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
        
    #搜索动作
    def search(self,type,key):
        search_button = self.driver.find_element(type, key)
        search_button.click()
    
    #点击动作
    def click(self,type,key):
        isSuccess = False
        while isSuccess == False:
            try:  
                button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((type, key))
                )
                button.click()
                isSuccess = True
            except exceptions.StaleElementReferenceException:
                self.logger.info("点击失效,重新尝试")
                time.sleep(0.1)
    
    #填写动作    
    def fill_in(self,type,key,content):
        box = WebDriverWait(self.driver, 100).until(
        EC.presence_of_element_located((type, key))
        )
        box.send_keys(content)
    
    #清空动作    
    def clear(self,type,key):
        box = WebDriverWait(self.driver, 100).until(
        EC.presence_of_element_located((type, key))
        )
        box.clear()
    
    #悬停动作
    def chains(self,type,key):   
        target =  WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((type,key))
        )
        actions = ActionChains(self.driver)
        actions.move_to_element(target).perform()
        
    #隐藏内容显示动作    
    def display(self,type,key,value):
        WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((type,key))
        )
        self.driver.execute_script(value)
    
    #等待动作 
    def waitload(self,type,key):
        target =  WebDriverWait(self.driver, 10).until(
        EC.invisibility_of_element_located((type,key))
        )
    
    #监测元素
    def check(self,type,key):
        try:
            self.driver.find_elements(type,key)
            return True
        except exceptions.NoSuchElementException:
            return False
    
    #读取元素内容及数量
    def get_elements(self,type,key):
        try:
            target = self.driver.find_elements(type,key)
            return target,len(target)
        except exceptions.NoSuchElementException:
            return None,0
    
    #滑动页面至底部
    def scroll_bottom(self):
        temp_height = 0
        isDone = False
        while isDone == False:
            self.driver.execute_script("window.scrollBy(0,600)")
            # sleep一下让滚动条反应一下
            time.sleep(1)
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            check_height = self.driver.execute_script(
                "return document.documentElement.scrollTop || window.pageYOffset || document.body.scrollTop;")
            # 如果两者相等说明到底了
            if check_height == temp_height:
                isDone = True
            temp_height = check_height
    
    #滑动一下页面
    def scroll_once(self):
        self.driver.execute_script("window.scrollBy(0,50)")
    
    #切换窗口动作   
    def switch_windows(self,not_this=None,handle=None):
        if handle == None:
            # 获取所有窗口句柄
            all_windows = self.driver.window_handles

            # 切换到新窗口
            if not_this == None:
                ignore_list = self.driver.current_window_handle
            else:
                ignore_list = not_this
            window_handle = [val for val in all_windows if val not in ignore_list]
            self.driver.switch_to.window(window_handle[0])

        else:
            self.driver.switch_to.window(handle)
    
    #关闭窗口动作        
    def close_windows(self):
        self.driver.close()
    
    #读取属性
    def read_attribute(self,type,key,value):
        # 等待元素可见
        attribute = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((type,key))
        )
        # 获取属性值
        attribute = attribute.get_attribute(value)
        
        return attribute 

    
    #读取值
    def read_value(self,type,key):
        # 等待元素可见
        value = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((type,key))
        )

        # 获取文本内容
        selected_count = value.text
        return selected_count
    
    
    # ---------复合操作--------------
    ## 输出参考文献
    def Get_Reference(self,type,key,name):
        content,num = self.get_elements(type,key + "//ul//li")
        self.logger.info('%s为%d篇',name,num)
        output = []
        if num != 0:
            sum = int(self.read_value(type,key + "//div//b//span"))
            n = 0  
            while n < num:
                # self.logger.info(content[n].text)
                output.append(content[n].text)
                #更新计数
                n += 1
                sum -= 1
                #判断是否读取完
                if n == num and sum != 0:
                    self.click(type,key + "//div//div//a[@class='next']")
                    time.sleep(1)
                    n = 0
                    content,num = self.get_elements(type,key + "//ul//li")         
        # self.logger.info(output)    
        return output