from selenium.webdriver.common.by import By
from WebAction import action
import time
import os
import logging
from FileAction import file

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

download_dir = "/home/lgabyss/project/cnki"  # 设置下载目录
# file_path = "/home/lgabyss/project/cnki/reference.xlsx" #设置活动Excel文件
url = 'https://sci-hub.se/'  # 目标网页 #10.1080/00220671.1956.10882361
# MAX_NUM = 200 #最大选中数量
START_NUM = 25 #开始下载序号（防止中途停止）

# 初始化 WebDriver 及下载路径设置
Controller = action(download_dir)
current_window = []

# 初始化活动Excel
# File = file(file_path)


try:
    # 打开目标网页并调整大小
    Controller.driver.get(url)
    logger.info('页面已打开')
    Controller.driver.set_window_size(1000,1000)
    
    #记录当前窗口
    current_window.append(Controller.driver.current_window_handle)
    
    #------------开始动作-----------------#
    #进入高级检索页面
    # Controller.click(By.ID,"highSearch")
    
    #记录高级检索页面
    # Controller.switch_windows()
    # current_window.append(Controller.driver.current_window_handle)
    
    
    # 等待搜索栏出现并输入关键词
    # Controller.fill_in(By.ID, 'txt_SearchText','NLP') #普通搜索
    Controller.fill_in(By.XPATH, "//textarea[@id='request']",'10.1080/00220671.1956.10882361') #高级检索
    logger.info('关键词已输入')
    
    # 点击搜索按钮
    # Controller.search(By.CLASS_NAME, 'search-btn') #普通搜索
    Controller.search(By.XPATH, '//button[@type="submit"]') #高级检索
    logger.info('开始搜索')
    
finally:
    # 关闭浏览器 
    logger.info('关闭浏览器,本次一共下载%d篇,从%d页到%d页',TotalDownload,OriginPage,EndPage)
    Controller.driver.quit()