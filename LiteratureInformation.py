from selenium.webdriver.common.by import By
from WebAction import action
import time
import os
import logging

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

download_dir = "/home/lgabyss/project/cnki"  # 设置下载目录
url = 'https://www.cnki.net/?fl=1973326409'  # 目标网页
MAX_NUM = 200 #最大选中数量

# 初始化 WebDriver 及下载路径设置
Controller = action(download_dir)
current_window = []

try:
    # 打开目标网页并调整大小
    Controller.driver.get(url)
    logger.info('页面已打开')
    Controller.driver.set_window_size(1000,1000)
    
    #记录当前窗口
    current_window.append(Controller.driver.current_window_handle)
    
    #------------开始动作-----------------#
    #进入高级检索页面
    Controller.click(By.ID,"highSearch")
    
    #记录高级检索页面
    Controller.switch_windows()
    current_window.append(Controller.driver.current_window_handle)
    
    
    # 等待搜索栏出现并输入关键词
    # Controller.fill_in(By.ID, 'txt_SearchText','NLP') #普通搜索
    Controller.fill_in(By.CSS_SELECTOR, '#gradetxt > dd:nth-child(2) > div:nth-child(2) > input:nth-child(2)','高等教育 + 大学 + 本科生') #高级检索
    logger.info('关键词已输入')
    
    # 点击搜索按钮
    # Controller.search(By.CLASS_NAME, 'search-btn') #普通搜索
    Controller.search(By.CLASS_NAME, 'btn-search') #高级检索
    logger.info('开始搜索')
    
    # 选择主题
    ifOpen = Controller.read_attribute(By.XPATH,"//dl[@groupid='ZYZT|||CYZT']",'class')
    if ifOpen == 'is-up-fold off ':
        Controller.click(By.XPATH,"//dt[@groupid='ZYZT|||CYZT']")
        Controller.waitload(By.CLASS_NAME,"divLoading")
    Controller.click(By.XPATH,"//input[@title='民办高校']")
    Controller.waitload(By.CLASS_NAME,"divLoading")
    # time.sleep(1)
    
    # 选择来源类别
    ifOpen = Controller.read_attribute(By.XPATH,"//dl[@groupid='LYBSM']",'class')
    if ifOpen == 'is-up-fold off ':
        Controller.click(By.XPATH,"//dt[@groupid='LYBSM']")
        Controller.waitload(By.CLASS_NAME,"divLoading")
    Controller.click(By.XPATH,"//input[@title='北大核心']")
    Controller.waitload(By.CLASS_NAME,"divLoading")
    
    # 选择年度
    # ifOpen = Controller.read_attribute(By.XPATH,"//dl[@groupid='YE']",'class')
    # if ifOpen == 'is-up-fold off ':
    #     Controller.click(By.XPATH,"//dt[@groupid='YE']")
    #     time.sleep(1)
    # Controller.click(By.XPATH,"//input[@title='2024年']")
    # time.sleep(1)
    
    # 更换每页显示数量
    Controller.click(By.XPATH,"//*[@id='perPageDiv']/div[@class='sort-default']")
    Controller.click(By.XPATH,"//*[@id='perPageDiv']/ul/li[@data-val='50']")
    Controller.waitload(By.CLASS_NAME,"divLoading")
    per_page = int(Controller.read_value(By.XPATH, "//div[@id='perPageDiv']//div//span"))
    logger.info('更改每页显示数量为%s',per_page)
    
    # 读取总篇数和总页数
    sum = int(Controller.read_value(By.XPATH, "//span[@class='pagerTitleCell']//em").replace(',',''))
    if sum > 50:
        page = int(Controller.read_attribute(By.XPATH, '//span[@class="countPageMark"]','data-pagenum'))
    else:
        page = 1
    logger.info('一共搜索到%s页，需要下载%s篇',page,sum)
    
    #初始化计数并循环
    i = 1           #开始首页
    OriginPage = i #起始页面
    BeginPage = i  #开始页面
    TotalDownload = 0   #此次下载总数量
    
    while i <= page:
        #记录页数
        EndPage = i #结束页面
        
        # 等待新页面加载并点击全选
        Controller.click(By.ID, 'selectCheckAll1')
        logger.info('点击全选')
        # time.sleep(1)
        
        #判断是否超过最大数量
        num = int(Controller.read_value(By.XPATH, '//em[@id="selectCount"]'))
        logger.info('已选中%d篇',num)
        
        if num > MAX_NUM - per_page or i == page:
            # 悬停按钮,出现下拉菜单
            Controller.chains(By.CSS_SELECTOR, '#batchOpsBox > li:nth-child(2) > a:nth-child(1)')
            Controller.chains(By.CSS_SELECTOR, '#batchOpsBox > li:nth-child(2) > ul:nth-child(3) > li:nth-child(1) > a:nth-child(1)')
            
            # 点击“自定义”按钮
            Controller.click(By.CSS_SELECTOR, '#batchOpsBox > li:nth-child(2) > ul:nth-child(3) > li:nth-child(1) > ul:nth-child(3) > li:nth-child(13) > a:nth-child(1)')
            logger.info('跳转至下载界面')

            #切换到新窗口
            Controller.switch_windows(current_window)
            
            #点击全选
            Controller.click(By.XPATH, "//a[@onclick='checkAllFields(true)']")
            
            #取消选择URL
            Controller.click(By.XPATH, "//input[@value='URL-网址']")            
            
            # 点击“下载”按钮
            Controller.click(By.ID, 'litoexcel')
            # 等待下载完成并重命名文件
            
            time.sleep(3)  # 根据网络速度调整等待时间
            downloaded_file = max([os.path.join(download_dir, f) for f in os.listdir(download_dir)], key=os.path.getctime)
            filename = str(BeginPage) + '_' + str(EndPage) + ".xls"
            os.rename(downloaded_file, os.path.join(download_dir, filename))
            logger.info('已下载,文件名为%s',filename)
            
            #更新计数
            BeginPage = i
            EndPage = i
            TotalDownload += num
            
            #关闭窗口并更改活动窗口 
            Controller.close_windows()
            Controller.switch_windows(current_window[0]) 
            
            # 点击“清除”按钮
            Controller.click(By.XPATH, '//a[contains(@href, "javascript:$(this).filenameClear();")]')
            logger.info('清除选中成功')
        
        # 等待新页面加载并点击“下一页”
        if i < page:
            Controller.click(By.ID, 'Page_next_top')
            logger.info('点击下一页')
            Controller.waitload(By.CLASS_NAME,"divLoading")
            # time.sleep(2)
        
        #更新计数
        i += 1
        

finally:
    # 关闭浏览器 
    logger.info('关闭浏览器,本次一共下载%d篇,从%d页到%d页',TotalDownload,OriginPage,EndPage)
    Controller.driver.quit()
