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
file_path = "/home/lgabyss/project/cnki/reference.xlsx" #设置活动Excel文件
url = 'https://www.cnki.net/?fl=1973326409'  # 目标网页
# MAX_NUM = 200 #最大选中数量
START_NUM = 25 #开始下载序号（防止中途停止）

# 初始化 WebDriver 及下载路径设置
Controller = action(download_dir)
current_window = []

# 初始化活动Excel
File = file(file_path)


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
    
    #定位到上次目标开始文献
    
    
    #初始化计数并循环
    i = 1           #开始首页
    OriginPage = i #起始页面
    BeginPage = i  #开始页面
    TotalDownload = 0   #此次下载总数量
    
    while i <= page:
        #记录页数
        EndPage = i #结束页面
        
        #初始化每页数量
        j = 1
        
        while j <= per_page:
            if START_NUM == 1:
                #生成目标CSS选择器
                # CSS_Selector = ".result-table-list > tbody:nth-child(2) > tr:nth-child(" + str(j) + ") > td:nth-child(2) > a:nth-child(1)"
                CSS_Selector = "#gridTable > div > div > table > tbody > tr:nth-child(" + str(j) + ") > td.name > a"
                #gridTable > div > div > table > tbody > tr:nth-child(19) > td.name > a
                #gridTable > div > div > table > tbody > tr:nth-child(17) > td.name > a
                #.result-table-list > tbody:nth-child(2) > tr:nth-child(19) > td:nth-child(2) > a:nth-child(1)
                #.result-table-list > tbody:nth-child(2) > tr:nth-child(47) > td:nth-child(2) > a:nth-child(1)
                #点击文献链接
                Controller.click(By.CSS_SELECTOR,CSS_Selector)
                logger.info('跳转至下载界面')

                #切换到新窗口
                Controller.switch_windows(current_window)
                
                #滑动到底端，确保参考文献被加载
                Controller.scroll_bottom()
                
                #等待加载完毕
                isDone = False
                while isDone == False:
                    isDone = Controller.check(By.XPATH,"//div[@id='nxgp-kcms-data-ref-references-crldeng']")
                    time.sleep(1)
                
                # -------------开始下载数据------------
                ## 读取文献名称
                name = Controller.read_value(By.XPATH, "//div[@class='wx-tit']//h1")
                logger.info('文献名称:%s',name)
                ## 读取目标内容
                content = []
                content = content + Controller.Get_Reference(By.XPATH,"//div[@id='nxgp-kcms-data-ref-references-journal']",'参考国内文献')
                content = content + Controller.Get_Reference(By.XPATH,"//div[@id='nxgp-kcms-data-ref-references-journal-w']",'参考国际文献')
                content = content + Controller.Get_Reference(By.XPATH,"//div[@id='nxgp-kcms-data-ref-references-book']",'参考图书')
                
                ## 转化为单个字符串
                reference = ''
                for row in content:
                    reference = reference + row + '\n'
                logger.info('参考文献为:%s',reference.rstrip('\n'))
                ## 输出至excel
                File.write_endrow([[name , reference]])
                # -------------下载数据 End------------
                
                #更新本次下载总量
                TotalDownload += 1
                logger.info('已下载第%d/%d篇，本次共下载%d篇',(i-1) * per_page + j, sum ,TotalDownload)
                
                #关闭窗口并更改活动窗口 
                Controller.close_windows()
                Controller.switch_windows(current_window[0])
            
            else:
                START_NUM -= 1
            
            #更新计数
            j += 1

            # 往下滑动窗口，防止点击不到目标文献
            Controller.scroll_once()
    
        
        
        # 等待新页面加载并点击“下一页”
        if i < page:
            Controller.click(By.ID, 'PageNext')
            logger.info('点击下一页')
            Controller.waitload(By.CLASS_NAME,"divLoading")
            time.sleep(2)
        
        #更新计数
        i += 1      
finally:
    # 关闭浏览器 
    logger.info('关闭浏览器,本次一共下载%d篇,从%d页到%d页',TotalDownload,OriginPage,EndPage)
    Controller.driver.quit()
