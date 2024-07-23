import pandas as pd
import openpyxl

class file:
    #将目标文件视为活动文件
    def __init__(self,path):
        self.path = path
        try:
            self.wb = openpyxl.load_workbook(self.path)
            self.ws = self.wb.active
        except FileNotFoundError:
            self.wb = openpyxl.Workbook()
            self.ws = self.wb.active
            self.ws.title = 'Sheet1'
            # 如果是新文件，写入第一行标题
            headers = ['文献名称', '参考文献']
            self.ws.append(headers)

    #在文件末行添加数据
    def write_endrow(self,data):
        # 找到最后一行的下一个空行
        next_row = self.ws.max_row + 1

        # 填充数据
        for row in data:
            self.ws.append(row)

        # 保存文件
        self.wb.save(self.path)
        