# --*-- coding : utf-8 --*--
# --*-- author : Euclid Jie --*--

from bs4 import BeautifulSoup
from lxml import etree
from xlutils.copy import copy
import re
import xlrd
import xlwt
import numpy as np

# --*--设置参数 --*--
#html文件路径
html = open("D:\Euclid_Jie\Python\天府金融科技指数\专利\人工智能\已完成\人工智能专利公开年份2020申请年份2017-》【100+53+6+4】.html",encoding="utf-8")
#导出数据路径
path = "D:\Euclid_Jie\Python\天府金融科技指数\专利\人工智能\已完成\用于存储转化后的原始数据的表格.xls"
#sheet名
sheet_name = "测试1"
##表头
value_title = [["专利名","公司名","阶段","申请号", " 申请日", "发明人","公开(公告)号","公开(公告)日"]]

##用于写入数据
def write_excel_xls(path, sheet_name, value):
	index = len(value)  # 获取需要写入数据的行数
	workbook = xlwt.Workbook()  # 新建一个工作簿
	sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
	for i in range(0, index):
		for j in range(0, len(value[i])):
			sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
	workbook.save(path)  # 保存工作簿
	print("xls格式表格写入数据成功！")

##用于增补数据
def write_excel_xls_append(path, value):
	workbook = xlrd.open_workbook(path)  # 打开工作簿
	sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
	worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
	rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
	new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
	new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
	for i in range(0, len(value)):
		new_worksheet.write(rows_old, i, value[i])  # 追加写入数据，注意是从i+rows_old行开始写入
	new_workbook.save(path)  # 保存工作簿

#--*--正式开始工作--*--
write_excel_xls(path, sheet_name, value_title)
soup = BeautifulSoup(html,"html.parser")
invents = soup.find_all('tr');
numbers  = len(invents)
inventdatas = soup.find_all('span',{'class':'val'});
print("共有"+str(numbers)+"条数据待写入")


for i in range(0,numbers):
	inventname = invents[i].td.div.a;
	corname =invents[i].td.div.find('a',target = "_blank");
	stage = invents[i].td.div.div.span
	value = [];
	value.append(inventname.get_text())
	value.append(corname.get_text())
	value.append(stage.get_text())
	value.append([inventdatas[6*i].string])
	value.append([inventdatas[6*i+1].string])
	value.append([inventdatas[6*i+2].string.replace(" ", "")])
	value.append([inventdatas[6*i+3].string])
	value.append([inventdatas[6*i+4].string])
	write_excel_xls_append(path,value)
