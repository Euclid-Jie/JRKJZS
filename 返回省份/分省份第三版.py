"""
Created on 2021-07-13
coding: utf-8
vesion: Python3.8.2
@author: 欧玮杰
"""

import os
import xlrd
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import warnings
warnings.filterwarnings("ignore")

def getLocalFile():
    root=tk.Tk()
    root.withdraw()

    filePath=filedialog.askopenfilename()

    print('文件路径：',filePath)
    return filePath

print("请选择需要提取省份信息的文件")
path1 = getLocalFile()
df=pd.read_excel(path1)
cols = list(df.columns.values)#列标签
"""
输出文件首行
"""
print(df.head(1))
print("请输入所需要处理的数据列数【请输入int数】")
clo = int(input())
colname = cols[clo-1]
df[colname].fillna(value='无数据')


# 导入省份数据文件
#path = os.getcwd(); #返回当前文件夹路径
df1 = pd.read_excel('分省份第三版支持数据.xlsx',engine='openpyxl') 


#处理
print("正在提取省份信息")
df['省份'] = df[colname]
for j in range(1,len(df1['关键词'])):
    df['省份'][df[colname].str.contains(df1['关键词'][j])] = df1['省份'][j]


#保存数据
print("处理完成")
df.to_excel(path1)
