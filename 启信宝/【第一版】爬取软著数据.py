"""
Created on 2021-07-20
coding: utf-8
vesion: Python3.8.2
@author: 欧玮杰
"""
# 导入相关库
import pyautogui as pg
import time

# 设置休眠时间3秒
time.sleep(3)
#设置每次点击时间间隔0.5秒
pg.PAUSE=0.5 

"""
设置参数
"""

##页码输入框坐标
X1 = 842; #在第1页时候的页码X坐标
X2 = 854; #在第10页时候的页码X坐标
X3 = 864; #在第100页时候的页码X坐标
Y = 63; #纵坐标

# --*-- 自动输入--*--
pages_begin = input("请输入开始页码：")
pages_begin = int(pages_begin)
pages_over = input("请输入结束页码：")
pages_over = int(pages_over)

# --*-- 手动输入--*--
#pages_begin = 1
#pages_over = 83

##定义相关函数
def choose_page(pages,X,Y):    
    
    ##点击网址页码输入框
    pg.click(X,Y)

    ##左键点击第两次
    pg.click(X,Y)
    
    if pages<=10:
        pg.press('Backspace')
    elif pages>10 and pages<=100:
        pg.press('Backspace')
        pg.press('Backspace')
    elif pages>100 and pages<1000:
        pg.press('Backspace')
        pg.press('Backspace')
        pg.press('Backspace')   
    elif pages>1000 and pages<10000:
        pg.press('Backspace')
        pg.press('Backspace')
        pg.press('Backspace')       
        pg.press('Backspace')
    
    ##处理页码为数字并输入       
    strs = str(pages)    
    for i in range(0,len(strs)):
        pg.press(strs[i])   
    
    ##点击确定

    pg.press('enter')
    time.sleep(4) #留存网页加载时间

def choose_tboby():    
    ##复制body为HTML
    pg.click(1074,242,button='right')
    pg.click(1202,393)
    pg.click(1423,496)

##切换页面至企查查##
pg.keyDown('alt') #按住不松手
pg.press('tab') #单次点击
pg.keyUp('alt') #松手

choose_tboby()

##切换页面至记事本##
pg.keyDown('alt') #按住不松手
pg.press('tab') #单次点击
pg.press('tab') #单次点击
pg.keyUp('alt') #松手

##复制内容至记事本##
pg.keyDown('ctrl') #按住不松手
pg.press('v') #单次点击
pg.keyUp('ctrl') #松手
pg.press('enter') #单次点击

##循环复制page
for i in range(pages_begin+1,pages_over+1):
    ##切换页面至企查查##
    pg.keyDown('alt') #按住不松手
    pg.press('tab') #单次点击
    pg.keyUp('alt') #松手
    
    ##选择页码
    ##坐标选择
    if i <=10:
        X = X1
    elif i<=100:
        X = X2
    elif i<=1000:
        X = X3
    else:
        X = X3+10
        
    choose_page(i,X,Y)
    ##选择tbody
    choose_tboby()
    
    ##切换页面至记事本##
    pg.keyDown('alt') #按住不松手
    pg.press('tab') #单次点击
    pg.keyUp('alt') #松手
    
    ##复制内容至记事本##
    pg.keyDown('ctrl') #按住不松手
    pg.press('v') #单次点击
    pg.keyUp('ctrl') #松手
    pg.press('enter') #单次点击
    





