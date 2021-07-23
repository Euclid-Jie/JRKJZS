"""
Created on 2021-07-11
coding: utf-8
vesion: Python3.8.2
@author: 欧玮杰
"""

import pyautogui as pg
#import pyperclip as pp
import time

###网页缩放比例80%###
time.sleep(3)
pg.PAUSE=0.5 

"""
设置参数
"""

##页码输入框坐标
X1 = 600;
X2 = 640;
Y = 360;
# --*-- 自动输入--*--
pages_begin = input("请输入开始页码：")
pages_over = input("请输入结束页码：")

# --*-- 手动输入--*--
#pages_begin = 1
#pages_over = 83

##定义相关函数
def choose_page(pages,X,Y):
    ##滚动左边侧到最下##
    pg.click(950,180)
    pg.scroll(-3000)
    
    ##点击页码输入框
    pg.click(X,Y)
    ##处理页码为数字并输入       
    strs = str(pages)    
    for i in range(0,len(strs)):
        pg.press(strs[i])   
    
    ##点击确定
    pg.click(X+125,Y)    
    time.sleep(4) #留存网页加载时间

def choose_tboby():    
    ##滚动右侧到最下##
    pg.click(1779,310)
    pg.scroll(-3000)
    time.sleep(2) #留存网页加载时间
    
    #搜索body
    pg.click(1026,1003)
    pg.press('enter') #按下enter
    time.sleep(1) #留存网页加载时间
    pg.click(1026,1003)
    pg.press('enter') #按下enter
    time.sleep(5) #留存加载时间
    
    ##滚动右侧到最下##
    pg.click(1779,310)
    pg.scroll(-3000)

    ##复制tbody为HTML
    pg.click(1164,164+176,button='right')
    pg.click(1293,393+176)
    pg.click(1526,555+176)

##切换页面至企查查##
pg.keyDown('alt') #按住不松手
pg.press('tab') #单次点击
pg.keyUp('alt') #松手

choose_page('100')
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
    choose_page(str(i),X1+(X2-X1)*i/(pages_over-pages_begin),Y)
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
    





