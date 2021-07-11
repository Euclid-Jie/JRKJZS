# -*- coding: utf-8 -*-
"""
Created on Fri Jun 18 14:46:18 2021

@author: 欧玮杰
"""
###设置参数####
pages_begin = 1
pages_over = 83


import pyautogui as pg
#import pyperclip as pp
import time

###网页缩放比例80%###
time.sleep(3)
pg.PAUSE=0.2 

##定义相关函数
def choose_page(pages):
    pg.click(1450,74)
    pg.click(1450,74)
    
    if pages<=10:
        pg.press('back')
    elif pages>10 and pages<=100:
        pg.press('back')
        pg.press('back')
    elif pages>100 and pages<1000:
        pg.press('back')
        pg.press('back')
        pg.press('back')   
    elif pages>1000 and pages<10000:
        pg.press('back')
        pg.press('back')
        pg.press('back')       
        pg.press('back') 
        
    strs = str(pages)    
    for i in range(0,len(strs)):
        pg.press(strs[i])   
    
    pg.press('enter') #按下enter    
    time.sleep(4) #留存网页加载时间

def choose_tboby():    
    ##滚动右侧到最下##
    pg.click(1779,260)
    pg.scroll(-3000)
    time.sleep(2) #留存网页加载时间
    
    #搜索tbody
    pg.click(1026,1003)
    pg.press('enter') #按下enter
    time.sleep(1) #留存网页加载时间
    pg.click(1026,1003)
    pg.press('enter') #按下enter
    time.sleep(5) #留存加载时间

    ##复制tbody为HTML
    pg.click(1164,164,button='right')
    pg.click(1293,393)
    pg.click(1526,555)
    #pg.click(1225,692,button='right')
    #pg.click(1399,471)
    #pg.click(1535,628)

##切换页面至企查查##
pg.keyDown('alt') #按住不松手
pg.press('tab') #单次点击
pg.keyUp('alt') #松手

choose_page(pages_begin)
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
    choose_page(i)
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