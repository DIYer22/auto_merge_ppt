# -*- coding:utf-8 -*-
"""
Created on Thu Jul 28 00:36:24 2016

@author: yl
"""
from __future__ import unicode_literals

import webbrowser
import os,time


SLEEP_TIME = 5  # 操作间隔时间

def p(l):
    for i in l:
        print i

def getFiles():
    def getNum(name,i):
        def changeToNum(char):
            return [1, 2, 3, 4, 5, 6, 7, 8, 9]['一二三四五六七八九'.index(char)]

        if name[i] in '一二三四五六七八九十':
            if name[i:i+3] == '十':
                if name[i+1] in '一二三四五六七八九':
                    return 10 + changeToNum(name[i+1])
                else:
                    return 10
            else:
                if name[i+1] == '十':

                    return changeToNum(name[i:i+3])*10 +\
                    (changeToNum(name[i+2]) if name[i+2] in '一二三四五六七八九' else 0)

            return changeToNum(name[i])
        if name[i] in '0123456789':
            end = i
            while 1:
                if name[end] not in '0123456789' :
                    break
                end += 1
            return int(name[i:end])
        print '合并失败，命名不符合规范！\n请让所有文件夹内的ppt命名规范一致！或者移除命名规范不一致的ppt\n请到 https://github.com/DIYer22/auto_merge_ppt 查看文件说明'
        webbrowser.open('https://github.com/DIYer22/auto_merge_ppt')
        raw_input()
        exit()

    dirr = os.path.abspath('.')

    files = []
    for f in os.listdir(dirr):
        if f[-4:].lower() in ('.ppt','pptx'):
            files += [f]

    def sortNum(files) :
        i = 0
        flag = 0
        while 1:
            for name in files:

                if name[i] != files[0][i]:
                    flag = 1
                    break
            if flag :
                break
            i += 1
        a = files[0]
        return map(lambda x: getNum(x, i), files)

    nums = zip(files,sortNum(files))
    nums.sort(lambda x,y:x[1]-y[1])

    files = [x[0] for x in nums]

    for i in files:
        print i
    return files


import win32com.client, os,sys
import shutil



# 使用时候请命名规范 程序会根据 数字 或 汉字数字自动排序

fileList = getFiles()
files = map(lambda x:os.path.abspath('.') + '\\' + x,
            fileList)
newPPT = os.path.abspath('.')+'\\'+os.path.basename(os.path.abspath('.')).replace(':','').replace('\\','')+'.ppt'
if files[0][-1] == 'x':
    newPPT += 'x'
shutil.copyfile(files[0],newPPT)
files = files[1:]

Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
# Create new presentation
new_ppt = Application.Presentations.Open(newPPT)
time.sleep(SLEEP_TIME)


for f in files:
    # 先打开一遍读页数
    exit_ppt = Application.Presentations.Open(f)
    time.sleep(SLEEP_TIME)
    page_num = exit_ppt.Slides.Count
    exit_ppt.Close()
    num = new_ppt.Slides.InsertFromFile(f,new_ppt.Slides.Count,1,page_num)
    time.sleep(SLEEP_TIME)
new_ppt.Save()
Application.Quit()
