#coding:utf-8
# from tkinter import tk

import tkinter as tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
RANGE = range(3, 8)

root = tk.Tk()
# root.deiconify()# 让窗口显示出来


def excel():
    app = 'Excel'#Excel程序
    # xl = win32.gencache.ensuredispath('%s.Application' % app)#静态调度，需要 COM Makepy utility
    xl = win32.Dispatch('%s.Application' % app)#动态调度，打开应用
    ss = xl.Workbooks.Add() #增加一个工作薄，内含多个工作表
    sh = ss.ActiveSheet #取得活动工作表的句柄
    xl.Visible = True#应用在桌面上可见
    sleep(1)#让程序休眠，具体方法是time.sleep(秒数)

    sh.Cells(1, 1).Value = 'Python-to-%s Demo' % app#将Python-to-%s Demo写入到第一个单元格中
    sleep(1)
    for i in RANGE:
        sh.Cells(i, 1).Value = 'Line %d' % i
        sleep(1)
    sh.Cells(i+2, 1).Value = "Th-th-th-that's all folks!"

    warn(app)#警告窗口
    ss.Close(False)#不保存表格内容
    xl.Application.Quit()#退出应用


if __name__ == '__main__':
    root.withdraw()#初始化tk
    excel()