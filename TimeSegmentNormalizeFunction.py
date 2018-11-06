#coding:utf-8
"""
contact:adau22@163.com
all right reserved
timed: 0:24 ; 11.17 2018
"""
""" 
motivation: 学习python 和 Excel 的交联操作，不想在Excel 上输入太多重复的文字，还可以探索其他作用
"""
import tkinter as tk
from time import sleep
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
# RANGE = range(3, 8)

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

    sh.Cells(1, 1).Value = "timeSegment"  #% app 将Python-to-%s Demo写入到第一个单元格中
    sh.Cells(1, 2).Value = "Name"# % app 将Python-to-%s Demo写入到第一个单元格中
    sleep(1)

    i = 2#数据起始行标题栏，属性名称栏之后

    TIME_RANGE = range(8, 22)#时间范围，时
    start_time = 10#起始时间，分
    time_return = 0#截止时间，分
    time_interval = 20#间隔时间
    segment_interval = 5#每一段之间的时间间隔

    for j in TIME_RANGE:
        flag = 1
        k = j
        if j == 8:
            start_minu = start_time
        while flag:
            time_return = start_minu + 20
            if time_return <= 60:#start_minu + 20
                if time_return == 60:
                    # 尾巴等于60
                    time_return = 0
                    k = j + 1#时值 +1
                    sh.Cells(i, 1).Value = '%d:%02d-%d:%02d' % (j, start_minu, k, time_return)
                    start_minu = time_return + segment_interval#重置起始值（分）
                    flag = 0#结束当前时循环标志

                else:#尾巴小于60
                    sh.Cells(i, 1).Value = '%d:%02d-%d:%02d' % (j, start_minu, k, time_return)#str(i)+":%d-" + str(k) + ":%d" % start_minu % start_minu+20#'Line %d' % i
                    start_minu = time_return + segment_interval#重置起始值
                    # 下一次的头部 超过或等于 60
                    if start_minu >= 60:#判断起始值（分）是否超出
                        start_minu -= 60
                        flag = 0
            else:#尾巴大于60
                k = j + 1
                time_return -= 60
                sh.Cells(i, 1).Value = '%d:%02d-%d:%02d' % (j, start_minu, k, time_return) #str(i) + ":%d-" + str(k) + ":%d" % start_minu % start_minu + 20 - 60
                start_minu = time_return + 5
                flag = 0
                """
                冗余判断，用于time_interval 过大
                if start_minu >= 60:#判断起始值（分）是否超出
                    start_minu -= 60
                    flag = 0
                """
            i += 1
            sleep(1)
            if flag == 0:
                break
    sh.Cells(i+2, 1).Value = "ada invention"

    warn(app)#警告窗口
    ss.Close(False)#不保存表格内容 ss.Close([SaveChanges=]False)
    xl.Application.Quit()#退出应用


if __name__ == '__main__':
    root.withdraw()#初始化tk
    excel()