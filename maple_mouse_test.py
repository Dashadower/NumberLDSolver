# -*- coding:utf-8 -*-

import win32api, win32con, time

def click(x,y):
    win32api.SetCursorPos((x,y))
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

time.sleep(1)
click(50,180)
time.sleep(2)
click(50,600)
