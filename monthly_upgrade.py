import win32gui
import win32api
import win32con
import win32com.client
import time
import sys
import keyboard

titlename = "PrincessConnectReDive"
hwnd = win32gui.FindWindow(0, titlename)
left, top, right, bottom = win32gui.GetWindowRect(hwnd)
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')
win32gui.BringWindowToTop(hwnd)
win32gui.SetForegroundWindow(hwnd)
# win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
print("这是一个自动拉角色的脚本")
print("请确保以管理员身份运行，\n并且游戏（dmm版）已打开以角色等级排序(降序)的第一个未满级角色")
print("不会自动停下所以全拉完后或者希望停下来时摁住s直到脚本关闭")
print("确保把只升级角色和技能给勾选上否则会装上装备")
L=right-left
W=bottom-top
time.sleep(0.1)
i=0
p1x=left+int(L*(493/1280))
p1y=bottom-int(W*(140/760))
p2x=left+int(L*(790/1280))
p2y=bottom-int(W*(85/760))
p3x=right-int(L*(25/1280))
p3y=bottom-int(W*(360/760))

while 1:
    if keyboard.is_pressed('s'):
        break
    win32api.SetCursorPos((p1x,p1y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.75)
    if keyboard.is_pressed('s'):
        break
    win32api.SetCursorPos((p2x,p2y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(2)
    if keyboard.is_pressed('s'):
        break
    win32api.SetCursorPos((p3x,p3y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(1.5)