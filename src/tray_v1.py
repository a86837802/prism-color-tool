# -*- coding: utf-8 -*-
"""tray_v1.py —— 系统托盘图标，右键菜单退出"""

import threading

import win32api
import win32con
import win32gui


# ============================================================================
# 【使用说明】
#
# 为什么需要这个：
#     工具是常驻后台的窗口程序，exe 版没有控制台、没有主窗口、面板只在
#     Prism 对话框打开时才出现。没有托盘图标的话，用户找不到任何退出入口，
#     只能去任务管理器结束进程。
#
# 实现方式：
#     不开子进程、不依赖第三方库。在一个后台线程里建一个隐藏窗口，
#     用 Shell_NotifyIcon 注册托盘图标，用 PumpMessages 处理点击事件。
#     右键弹出菜单，选"退出"即回调 on_quit。
#
# 注意：
#     on_quit 在托盘线程里被调用，不能直接操作 tkinter。
#     调用方应当只设置一个标志位，由主线程的消息循环去执行真正的退出。
# ============================================================================


# ============ 变量区 ============
WM_TRAY = win32con.WM_USER + 20
TRAY_ID = 1
MENU_QUIT = 1001
CLASS_NAME = "PrismColorToolTrayWnd"
TOOLTIP = "Prism 取色工具（右键退出）"
# 不加这个标志时 TrackPopupMenu 一律返回 0，选择结果只通过 WM_COMMAND 发出，
# 靠返回值判断"点了哪一项"就会永远失效。
TPM_RETURNCMD = getattr(win32con, "TPM_RETURNCMD", 0x0100)


class TrayIcon:
    def __init__(self, on_quit, tooltip=TOOLTIP):
        self.on_quit = on_quit
        self.tooltip = tooltip
        self.hwnd = 0
        self._thread = None
        self._ready = threading.Event()
        self._quit_fired = False

    def _fire_quit(self):
        """防止返回值与 WM_COMMAND 两条路径重复触发。"""
        if self._quit_fired:
            return
        self._quit_fired = True
        if self.on_quit:
            self.on_quit()

    def start(self, timeout=3.0):
        """启动托盘线程。返回是否在超时前完成注册。"""
        self._thread = threading.Thread(target=self._run, name="tray", daemon=True)
        self._thread.start()
        return self._ready.wait(timeout)

    def _run(self):
        try:
            hinst = win32api.GetModuleHandle(None)
            wc = win32gui.WNDCLASS()
            wc.hInstance = hinst
            wc.lpszClassName = CLASS_NAME
            wc.lpfnWndProc = self._wndproc
            try:
                atom = win32gui.RegisterClass(wc)
            except Exception:
                atom = CLASS_NAME
            self.hwnd = win32gui.CreateWindow(
                atom, "tray", 0, 0, 0, 0, 0, 0, 0, hinst, None)
            icon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
            nid = (self.hwnd, TRAY_ID,
                   win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                   WM_TRAY, icon, self.tooltip)
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
            self._ready.set()
            win32gui.PumpMessages()
        except Exception:
            self._ready.set()

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg == WM_TRAY:
            if lparam in (win32con.WM_RBUTTONUP, win32con.WM_LBUTTONUP):
                self._popup_menu()
        elif msg == win32con.WM_COMMAND:
            # 兜底路径：部分环境下菜单选择走 WM_COMMAND 而不是返回值
            if win32api.LOWORD(wparam) == MENU_QUIT:
                self._fire_quit()
        elif msg == win32con.WM_DESTROY:
            try:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, TRAY_ID))
            except Exception:
                pass
            win32gui.PostQuitMessage(0)
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _popup_menu(self):
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_QUIT, "退出")
        try:
            x, y = win32gui.GetCursorPos()
        except Exception:
            x, y = 100, 100
        win32gui.SetForegroundWindow(self.hwnd)
        command = win32gui.TrackPopupMenu(
            menu,
            win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON | TPM_RETURNCMD,
            x, y, 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
        win32gui.DestroyMenu(menu)
        if command == MENU_QUIT:
            self._fire_quit()

    def stop(self):
        if self.hwnd:
            try:
                win32gui.PostMessage(self.hwnd, win32con.WM_DESTROY, 0, 0)
            except Exception:
                pass
            self.hwnd = 0
