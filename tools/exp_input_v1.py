# -*- coding: utf-8 -*-
"""
exp_input_v1.py —— 试验哪种输入方式能真正触发颜色对话框的状态更新

依次尝试三种方法，每试完读一次预览色块的屏幕像素，看哪种能让对话框认账：
  A. WM_SETTEXT 直接设文本
  B. 逐字符 WM_CHAR（让控件自己走输入流程，会自然产生 EN_CHANGE）
  C. 先点一个基本色让对话框进入"已选中"状态，再用 B

注意：C 会真实移动鼠标并点击对话框内的基本色格，请勿在此期间操作鼠标。
"""

import ctypes
import sys
import time

import win32api
import win32con
import win32gui


# ============ 变量区 ============
DIALOG_TITLE_KEYWORD = "choose color"
IDS = {"red": 706, "green": 707, "blue": 708}
BASIC_COLORS_ID = 720        # "Basic colors" 色块矩阵
PREVIEW_ID = 2712
EN_CHANGE = 0x0300
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_CHAR = 0x0102
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_END = 0x23
VK_BACK = 0x08


# ============ 工具 ============

def find_dialog():
    hits = []

    def cb(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) == "#32770":
                if DIALOG_TITLE_KEYWORD in win32gui.GetWindowText(hwnd).lower():
                    hits.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return hits[0] if hits else 0


def read_text(h):
    """
    跨进程读控件文本必须用 WM_GETTEXT（GetWindowText 只返回缓存）。
    先取长度再按长度分配缓冲区，避免截断：PyMakeBuffer 的单位是字节，
    而 WM_GETTEXT 的 wParam 单位是字符，UTF-16 下需乘以 2。
    """
    try:
        length = win32gui.SendMessage(h, 0x000E, 0, 0)   # WM_GETTEXTLENGTH
        if length <= 0:
            return ""
        buf = win32gui.PyMakeBuffer((length + 1) * 2)
        win32gui.SendMessage(h, WM_GETTEXT, length + 1, buf)
        return buf[:length * 2].tobytes().decode("utf-16-le", errors="ignore")
    except Exception:
        return ""


def read_all(hdlg):
    return {name: read_text(win32gui.GetDlgItem(hdlg, cid))
            for name, cid in IDS.items()}


def sample(x, y):
    hdc = None
    try:
        hdc = win32gui.GetDC(0)
        v = win32gui.GetPixel(hdc, x, y)
        if v == -1 or v is None:
            return None
        return (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)
    except Exception:
        return None
    finally:
        if hdc:
            try:
                win32gui.ReleaseDC(0, hdc)
            except Exception:
                pass


def preview_px(hdlg):
    child = win32gui.GetDlgItem(hdlg, PREVIEW_ID)
    if not child:
        return None, None
    left, top, right, bottom = win32gui.GetWindowRect(child)
    x, y = (left + right) // 2, (top + bottom) // 2
    return (x, y), sample(x, y)


def show(tag, hdlg):
    texts = read_all(hdlg)
    coord, px = preview_px(hdlg)
    print("  %-22s 文本=%s  预览像素=%s" %
          (tag, [texts["red"], texts["green"], texts["blue"]], px))
    return px


# ============ 三种方法 ============

def method_settext(hdlg, rgb):
    texts = {}
    for (name, cid), v in zip(IDS.items(), rgb):
        h = win32gui.GetDlgItem(hdlg, cid)
        win32gui.SendMessage(h, WM_SETTEXT, 0, str(v))
        time.sleep(0.05)
    for name, cid in IDS.items():
        h = win32gui.GetDlgItem(hdlg, cid)
        win32gui.PostMessage(hdlg, win32con.WM_COMMAND, (EN_CHANGE << 16) | cid, h)
        time.sleep(0.05)


def method_wmchar(hdlg, rgb):
    """逐字符 WM_CHAR，让 Edit 控件自己走输入路径。"""
    for (name, cid), v in zip(IDS.items(), rgb):
        h = win32gui.GetDlgItem(hdlg, cid)
        # 先全选并删除
        win32gui.SendMessage(h, WM_KEYDOWN, VK_END, 0)
        for _ in range(10):
            win32gui.SendMessage(h, WM_CHAR, VK_BACK, 0)
        for ch in str(v):
            win32gui.SendMessage(h, WM_CHAR, ord(ch), 0)
            time.sleep(0.02)
        time.sleep(0.05)


def click_basic_color(hdlg, index=0):
    """点基本色矩阵里的第 index 个格子。格子按 8 列 6 行排布。"""
    child = win32gui.GetDlgItem(hdlg, BASIC_COLORS_ID)
    if not child:
        print("  找不到基本色矩阵控件")
        return False
    left, top, right, bottom = win32gui.GetWindowRect(child)
    width, height = right - left, bottom - top
    cols, rows = 8, 6
    cw, chh = width / cols, height / rows
    col, row = index % cols, index // cols
    x = int(left + cw * (col + 0.5))
    y = int(top + chh * (row + 0.5))
    print("  点击基本色第 %d 格，屏幕坐标 (%d, %d)" % (index, x, y))
    win32gui.SetForegroundWindow(hdlg)
    time.sleep(0.2)
    win32api.SetCursorPos((x, y))
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.35)
    return True


# ============ 主流程 ============

def main():
    hdlg = find_dialog()
    if not hdlg:
        print("未找到颜色对话框。")
        return 1
    print("对话框 %08X" % hdlg)
    print("")
    print("初始状态：")
    before = show("baseline", hdlg)
    print("")

    # A：WM_SETTEXT
    print("[方法 A] WM_SETTEXT 写入 255,0,0")
    method_settext(hdlg, (255, 0, 0))
    time.sleep(0.4)
    px_a = show("after A", hdlg)
    print("")

    # B：WM_CHAR
    print("[方法 B] 逐字符 WM_CHAR 写入 0,255,0")
    method_wmchar(hdlg, (0, 255, 0))
    time.sleep(0.4)
    px_b = show("after B", hdlg)
    print("")

    # C 会真实移动并点击鼠标，仅在前两种方法都无效时才执行
    ab_worked = ((px_a is not None and px_a != before) or
                 (px_b is not None and px_b != before))
    if not ab_worked:
        print("[方法 C] 先点基本色，再逐字符写入 0,0,255")
        if click_basic_color(hdlg, 0):
            print("  点击后：")
            show("  after click", hdlg)
            method_wmchar(hdlg, (0, 0, 255))
            time.sleep(0.4)
            px_c = show("after C", hdlg)
        else:
            px_c = None
    else:
        print("[方法 C] 跳过（方法 A 或 B 已生效，不需要动鼠标）")
        px_c = None
    print("")

    print("=" * 60)
    print("结果对比：")
    print("  基准像素      %s" % (before,))
    print("  A / WM_SETTEXT %s" % (px_a,))
    print("  B / WM_CHAR    %s" % (px_b,))
    print("  C / 点击后输入 %s" % (px_c,))
    print("")
    if px_b != before and px_b is not None:
        print("→ 方法 B 有效：逐字符输入能触发对话框更新。")
    elif px_a != before and px_a is not None:
        print("→ 方法 A 有效。")
    elif px_c != before and px_c is not None:
        print("→ 必须先用鼠标点一次基本色，之后才接受程序写入。")
    else:
        print("→ 三种方法都未改变预览像素。可能需要对话框在前台并有真实焦点。")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
