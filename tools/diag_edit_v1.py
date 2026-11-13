# -*- coding: utf-8 -*-
"""
diag_edit_v1.py —— 诊断 RGB 输入框为何拒绝 WM_SETTEXT

输出每个输入框的窗口样式、只读标志、启用状态，以及 WM_SETTEXT 的返回值。
只读诊断，写入的值会立即读回校验。
"""

import sys
import time

import win32api
import win32con
import win32gui


# ============ 变量区 ============
DIALOG_TITLE_KEYWORD = "choose color"
STD_IDS = {"red": 706, "green": 707, "blue": 708}
ES_READONLY = 0x0800
GWL_STYLE = -16
GWL_EXSTYLE = -20
WM_GETTEXT = 0x000D
WM_SETTEXT = 0x000C


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


def style_text(style):
    flags = []
    if style & ES_READONLY:
        flags.append("ES_READONLY(只读)")
    if style & 0x0001:
        flags.append("ES_CENTER")
    if style & 0x0002:
        flags.append("ES_RIGHT")
    if style & 0x0080:
        flags.append("ES_AUTOHSCROLL")
    if style & 0x2000:
        flags.append("ES_NUMBER")
    return " | ".join(flags) if flags else "（无特殊样式）"


def main():
    hdlg = find_dialog()
    if not hdlg:
        print("未找到颜色对话框。")
        return 1

    print("对话框 %08X  %s" % (hdlg, win32gui.GetWindowText(hdlg)))
    print("")
    for name, cid in STD_IDS.items():
        child = win32gui.GetDlgItem(hdlg, cid)
        if not child:
            print("%-6s 控件不存在" % name)
            continue
        style = win32gui.GetWindowLong(child, GWL_STYLE)
        exstyle = win32gui.GetWindowLong(child, GWL_EXSTYLE)
        cls = win32gui.GetClassName(child)
        print("%-6s hwnd=%08X class=%s" % (name, child, cls))
        print("       style=0x%08X  %s" % (style, style_text(style)))
        print("       exstyle=0x%08X" % exstyle)
        print("       enabled=%s visible=%s" % (
            win32gui.IsWindowEnabled(child), win32gui.IsWindowVisible(child)))
        print("       当前文本=%r" % win32gui.GetWindowText(child))

        # 直接 WM_SETTEXT，看返回值（TRUE=1 表示接受）
        ret = win32gui.SendMessage(child, WM_SETTEXT, 0, "123")
        time.sleep(0.1)
        after = win32gui.GetWindowText(child)
        print("       WM_SETTEXT('123') 返回=%r，读回=%r" % (ret, after))

        # 用 WM_GETTEXT 再确认一次（排除 GetWindowText 缓存问题）
        try:
            buf = win32gui.PyMakeBuffer(64)
            n = win32gui.SendMessage(child, WM_GETTEXT, 64, buf)
            print("       WM_GETTEXT 返回长度=%r" % n)
        except Exception as exc:
            print("       WM_GETTEXT 失败：%s" % exc)
        print("")

    print("诊断完毕。若上面返回 0 且读回为空，说明控件被禁用或子类化拦截了写入。")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
