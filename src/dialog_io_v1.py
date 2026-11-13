# -*- coding: utf-8 -*-
"""dialog_io_v1.py —— Prism 颜色对话框的识别与读写"""

import ctypes
import ctypes.wintypes
import time

import win32api
import win32con
import win32gui
import win32process


# ============================================================================
# 【使用说明】
#
# 封装 Prism 颜色对话框的全部 Win32 交互。所有结论均已在本机实测：
#   - 对话框类名 #32770，标题 "Choose Color"，进程 prism.exe
#   - R/G/B 输入框 ID 706/707/708，Add 按钮 712，OK 1，Cancel 2
#   - 透明度输入框 ID 1117（Prism 追加的扩展控件）
#
# 关键坑（勿改）：
#   1. 读别的进程的 Edit 文本必须用 WM_GETTEXT + WM_GETTEXTLENGTH；
#      GetWindowText 对跨进程控件返回空，会造成"写入失败"的误判。
#   2. PyMakeBuffer 单位是字节，WM_GETTEXT 长度参数单位是字符，
#      UTF-16 下缓冲区要乘 2。
#   3. 写完 WM_SETTEXT 后，对话框会自行重算颜色（HSL 三框会同步变化），
#      不需要再发 EN_CHANGE，也不需要模拟键盘。
#
# 单独测试（需先打开 Prism 颜色对话框）：
#   python dialog_io_v1.py
# ============================================================================


# ============ 变量区 ============
DIALOG_CLASS = "#32770"
# 标题关键词只是快速路径。英文版是 "Choose Color"，中文版是「选择颜色」。
# 真正可靠的判定是 _has_rgb_edits：能取到 706/707/708 三个 Edit 控件，
# 就说明这是颜色对话框，与界面语言无关。
TITLE_KEYWORDS = ("choose color", "选择颜色")
PRISM_EXE_KEYWORD = "prism"

ID_RED = 706
ID_GREEN = 707
ID_BLUE = 708
ID_HUE = 703
ID_SAT = 704
ID_LUM = 705
ID_ADD = 712
ID_OK = 1
ID_CANCEL = 2
ID_TRANSPARENCY = 1117
ID_TRANSPARENCY_SLIDER = 1116

TBM_GETPOS = 0x0400
TBM_SETPOS = 0x0405
SB_THUMBPOSITION = 4

WM_GETTEXTLENGTH = 0x000E
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


# ============ 进程 ============

def get_process_name(pid):
    """取进程可执行文件全路径。失败返回空串。"""
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        try:
            return win32process.GetModuleFileNameEx(handle, 0)
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        pass
    try:
        handle = win32api.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = ctypes.wintypes.DWORD(1024)
            if ctypes.windll.kernel32.QueryFullProcessImageNameW(
                    handle, 0, buf, ctypes.byref(size)):
                return buf.value
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        pass
    return ""


# ============ 查找对话框 ============

def has_rgb_edits(hwnd):
    """
    判定一个对话框是不是颜色选择对话框，依据是它内部有没有
    R/G/B 三个 Edit 控件（dlgs.h 的 706/707/708）。
    这条判据不依赖界面语言，中文版、英文版、其他语言版都成立。
    Prism 的 Format Graph 等对话框虽然也是 #32770，但取不到这三个控件。
    """
    for cid in (ID_RED, ID_GREEN, ID_BLUE):
        try:
            child = win32gui.GetDlgItem(hwnd, cid)
        except Exception:
            return False
        if not child:
            return False
        try:
            if win32gui.GetClassName(child) != "Edit":
                return False
        except Exception:
            return False
    return True


def find_color_dialogs():
    """返回所有颜色对话框句柄。先看标题，再按控件结构兜底。"""
    hits = []

    def callback(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) != DIALOG_CLASS:
                return True
            title = win32gui.GetWindowText(hwnd).lower()
        except Exception:
            return True
        if any(keyword in title for keyword in TITLE_KEYWORDS) or has_rgb_edits(hwnd):
            hits.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    return hits


def find_color_dialog():
    hits = find_color_dialogs()
    return hits[0] if hits else 0


def dialog_pid(hdlg):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hdlg)
        return pid
    except Exception:
        return 0


# ============ 文本读写 ============

def read_text(hwnd):
    """跨进程读控件文本。必须用 WM_GETTEXT，见文件头说明。"""
    if not hwnd:
        return ""
    try:
        length = win32gui.SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
        if length <= 0:
            return ""
        buf = win32gui.PyMakeBuffer((length + 1) * 2)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buf)
        return buf[:length * 2].tobytes().decode("utf-16-le", errors="ignore")
    except Exception:
        return ""


def write_text(hwnd, text):
    """写控件文本。返回是否写入成功（以读回校验为准）。"""
    if not hwnd:
        return False
    try:
        win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, str(text))
        time.sleep(0.02)
        return read_text(hwnd) == str(text)
    except Exception:
        return False


# ============ 颜色读写 ============

def get_handles(hdlg):
    """取 R/G/B 三个输入框句柄。缺失的键值为 0。"""
    out = {}
    for name, cid in (("red", ID_RED), ("green", ID_GREEN), ("blue", ID_BLUE)):
        try:
            out[name] = win32gui.GetDlgItem(hdlg, cid)
        except Exception:
            out[name] = 0
    return out


def read_rgb(hdlg):
    """读当前 RGB。任一通道读不到或为空则返回 None。"""
    handles = get_handles(hdlg)
    values = []
    for key in ("red", "green", "blue"):
        text = read_text(handles[key]).strip()
        if not text:
            return None
        try:
            values.append(max(0, min(255, int(text))))
        except ValueError:
            return None
    return tuple(values)


def write_rgb(hdlg, r, g, b):
    """
    写入 RGB。实测：对话框会自行重算颜色，HSL 三框同步更新，
    无需补发 EN_CHANGE，也无需模拟键盘。
    返回 (成功与否, 实际读回的值)。
    """
    handles = get_handles(hdlg)
    for key, value in (("red", r), ("green", g), ("blue", b)):
        if not handles[key]:
            return False, None
        win32gui.SendMessage(handles[key], win32con.WM_SETTEXT, 0, str(int(value)))
        time.sleep(0.02)
    time.sleep(0.06)
    actual = read_rgb(hdlg)
    return (actual == (int(r), int(g), int(b))), actual


def read_transparency(hdlg):
    """
    读透明度百分比（0-100）。
    优先读滑块位置，它才是 Prism 真正使用的值；读不到再退回文本框。
    """
    try:
        slider = win32gui.GetDlgItem(hdlg, ID_TRANSPARENCY_SLIDER)
    except Exception:
        slider = 0
    if slider:
        try:
            position = win32gui.SendMessage(slider, TBM_GETPOS, 0, 0)
            if 0 <= position <= 100:
                return int(position)
        except Exception:
            pass
    try:
        handle = win32gui.GetDlgItem(hdlg, ID_TRANSPARENCY)
    except Exception:
        handle = 0
    if not handle:
        return None
    text = read_text(handle).strip().replace("%", "").strip()
    try:
        return max(0, min(100, int(float(text))))
    except ValueError:
        return None


def write_transparency(hdlg, percent):
    """
    设置透明度。实测：只写 1117 文本框时滑块不动，Prism 读的是滑块位置，
    所以必须 TBM_SETPOS 设位置，再补发 WM_HSCROLL 让对话框响应，
    这样文本框与内部状态才会同步。
    """
    percent = max(0, min(100, int(percent)))
    try:
        slider = win32gui.GetDlgItem(hdlg, ID_TRANSPARENCY_SLIDER)
    except Exception:
        slider = 0
    if not slider:
        return False
    try:
        win32gui.SendMessage(slider, TBM_SETPOS, 1, percent)
        win32gui.SendMessage(hdlg, win32con.WM_HSCROLL,
                             (SB_THUMBPOSITION << 16) | percent, slider)
        time.sleep(0.04)
        return read_transparency(hdlg) == percent
    except Exception:
        return False


def read_hsl(hdlg):
    """读 HSL 三框（Windows 的 winHSL 尺度）。用于交叉验证颜色是否被接受。"""
    values = []
    for cid in (ID_HUE, ID_SAT, ID_LUM):
        try:
            handle = win32gui.GetDlgItem(hdlg, cid)
        except Exception:
            handle = 0
        text = read_text(handle).strip()
        if not text:
            return None
        try:
            values.append(int(text))
        except ValueError:
            return None
    return tuple(values)


def click_add_to_custom(hdlg):
    """点"Add to Custom Colors"。用于把当前颜色存入自定义槽。"""
    try:
        button = win32gui.GetDlgItem(hdlg, ID_ADD)
    except Exception:
        button = 0
    if not button:
        return False
    win32gui.SendMessage(hdlg, win32con.WM_COMMAND, ID_ADD, button)
    return True


def press_ok(hdlg):
    try:
        button = win32gui.GetDlgItem(hdlg, ID_OK)
    except Exception:
        button = 0
    if not button:
        return False
    win32gui.SendMessage(hdlg, win32con.WM_COMMAND, ID_OK, button)
    return True


# ============ 自检 ============

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    dialogs = find_color_dialogs()
    if not dialogs:
        print("没有找到 Choose Color 对话框。请先在 Prism 里打开一个。")
        raise SystemExit(1)

    hdlg = dialogs[0]
    print("对话框 %08X  标题=%s" % (hdlg, win32gui.GetWindowText(hdlg)))
    print("进程 =%s  PID=%d" % (get_process_name(dialog_pid(hdlg)), dialog_pid(hdlg)))
    print("")
    print("当前 RGB      =", read_rgb(hdlg))
    print("当前 HSL      =", read_hsl(hdlg))
    print("当前透明度    =", read_transparency(hdlg))
    print("")

    before = read_rgb(hdlg)
    ok, actual = write_rgb(hdlg, 255, 0, 0)
    print("写入 255,0,0 → 成功=%s  读回=%s" % (ok, actual))
    print("写入后 HSL    =", read_hsl(hdlg))
    print("写入后透明度  =", read_transparency(hdlg))
    if before:
        write_rgb(hdlg, *before)
        print("已还原为写入前的值 %s" % (before,))
