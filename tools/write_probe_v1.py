# -*- coding: utf-8 -*-
"""
write_probe_v1.py —— 颜色对话框写入验证

把指定颜色写入 Prism 颜色对话框的 R/G/B 三个输入框，然后读取预览色块的
屏幕像素，自动判定写入是否被对话框接受并刷新。

只修改对话框控件，不保存、不点 OK、不触碰任何文件。
"""

import argparse
import sys
import time

import win32api
import win32con
import win32gui
import win32process


# ============================================================================
# 【使用说明】
#
# 目的
#     Phase 0 已确认 Prism 的颜色对话框是标准 #32770，R/G/B 位于 706/707/708。
#     本脚本验证下一步：用 WM_SETTEXT 写入值之后，对话框会不会刷新预览、
#     Prism 最终读到的是不是新值。
#
# 判定方式
#     不靠肉眼。写入前后各读一次预览色块（控件 2712）中心的屏幕像素，
#     像素颜色变化即说明对话框接受了写入并重绘。
#
# 环境
#     D:\\Conda\\envs\\dock\\python.exe   （已含 pywin32）
#
# 用法
#     python write_probe_v1.py                  写入 255,0,0（纯红）
#     python write_probe_v1.py --rgb 0,128,255
#     python write_probe_v1.py --dry            只读当前状态，不写入
#     python write_probe_v1.py --wait 0.8       延长等待刷新的时间
#
# 前置条件
#     Prism 里已经打开着一个 "Choose Color" 对话框，且它在前台可见。
#
# 安全
#     只对目标对话框的三个输入框发 WM_SETTEXT，并补发 EN_CHANGE 通知。
#     不点 Add、不点 OK、不点 Cancel。测试完你直接点 Cancel 即可丢弃。
# ============================================================================


# ============ 变量区（命令行参数会覆盖这里的默认值） ============
DIALOG_TITLE_KEYWORD = "choose color"   # 小写比较
STD_IDS = {"red": 706, "green": 707, "blue": 708}
PREVIEW_ID = 2712                        # 当前颜色预览（实测得到的 59×58 Static）
EN_CHANGE = 0x0300
DEFAULT_RGB = "255,0,0"
DEFAULT_WAIT = 0.5


# ============ 查找窗口 ============

def get_process_name(pid):
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
            False, pid)
        try:
            return win32process.GetModuleFileNameEx(handle, 0)
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return ""


def find_color_dialog():
    """找 Prism 的 Choose Color 对话框。返回 (hwnd, 进程路径)。"""
    hits = []

    def callback(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            return True
        if DIALOG_TITLE_KEYWORD in title.lower():
            hits.append(hwnd)
        return True

    win32gui.EnumWindows(callback, None)
    if not hits:
        return 0, ""
    hwnd = hits[0]
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return hwnd, get_process_name(pid)


def read_edit_text(hwnd):
    """
    跨进程读控件文本必须用 WM_GETTEXT。GetWindowText 对别的进程只返回
    缓存的标题，对 Edit 控件常常返回空，会造成"写入失败"的误判。
    """
    try:
        length = win32gui.SendMessage(hwnd, 0x000E, 0, 0)   # WM_GETTEXTLENGTH
        if length <= 0:
            return ""
        buf = win32gui.PyMakeBuffer((length + 1) * 2)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buf)
        return buf[:length * 2].tobytes().decode("utf-16-le", errors="ignore")
    except Exception:
        return "?"


def sample_screen(x, y):
    """读屏幕某点的 RGB。返回 (r, g, b) 或 None。"""
    hdc = None
    try:
        hdc = win32gui.GetDC(0)
        value = win32gui.GetPixel(hdc, x, y)
        if value == -1 or value is None:
            return None
        return (value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)
    except Exception:
        return None
    finally:
        if hdc:
            try:
                win32gui.ReleaseDC(0, hdc)
            except Exception:
                pass


def preview_pixel(hdlg):
    """取预览色块中心的屏幕坐标并读像素。"""
    try:
        child = win32gui.GetDlgItem(hdlg, PREVIEW_ID)
    except Exception:
        child = 0
    if not child:
        return None, None
    try:
        left, top, right, bottom = win32gui.GetWindowRect(child)
    except Exception:
        return None, None
    x = (left + right) // 2
    y = (top + bottom) // 2
    return (x, y), sample_screen(x, y)


def is_covered(hdlg):
    """粗查对话框是否在前台，避免读到被遮挡的像素。"""
    try:
        fg = win32gui.GetForegroundWindow()
    except Exception:
        return False
    if fg == hdlg:
        return True
    # 前台可能是它的子控件，向上找
    cur = fg
    for _ in range(4):
        try:
            cur = win32gui.GetParent(cur)
        except Exception:
            break
        if not cur:
            break
        if cur == hdlg:
            return True
    return False


# ============ 主流程 ============

def main():
    parser = argparse.ArgumentParser(description="验证颜色对话框能否被程序写入")
    parser.add_argument("--rgb", default=DEFAULT_RGB,
                        help="目标颜色，格式 R,G,B，默认 255,0,0")
    parser.add_argument("--dry", action="store_true", help="只读，不写入")
    parser.add_argument("--wait", type=float, default=DEFAULT_WAIT,
                        help="写入后等待刷新的秒数")
    parser.add_argument("--raise-window", action="store_true",
                        help="先把对话框带到前台，保证像素判定读到的是它本身")
    args = parser.parse_args()

    try:
        r, g, b = [int(v.strip()) for v in args.rgb.split(",")]
    except Exception:
        print("颜色格式错误，应为 R,G,B（如 255,0,0）")
        return 2
    if not all(0 <= v <= 255 for v in (r, g, b)):
        print("RGB 分量必须在 0-255 之间")
        return 2

    hdlg, proc = find_color_dialog()
    if not hdlg:
        print("未找到标题含 '%s' 的对话框。" % DIALOG_TITLE_KEYWORD)
        print("请确认 Prism 的颜色对话框已打开，且没有被最小化。")
        return 1

    if args.raise_window and not is_covered(hdlg):
        try:
            win32gui.SetForegroundWindow(hdlg)
            time.sleep(0.35)
            print("已尝试把对话框带到前台。")
        except Exception as exc:
            print("置前台失败：%s（可手动点一下对话框再重试）" % exc)

    print("=" * 58)
    print("目标对话框")
    print("  句柄：%08X" % hdlg)
    print("  标题：%s" % win32gui.GetWindowText(hdlg))
    print("  进程：%s" % (proc or "?"))
    print("  前台：%s" % ("是" if is_covered(hdlg) else "否（可能被遮挡，像素判定不可靠）"))
    print("=" * 58)

    edits = {}
    for name, cid in STD_IDS.items():
        try:
            edits[name] = win32gui.GetDlgItem(hdlg, cid)
        except Exception:
            edits[name] = 0
    missing = [n for n, h in edits.items() if not h]
    if missing:
        print("缺少控件：%s" % "、".join(missing))
        return 1

    coord_before, px_before = preview_pixel(hdlg)
    print("")
    print("写入前")
    print("  %s" % "  ".join("%s=%r" % (n, read_edit_text(edits[n]))
                            for n in ("red", "green", "blue")))
    print("  预览框坐标：%s  像素：%s" % (coord_before, px_before))

    if args.dry:
        print("")
        print("--dry 模式，未写入。")
        return 0

    # 写入三个通道
    for name, value in (("red", r), ("green", g), ("blue", b)):
        win32gui.SendMessage(edits[name], win32con.WM_SETTEXT, 0, str(value))
        time.sleep(0.05)

    # 补发 EN_CHANGE，通知父对话框数值已变
    for name, cid in STD_IDS.items():
        wparam = (EN_CHANGE << 16) | cid
        win32gui.PostMessage(hdlg, win32con.WM_COMMAND, wparam, edits[name])
        time.sleep(0.05)

    time.sleep(args.wait)

    coord_after, px_after = preview_pixel(hdlg)
    print("")
    print("写入后")
    print("  %s" % "  ".join("%s=%r" % (n, read_edit_text(edits[n]))
                            for n in ("red", "green", "blue")))
    print("  预览框坐标：%s  像素：%s" % (coord_after, px_after))

    print("")
    print("=" * 58)
    texts = [read_edit_text(edits[n]) for n in ("red", "green", "blue")]
    wrote_ok = texts == [str(r), str(g), str(b)]

    if not wrote_ok:
        print("结论：写入失败。三个输入框的实际文本与目标值不一致。")
        print("      说明这些 Edit 控件不接受 WM_SETTEXT，需改用键盘模拟。")
        return 1

    if px_before is None or px_after is None:
        print("结论：写入成功，但预览像素读取失败，无法自动判定刷新。")
        print("      请肉眼确认预览色块是否变成了目标颜色 %s。" % ((r, g, b),))
        return 0

    changed = px_before != px_after
    if changed:
        # 判断像素是否接近目标色（允许 8 级误差，应对取整与色彩管理）
        diff = max(abs(px_after[0] - r), abs(px_after[1] - g), abs(px_after[2] - b))
        print("结论：写入成功且预览已刷新。")
        print("      预览像素由 %s 变为 %s，与目标色 %s 的最大单通道偏差 %d。"
              % (px_before, px_after, (r, g, b), diff))
        if diff <= 8:
            print("      偏差在容差内，判定对话框已正确接收该颜色。")
        else:
            print("      偏差偏大，可能取到了色块边框或抗锯齿位置，请肉眼复核。")
    else:
        print("结论：值写进去了，但预览像素没有变化（%s）。" % (px_after,))
        print("      说明 WM_SETTEXT 未触发重绘，需要补充刷新手段。")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
