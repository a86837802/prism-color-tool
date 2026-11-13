# -*- coding: utf-8 -*-
"""_coord_check_v1.py —— 校准 ScreenGrab 的坐标映射

用 GetPixel 逐点读取屏幕真实颜色，与内存缓冲区的读取结果对照，
判断缓冲区是自上而下还是自下而上排列。
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import win32api
import win32gui
from picker_v1 import ScreenGrab


def split(value):
    return (value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)


def main():
    grab = ScreenGrab()
    print("缓冲尺寸      : %d × %d" % (grab.width, grab.height))
    print("GetSystemMetrics: %d × %d" % (
        win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)))
    print("")

    hdc = win32gui.GetDC(0)
    points = [
        (grab.width // 2, 20),
        (grab.width // 2, grab.height - 40),
        (80, 40),
        (80, grab.height - 60),
        (grab.width - 120, 30),
        (grab.width - 120, grab.height - 80),
    ]
    ok = flip = bad = 0
    for (x, y) in points:
        mine = grab.pixel(x, y)
        real = split(win32gui.GetPixel(hdc, x, y))
        flipped = grab.pixel(x, grab.height - 1 - y)
        if mine == real:
            verdict = "正常"
            ok += 1
        elif flipped == real:
            verdict = "上下翻转"
            flip += 1
        else:
            verdict = "都不一致"
            bad += 1
        print("(%4d,%4d) 缓冲=%-16s 真实=%-16s 翻转读取=%-16s → %s"
              % (x, y, mine, real, flipped, verdict))
    win32gui.ReleaseDC(0, hdc)

    print("")
    if flip > ok and flip > bad:
        print("结论：缓冲区是自上而下排列，pixel() 里的 height-1-y 需要去掉。")
    elif ok >= flip and ok >= bad:
        print("结论：坐标映射正常，偏移另有原因。")
    else:
        print("结论：两个方向都对不上，可能是抓屏尺寸与屏幕不一致（DPI 缩放）。")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
