# -*- coding: utf-8 -*-
"""_panel_size_check_v1.py —— 不依赖 Prism 检查面板尺寸与布局"""

import sys
import tkinter as tk

sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parent.parent / "src"))

import win32gui
import prism_color_tool_v1 as t


def main():
    root = tk.Tk()
    root.withdraw()
    tool = t.Tool.__new__(t.Tool)
    tool.root = root
    tool.panel = None
    tool.hdlg = 0

    panel = t.Panel(tool, 0)
    for _ in range(5):
        root.update()

    frame = int(panel.win.wm_frame(), 16)
    rect = win32gui.GetWindowRect(frame)
    print("geometry()   =", panel.win.geometry())
    print("winfo w/h    = %d × %d" % (panel.win.winfo_width(), panel.win.winfo_height()))
    print("实际窗口矩形 =", rect, " 宽=%d 高=%d" % (rect[2] - rect[0], rect[3] - rect[1]))
    print("DPI 缩放     = %.2f" % (panel.win.winfo_fpixels("1i") / 96.0))
    print("")
    print("canvas 高度  =", panel.canvas.winfo_height())
    print("内容总高度   =", panel.inner.winfo_reqheight())
    print("滑块已绘制   =", panel._thumb_box is not None)
    print("色板块数     =", sum(1 for c in panel.inner.winfo_children()
                              if isinstance(c, tk.Frame) and c.winfo_height() > 1))

    # 检查第一套色板那一行是否放得下
    rows = []
    for child in panel.win.winfo_children():
        pass
    print("")
    print("若宽度 < 248，8 个色块会被挤扁，需要调大 PANEL_WIDTH。")

    panel.destroy()
    root.destroy()
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
