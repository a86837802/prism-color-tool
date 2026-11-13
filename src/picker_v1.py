# -*- coding: utf-8 -*-
"""picker_v1.py —— 屏幕吸管（抓屏 + 放大镜 + 邻域中位数采样）"""

import statistics
import tkinter as tk

import win32api
import win32con
import win32gui
import win32ui


# ============================================================================
# 【使用说明】
#
# 提供全屏吸管：点一下屏幕任意位置，取该点颜色，并按 5×5 邻域取中位数，
# 同时计算邻域色差用于提示"是否吸到了边缘/渐变区"。
#
# 为什么不用单点取样：
#     已发表插图的色块带抗锯齿边缘和压缩噪声，单点很容易吸到过渡色，
#     得到一个"看着差不多、实际不对"的值。中位数对这类噪声不敏感。
#
# 为什么不逐点调 GetPixel：
#     那样每移动一像素就要一次系统调用，会明显卡顿。这里一次性抓整屏到
#     内存缓冲区，之后全部在内存里查表，鼠标跟随无延迟。
#
# 单独测试（抓屏 + 读几个像素）：
#   python picker_v1.py
# ============================================================================


# ============ 变量区 ============
MAGNIFY = 10              # 放大镜每个像素放大成多少屏幕像素
GRID = 11                 # 放大镜显示的格子数，必须是奇数
SAMPLE_RADIUS = 0         # 采样半径。0 = 单点；1 = 3×3；2 = 5×5
OFFSET = 18               # 放大镜相对光标的偏移
COLOR_DIFF_WARN = 12      # 邻域色差超过此值提示"可能取到边缘"


class ScreenGrab:
    """一次性把整屏抓到内存，之后按坐标查像素，不再调系统 API。"""

    def __init__(self):
        self.width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        self.height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        hdesktop = win32gui.GetDesktopWindow()
        hwnd_dc = win32gui.GetWindowDC(hdesktop)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfc_dc, self.width, self.height)
        save_dc.SelectObject(bitmap)
        save_dc.BitBlt((0, 0), (self.width, self.height), mfc_dc,
                       (0, 0), win32con.SRCCOPY)
        self.bits = bitmap.GetBitmapBits(True)
        win32gui.DeleteObject(bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hdesktop, hwnd_dc)

    def pixel(self, x, y):
        """返回 (r, g, b)。越界返回黑色。"""
        if not (0 <= x < self.width and 0 <= y < self.height):
            return (0, 0, 0)
        # 实测（2026-11-13，win32ui.GetBitmapBits）：缓冲区是自上而下排列，
        # 不是 Windows DIB 惯例的自下而上。用 height-1-y 会让取色点上下颠倒。
        index = (y * self.width + x) * 4
        return (self.bits[index + 2], self.bits[index + 1], self.bits[index])

    def sample(self, x, y, radius=SAMPLE_RADIUS):
        """
        取 (2r+1)² 邻域。返回 (中位数颜色, 最大单通道色差)。
        中位数比平均值更抗孤立噪点。
        """
        reds, greens, blues = [], [], []
        for dy in range(-radius, radius + 1):
            for dx in range(-radius, radius + 1):
                r, g, b = self.pixel(x + dx, y + dy)
                reds.append(r)
                greens.append(g)
                blues.append(b)
        color = (int(statistics.median(reds)),
                 int(statistics.median(greens)),
                 int(statistics.median(blues)))
        spread = max(max(reds) - min(reds),
                     max(greens) - min(greens),
                     max(blues) - min(blues))
        return color, spread


class Picker:
    """
    全屏吸管。
      on_pick(rgb, spread)  选中时回调
      on_cancel()           按 Esc 或右键取消
    """

    def __init__(self, master, on_pick, on_cancel=None):
        self.on_pick = on_pick
        self.on_cancel = on_cancel
        self.grab = ScreenGrab()
        self._closed = False

        self._build_overlay(master)
        self._build_magnifier(master)
        self._update_magnifier(0, 0)

    # ---------- 构建 ----------

    def _build_overlay(self, master):
        self.overlay = tk.Toplevel(master)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        # alpha 只影响视觉，不影响鼠标事件；0.01 几乎不可见但不会穿透
        self.overlay.attributes("-alpha", 0.01)
        self.overlay.configure(bg="black", cursor="crosshair")
        self.overlay.geometry("%dx%d+0+0" % (self.grab.width, self.grab.height))
        self.overlay.bind("<Motion>", self._on_motion)
        self.overlay.bind("<Button-1>", self._on_click)
        self.overlay.bind("<Escape>", self._on_cancel)
        self.overlay.bind("<Button-3>", self._on_cancel)
        self.overlay.focus_force()

    def _build_magnifier(self, master):
        size = GRID * MAGNIFY
        self.mag = tk.Toplevel(master)
        self.mag.overrideredirect(True)
        self.mag.attributes("-topmost", True)
        self.mag.configure(bg="#000000")

        self.canvas = tk.Canvas(self.mag, width=size, height=size,
                                bd=0, highlightthickness=0, bg="#000000")
        self.canvas.pack()

        self.cells = []
        for gy in range(GRID):
            for gx in range(GRID):
                rect = self.canvas.create_rectangle(
                    gx * MAGNIFY, gy * MAGNIFY,
                    (gx + 1) * MAGNIFY, (gy + 1) * MAGNIFY,
                    outline="", fill="#000000")
                self.cells.append((rect, gx, gy))

        # 红框标出真正会被采样的区域。取的是 5×5 中位数，
        # 所以框要覆盖这 25 格，而不是中心那一格——否则用户以为
        # 取的是单点，实际结果却和框里那个像素不一致。
        center = GRID // 2
        radius = SAMPLE_RADIUS
        self.canvas.create_rectangle(
            (center - radius) * MAGNIFY, (center - radius) * MAGNIFY,
            (center + radius + 1) * MAGNIFY, (center + radius + 1) * MAGNIFY,
            outline="#FF3B30", width=2)
        # 中心十字，标出光标所在的那一个像素
        cx = center * MAGNIFY + MAGNIFY // 2
        cy = center * MAGNIFY + MAGNIFY // 2
        self.canvas.create_line(cx - 5, cy, cx + 5, cy, fill="#FF3B30")
        self.canvas.create_line(cx, cy - 5, cx, cy + 5, fill="#FF3B30")

        self.label = tk.Label(self.mag, text="", bg="#1E1E1E", fg="#E6E6E6",
                              font=("Consolas", 9), anchor="w", padx=4, pady=2)
        self.label.pack(fill="x")

    # ---------- 事件 ----------

    def _on_motion(self, event):
        # 不用 event.x_root / y_root：Tk 8.6 在 DPI 感知进程里可能给出
        # 逻辑坐标（例如 1536 体系），而抓屏缓冲区是物理像素（1920 体系），
        # 两者混用会让取色点系统性偏移。统一改用 Win32 的物理坐标。
        x, y = win32api.GetCursorPos()
        self._update_magnifier(x, y)

    def _on_click(self, event):
        x, y = win32api.GetCursorPos()
        color, spread = self.grab.sample(x, y)
        self._close()
        if self.on_pick:
            self.on_pick(color, spread)

    def _on_cancel(self, _event=None):
        self._close()
        if self.on_cancel:
            self.on_cancel()

    # ---------- 渲染 ----------

    def _update_magnifier(self, x, y):
        half = GRID // 2
        for rect, gx, gy in self.cells:
            r, g, b = self.grab.pixel(x - half + gx, y - half + gy)
            self.canvas.itemconfig(rect, fill="#%02X%02X%02X" % (r, g, b))
        color, spread = self.grab.sample(x, y)
        if SAMPLE_RADIUS == 0:
            mode = "单点"
        else:
            side = SAMPLE_RADIUS * 2 + 1
            mode = "%d×%d 中位数" % (side, side)
        hint = " 边缘?" if spread > COLOR_DIFF_WARN else ""
        self.label.config(
            text="#%02X%02X%02X  %d,%d,%d\n%s%s"
                 % (color + color + (mode, hint)))

        size = GRID * MAGNIFY
        screen_w = self.grab.width
        screen_h = self.grab.height
        px = x + OFFSET
        py = y + OFFSET
        if px + size > screen_w:
            px = x - OFFSET - size
        if py + size + 30 > screen_h:
            py = y - OFFSET - size - 30
        self.mag.geometry("+%d+%d" % (max(0, px), max(0, py)))

    def _close(self):
        if self._closed:
            return
        self._closed = True
        for window in (self.overlay, self.mag):
            try:
                window.destroy()
            except Exception:
                pass


def pick_color(master, on_pick, on_cancel=None):
    """便捷入口。"""
    return Picker(master, on_pick, on_cancel)


if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    grab = ScreenGrab()
    print("屏幕尺寸：%d × %d" % (grab.width, grab.height))
    for point in [(0, 0), (100, 100), (grab.width // 2, grab.height // 2)]:
        color, spread = grab.sample(*point)
        print("  %-12s → RGB%s  色差=%d" % (str(point), color, spread))
