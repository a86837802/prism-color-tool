# -*- coding: utf-8 -*-
"""prism_color_tool_v1.py —— Prism 颜色对话框接管工具"""

import ctypes
import sys
import time
import tkinter as tk

import win32api
import win32event
import winerror

import dialog_io_v1 as dio
import color_convert_v1 as cc
import palette_data_v1 as pd
import store_v1 as st
import tray_v1 as tr


# ============================================================================
# 【使用说明】
#
# 作用
#     常驻后台，检测到 Prism 弹出 "Choose Color" 对话框时，在旁边浮出一个
#     取色面板。点色块或输入 hex，颜色立即写进对话框的 R/G/B 三个输入框。
#
# 环境
#     D:\\Conda\\envs\\dock\\python.exe   （含 pywin32 与 tkinter）
#
# 启动
#     python prism_color_tool_v1.py
#     或双击同目录的 启动Prism取色工具.bat / PrismColorTool.exe
#
# 交互
#     点色块            立即写入该颜色
#     输入 hex 后回车   写入（支持 #RRGGBB、RRGGBB、#RGB、rgb(r,g,b)）
#     吸管              全屏取色，左键确认，Esc 或右键取消
#     色板标题旁的 +    把当前色加入该色板，写回 palettes_v1.txt
#     色块右键          从自定义色板中移除
#     渐变 A → B        设置两端后按步数生成，点生成的色块应用
#     ×                 关闭面板（Prism 对话框不受影响）
#
# 自定义色板
#     存在同目录的 palettes_v1.txt，格式见该文件开头的注释。
#     改完需要重启工具生效。
# ============================================================================


# ============ 变量区 ============
POLL_MS = 250
PANEL_WIDTH = 272
SWATCH_SIZE = 26
SWATCH_GAP = 5
SWATCH_PER_ROW = 8
MAX_SCROLL_HEIGHT = 400       # 色板区最大高度，超出则出现滚动条
SCROLL_SCREEN_RATIO = 0.45    # 小屏幕上按屏幕高度的比例收缩

WS_EX_NOACTIVATE = 0x08000000
GWL_EXSTYLE = -20
GWLP_HWNDPARENT = -8
MUTEX_NAME = "PrismColorTool_SingleInstance_v1"
_MUTEX_HANDLE = None

BG = "#2B2B2B"
BG_INPUT = "#1E1E1E"
FG = "#E6E6E6"
FG_SUB = "#9A9A9A"
FG_ACCENT = "#7FB3D5"
BORDER = "#3C3C3C"


def set_noactivate(window, enable):
    """控制窗口是否抢焦点。NOACTIVATE 下点击不会把焦点从 Prism 抢走。"""
    try:
        hwnd = int(window.wm_frame(), 16)
    except Exception:
        return
    try:
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if enable:
            style |= WS_EX_NOACTIVATE
        else:
            style &= ~WS_EX_NOACTIVATE
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
    except Exception:
        pass


class Panel:
    """浮在颜色对话框旁边的取色面板。"""

    def __init__(self, tool, hdlg):
        self.tool = tool
        self.hdlg = hdlg
        self.last_rect = None
        self.grad_a = None
        self.grad_b = None
        self.last_rgb = None       # 上次从对话框读到的颜色，用于判断是否变化
        self.last_alpha = None     # 上次读到的透明度
        self._typing = False       # 用户正在 hex 框里打字时不要覆盖
        self._syncing = False      # 正在把对话框的值灌进滑块，不要回写
        self._alpha_job = None     # 透明度防抖任务
        self._panel_size = st.load_size()   # 上次记住的尺寸，None 表示按内容自动
        self._size_origin = None   # 拖动调整大小的起点
        self._scroll_first = 0.0   # 自绘滚动条：可见区域起点比例
        self._scroll_last = 1.0    # 自绘滚动条：可见区域终点比例
        self._thumb_box = None     # 滑块当前的 (top, bottom)
        self._drag_offset = 0      # 拖动滑块时的抓取偏移

        self.win = tk.Toplevel(tool.root)
        # 初始化任何一步抛异常，都必须把已创建的窗口销毁。
        # 否则 self.panel 仍是 None，下一轮轮询又建一个，会堆出多个孤儿面板。
        try:
            self.win.withdraw()
            self.win.overrideredirect(True)
            self.win.configure(bg=BORDER)

            self._build()
            self._place_near_dialog()
            self.win.deiconify()
            self.win.update_idletasks()
            set_noactivate(self.win, True)
            if not self.attach_owner():
                self.win.attributes("-topmost", True)
            self.refresh_from_dialog()
        except Exception:
            try:
                self.win.destroy()
            except Exception:
                pass
            raise

    def attach_owner(self):
        """把面板设为颜色对话框的从属窗口，层级跟着对话框走。"""
        try:
            hwnd = int(self.win.wm_frame(), 16)
            ctypes.windll.user32.SetWindowLongPtrW(hwnd, GWLP_HWNDPARENT, self.hdlg)
            return True
        except Exception:
            return False

    def sync_position(self):
        try:
            import win32gui
            rect = win32gui.GetWindowRect(self.hdlg)
        except Exception:
            return
        if rect == self.last_rect:
            return
        self.last_rect = rect
        self._place_near_dialog()

    # ---------- 界面骨架 ----------

    def _build(self):
        outer = tk.Frame(self.win, bg=BG)
        outer.pack(fill="both", expand=True, padx=1, pady=1)

        self._build_header(outer)
        self._build_scroll_area(outer)
        self._build_alpha_row(outer)
        self._build_footer(outer)
        self._build_resize_grip(outer)

        self._fill_content()
        self._place_near_dialog()
        self._sync_scroll()

    def _build_header(self, parent):
        top = tk.Frame(parent, bg=BG)
        top.pack(fill="x", padx=10, pady=(8, 6))

        self.preview = tk.Frame(top, bg="#000000", width=34, height=34,
                                highlightthickness=1, highlightbackground=BORDER)
        self.preview.pack(side="left")
        self.preview.pack_propagate(False)

        self.entry = tk.Entry(top, bg=BG_INPUT, fg=FG, relief="flat",
                              insertbackground=FG, width=11,
                              font=("Consolas", 11), justify="center")
        self.entry.pack(side="left", padx=(8, 0), ipady=5)
        self.entry.bind("<Return>", self._on_entry_return)
        self.entry.bind("<Button-1>", self._on_entry_click)
        self.entry.bind("<FocusOut>", self._on_entry_blur)

        close = tk.Label(top, text="✕", bg=BG, fg=FG_SUB,
                         font=("Segoe UI", 10), cursor="hand2", padx=6)
        close.pack(side="right")
        close.bind("<Button-1>", lambda e: self.tool.hide_panel())

        picker = tk.Label(top, text="吸管", bg=BG, fg=FG_ACCENT,
                          font=("Microsoft YaHei UI", 9), cursor="hand2", padx=4)
        picker.pack(side="right", padx=(0, 4))
        picker.bind("<Button-1>", lambda e: self.start_picker())

    def _build_scroll_area(self, parent):
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(body, bg=BG, highlightthickness=0, bd=0,
                                width=PANEL_WIDTH - 26)
        self.canvas.pack(side="left", fill="both", expand=True)

        # 自绘滚动条：一条 6 像素的窄轨 + 圆角色块。
        # 系统自带的 tk.Scrollbar 是经典 3D 凸起样式，在深色面板里很突兀，
        # 而且无法改配色，所以这里整个自己画。
        self.scroll_track = tk.Canvas(body, width=6, bg=BG,
                                      highlightthickness=0, bd=0,
                                      cursor="arrow")
        self.scroll_track.pack(side="right", fill="y", padx=(2, 3))

        self.inner = tk.Frame(self.canvas, bg=BG, padx=10)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner,
                                                  anchor="nw")
        self.inner.bind("<Configure>", lambda e: self._on_inner_resize())
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.configure(yscrollcommand=self._on_canvas_scroll)

        self.scroll_track.bind("<Button-1>", self._on_track_press)
        self.scroll_track.bind("<B1-Motion>", self._on_track_drag)
        self.scroll_track.bind("<ButtonRelease-1>", self._on_track_release)

    def _build_alpha_row(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", padx=10, pady=(4, 0))

        tk.Label(row, text="透明度", bg=BG, fg=FG_SUB,
                 font=("Microsoft YaHei UI", 8)).pack(side="left")

        self.alpha_var = tk.IntVar(value=0)
        self.alpha_scale = tk.Scale(
            row, from_=0, to=100, orient="horizontal",
            variable=self.alpha_var, command=self._on_alpha_change,
            bg=BG, fg=FG, troughcolor="#1B1B1B", highlightthickness=0,
            bd=0, relief="flat", sliderrelief="flat",
            showvalue=False, length=96, width=10)
        self.alpha_scale.pack(side="left", fill="x", expand=True, padx=6)

        # 数值框可直接输入，回车或失焦生效；和滑块双向联动
        self.alpha_entry = tk.Entry(
            row, width=4, bg=BG_INPUT, fg=FG, relief="flat",
            insertbackground=FG, font=("Consolas", 9), justify="center")
        self.alpha_entry.insert(0, "0")
        self.alpha_entry.pack(side="left", ipady=2)
        self.alpha_entry.bind("<Return>", self._on_alpha_entry_return)
        self.alpha_entry.bind("<FocusOut>", self._on_alpha_entry_blur)
        self.alpha_entry.bind("<Button-1>", self._on_alpha_entry_click)

        tk.Label(row, text="%", bg=BG, fg=FG_SUB,
                 font=("Consolas", 9)).pack(side="left", padx=(1, 0))

    def _set_alpha_display(self, alpha):
        """同步滑块与输入框的显示，不触发写回。"""
        self._syncing = True
        self.alpha_var.set(alpha)
        self.alpha_entry.delete(0, "end")
        self.alpha_entry.insert(0, str(alpha))
        self._syncing = False

    def _on_alpha_change(self, value):
        if self._syncing:
            return
        alpha = int(float(value))
        self.alpha_entry.delete(0, "end")
        self.alpha_entry.insert(0, str(alpha))
        # 拖动时持续触发，用短延迟合并成一次写入，避免每帧都发跨进程消息
        if self._alpha_job:
            try:
                self.win.after_cancel(self._alpha_job)
            except Exception:
                pass
        self._alpha_job = self.win.after(120, self._commit_alpha)

    def _commit_alpha(self):
        self._alpha_job = None
        self._apply_alpha(int(self.alpha_var.get()))

    def _apply_alpha(self, alpha):
        alpha = max(0, min(100, int(alpha)))
        if dio.write_transparency(self.hdlg, alpha):
            self.last_alpha = alpha
            self._refresh_alpha_text()
        else:
            self.status.config(text="透明度写入失败，对话框可能已关闭")

    def _on_alpha_entry_click(self, _event):
        set_noactivate(self.win, False)
        self.alpha_entry.focus_set()

    def _on_alpha_entry_return(self, _event):
        self._apply_alpha_from_entry()
        return "break"

    def _on_alpha_entry_blur(self, _event):
        self._apply_alpha_from_entry()
        set_noactivate(self.win, True)

    def _apply_alpha_from_entry(self):
        raw = self.alpha_entry.get().strip().replace("%", "")
        try:
            alpha = max(0, min(100, int(float(raw))))
        except ValueError:
            alpha = self.alpha_var.get()
        self._set_alpha_display(alpha)
        self._apply_alpha(alpha)

    def _build_footer(self, parent):
        self.status = tk.Label(parent, text="", bg=BG, fg=FG_SUB,
                               font=("Microsoft YaHei UI", 8), anchor="w",
                               justify="left")
        self.status.pack(fill="x", padx=10, pady=(2, 6))

    def _fill_content(self):
        """滚动区内只有自定义色板与渐变。内置预设色板已按要求移除。"""
        custom = st.load_palettes()
        if not custom:
            hint = tk.Label(self.inner,
                            text="palettes_v1.txt 里还没有色板\n"
                                 "用记事本编辑该文件，或点色板标题旁的 + 收藏颜色",
                            bg=BG, fg="#6A6A6A",
                            font=("Microsoft YaHei UI", 8),
                            anchor="w", justify="left")
            hint.pack(fill="x", pady=(8, 0))
        for palette in custom:
            self._add_palette_block(self.inner, palette["name"],
                                    palette["colors"], removable=True,
                                    apply_alpha=True)
        self._add_new_palette_button(self.inner)
        self._add_gradient_block(self.inner)
        # 滚轮事件不冒泡，创建完控件后统一递归绑定
        self._bind_wheel_tree(self.inner)

    def _build_resize_grip(self, parent):
        """右下角的拖拽点，按住拖动可调整面板大小。"""
        self.grip = tk.Canvas(parent, width=12, height=12, bg=BG,
                              highlightthickness=0, bd=0, cursor="size_nw_se")
        self.grip.place(relx=1.0, rely=1.0, anchor="se", x=-3, y=-3)
        for offset in (0, 4, 8):
            self.grip.create_line(offset + 2, 10, 10, offset + 2,
                                  fill="#6A6A6A", width=1)
        self.grip.bind("<Button-1>", self._on_grip_press)
        self.grip.bind("<B1-Motion>", self._on_grip_drag)
        self.grip.bind("<ButtonRelease-1>", self._on_grip_release)

    def _on_grip_press(self, event):
        self._size_origin = (event.x_root, event.y_root,
                             self.win.winfo_width(), self.win.winfo_height())

    def _on_grip_drag(self, event):
        if not self._size_origin:
            return
        x0, y0, w0, h0 = self._size_origin
        width = max(240, w0 + (event.x_root - x0))
        height = max(240, h0 + (event.y_root - y0))
        self._panel_size = (width, height)
        # 只改尺寸，位置由 owner 关系跟随对话框
        self.win.geometry("%dx%d" % (width, height))

    def _on_grip_release(self, _event):
        """松手时才写文件，拖动过程中不落盘。"""
        if self._panel_size:
            if st.save_size(*self._panel_size):
                self.status.config(text="面板尺寸已记住（%d × %d）"
                                        % self._panel_size)
            else:
                self.status.config(text="尺寸保存失败")

    # ---------- 色板块 ----------

    def _add_palette_block(self, parent, name, colors, removable,
                           apply_alpha=False):
        header = tk.Frame(parent, bg=BG)
        header.pack(fill="x", pady=(8, 3))

        tk.Label(header, text=name, bg=BG, fg=FG_SUB,
                 font=("Microsoft YaHei UI", 8)).pack(side="left")

        if removable:
            add = tk.Label(header, text="+", bg=BG, fg=FG_ACCENT,
                           font=("Consolas", 10, "bold"), cursor="hand2", padx=4)
            add.pack(side="right")
            add.bind("<Button-1>", lambda e, n=name: self._add_to_palette(n))

        grid = tk.Frame(parent, bg=BG)
        grid.pack(fill="x")
        if not colors:
            tk.Label(grid, text="（空）", bg=BG, fg="#6A6A6A",
                     font=("Microsoft YaHei UI", 8)).grid(row=0, column=0, sticky="w")
            return
        for index, item in enumerate(colors):
            # 自定义色板允许 '#RRGGBB@透明度' 写法，内置色板不带透明度
            if apply_alpha:
                hex_value, alpha = st.split_color(item)
            else:
                hex_value, alpha = item, None
            rgb = cc.hex_to_rgb(hex_value)
            if rgb is None:
                continue
            swatch = tk.Frame(grid, bg=hex_value, width=SWATCH_SIZE,
                              height=SWATCH_SIZE, cursor="hand2")
            swatch.grid(row=index // SWATCH_PER_ROW,
                        column=index % SWATCH_PER_ROW,
                        padx=(0, SWATCH_GAP), pady=(0, SWATCH_GAP), sticky="w")
            swatch.pack_propagate(False)
            swatch.bind("<Button-1>",
                        lambda e, c=rgb, h=hex_value, a=alpha: self.apply(c, h, a))
            if removable:
                swatch.bind("<Button-3>",
                            lambda e, it=item: self._remove_color(it))

    def _add_to_palette(self, palette_name):
        rgb = cc.parse_color(self.entry.get().strip())
        if not rgb:
            self.status.config(text="当前色无效，无法加入")
            return
        hex_value = cc.rgb_to_hex(*rgb)
        # 连同当前透明度一起存，写进 txt 时格式为 #RRGGBB@透明度
        alpha = self.last_alpha or 0
        item = "%s@%d" % (hex_value, alpha) if alpha else hex_value
        added, _ = st.add_color(item, palette_name)
        self._reload_palettes()
        self.status.config(
            text="已加入「%s」：%s" % (palette_name, item) if added
            else "「%s」里已有该颜色" % palette_name)

    def _remove_color(self, hex_value):
        removed, _ = st.remove_color(hex_value)
        if removed:
            self._reload_palettes()
            self.status.config(text="已从色板移除 %s" % hex_value)

    def _reload_palettes(self):
        """重绘整个滚动区，让自定义色板的增删立即反映出来。"""
        for child in self.inner.winfo_children():
            child.destroy()
        self._fill_content()
        self.inner.update_idletasks()
        self._sync_scroll()
        self._place_near_dialog()

    # ---------- 渐变 ----------

    def _add_new_palette_button(self, parent):
        """
        色板列表下方的"新建色板"入口。
        点击后在原位展开一个输入行，而不是弹系统对话框——
        simpledialog 是白底系统样式，和深色面板格格不入。
        """
        self.new_wrap = tk.Frame(parent, bg=BG)
        self.new_wrap.pack(fill="x", pady=(10, 0))

        self.new_btn = tk.Label(self.new_wrap, text="＋ 新建色板", bg=BG,
                                fg=FG_ACCENT, font=("Microsoft YaHei UI", 8),
                                cursor="hand2", anchor="w")
        self.new_btn.pack(fill="x")
        self.new_btn.bind("<Button-1>", lambda e: self._show_new_palette_row())

        self.new_row = tk.Frame(self.new_wrap, bg=BG)
        self.new_entry = tk.Entry(self.new_row, bg=BG_INPUT, fg=FG,
                                  relief="flat", insertbackground=FG,
                                  font=("Microsoft YaHei UI", 8))
        self.new_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self.new_entry.bind("<Return>", lambda e: self._confirm_new_palette())
        self.new_entry.bind("<Escape>", lambda e: self._cancel_new_palette())

        ok = tk.Label(self.new_row, text="确定", bg="#3C3C3C", fg=FG, padx=6,
                      font=("Microsoft YaHei UI", 8), cursor="hand2")
        ok.pack(side="left", padx=(4, 0))
        ok.bind("<Button-1>", lambda e: self._confirm_new_palette())

        cancel = tk.Label(self.new_row, text="取消", bg=BG, fg=FG_SUB, padx=4,
                          font=("Microsoft YaHei UI", 8), cursor="hand2")
        cancel.pack(side="left")
        cancel.bind("<Button-1>", lambda e: self._cancel_new_palette())

    def _show_new_palette_row(self):
        self.new_btn.pack_forget()
        self.new_row.pack(fill="x")
        self.new_entry.delete(0, "end")
        set_noactivate(self.win, False)
        self.new_entry.focus_set()
        self.inner.update_idletasks()
        self._sync_scroll()
        self._place_near_dialog()

    def _cancel_new_palette(self):
        self.new_row.pack_forget()
        self.new_btn.pack(fill="x")
        set_noactivate(self.win, True)
        self.inner.update_idletasks()
        self._sync_scroll()
        self._place_near_dialog()

    def _confirm_new_palette(self):
        name = self.new_entry.get().strip()
        if not name:
            self._cancel_new_palette()
            return
        created, _ = st.create_palette(name)
        if created:
            set_noactivate(self.win, True)
            self.status.config(text="已新建「%s」，点它右边的 + 往里加色" % name)
            self._reload_palettes()
        else:
            self.status.config(text="「%s」已存在，换个名字" % name)
            self.new_entry.select_range(0, "end")
            self.new_entry.focus_set()

    def _add_gradient_block(self, parent):
        divider = tk.Frame(parent, bg=BORDER, height=1)
        divider.pack(fill="x", pady=(12, 0))

        header = tk.Frame(parent, bg=BG)
        header.pack(fill="x", pady=(8, 3))
        tk.Label(header, text="渐变", bg=BG, fg=FG_SUB,
                 font=("Microsoft YaHei UI", 8)).pack(side="left")

        self.btn_a = tk.Label(header, text="A", bg="#444444", fg=FG, width=3,
                              font=("Consolas", 9), cursor="hand2")
        self.btn_a.pack(side="left", padx=(8, 0))
        self.btn_a.bind("<Button-1>", lambda e: self._set_endpoint("a"))

        tk.Label(header, text="→", bg=BG, fg=FG_SUB).pack(side="left", padx=3)

        self.btn_b = tk.Label(header, text="B", bg="#444444", fg=FG, width=3,
                              font=("Consolas", 9), cursor="hand2")
        self.btn_b.pack(side="left")
        self.btn_b.bind("<Button-1>", lambda e: self._set_endpoint("b"))

        self.steps = tk.Entry(header, width=3, bg=BG_INPUT, fg=FG, relief="flat",
                              justify="center", font=("Consolas", 9))
        self.steps.insert(0, "5")
        self.steps.pack(side="left", padx=(8, 2), ipady=2)
        tk.Label(header, text="色", bg=BG, fg=FG_SUB,
                 font=("Microsoft YaHei UI", 8)).pack(side="left")

        generate = tk.Label(header, text="生成", bg="#3C3C3C", fg=FG, padx=6,
                            cursor="hand2", font=("Microsoft YaHei UI", 8))
        generate.pack(side="right")
        generate.bind("<Button-1>", lambda e: self._generate_gradient())

        self.grad_row = tk.Frame(parent, bg=BG)
        self.grad_row.pack(fill="x", pady=(4, 0))

    def _set_endpoint(self, which):
        rgb = cc.parse_color(self.entry.get().strip())
        if not rgb:
            self.status.config(text="当前色无效，无法设为端点")
            return
        hex_value = cc.rgb_to_hex(*rgb)
        if which == "a":
            self.grad_a = rgb
            self.btn_a.config(bg=hex_value, text="")
        else:
            self.grad_b = rgb
            self.btn_b.config(bg=hex_value, text="")
        self.status.config(text="端点 %s 设为 %s" % (which.upper(), hex_value))

    def _generate_gradient(self):
        if not self.grad_a or not self.grad_b:
            self.status.config(text="请先点 A 和 B，把当前色设为两端")
            return
        try:
            steps = int(self.steps.get())
        except ValueError:
            steps = 5
        steps = max(2, min(20, steps))
        colors = cc.make_gradient(self.grad_a, self.grad_b, steps, "lab")

        for child in self.grad_row.winfo_children():
            child.destroy()
        width = max(10, min(34, (PANEL_WIDTH - 24) // len(colors) - 3))
        for rgb in colors:
            hex_value = cc.rgb_to_hex(*rgb)
            swatch = tk.Frame(self.grad_row, bg=hex_value, width=width,
                              height=22, cursor="hand2")
            swatch.pack(side="left", padx=(0, 3))
            swatch.pack_propagate(False)
            swatch.bind("<Button-1>", lambda e, c=rgb, h=hex_value: self.apply(c, h))
        self.status.config(text="已生成 %d 色渐变（LAB 插值），点色块应用" % len(colors))
        self.inner.update_idletasks()
        self._sync_scroll()
        self._place_near_dialog()

    # ---------- 滚动 ----------

    def _on_inner_resize(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._draw_thumb()

    def _on_canvas_resize(self, event):
        self.canvas.itemconfig(self.inner_id, width=event.width)
        self._draw_thumb()

    # ---------- 自绘滚动条 ----------

    def _on_canvas_scroll(self, first, last):
        """canvas 滚动时回调，记录可见区比例并重绘滑块。"""
        self._scroll_first = float(first)
        self._scroll_last = float(last)
        self._draw_thumb()

    def _draw_thumb(self):
        self.scroll_track.delete("all")
        height = self.scroll_track.winfo_height()
        if height <= 4:
            self._thumb_box = None
            return
        # 内容装得下就不画滑块
        if self._scroll_last - self._scroll_first >= 0.999:
            self._thumb_box = None
            return
        top = self._scroll_first * height
        bottom = self._scroll_last * height
        if bottom - top < 20:            # 滑块太短不好抓，给个下限
            center = (top + bottom) / 2
            top = max(0, center - 10)
            bottom = min(height, top + 20)
            top = max(0, bottom - 20)
        self._thumb_box = (top, bottom)
        self.scroll_track.create_rectangle(1, top, 5, bottom,
                                           fill="#5C5C5C", outline="")

    def _on_track_press(self, event):
        if self._thumb_box and self._thumb_box[0] <= event.y <= self._thumb_box[1]:
            self._drag_offset = event.y - self._thumb_box[0]
        else:
            # 点在空白轨道上：让滑块中心跳到点击位置
            self._drag_offset = 10
            self._scroll_to_thumb_top(event.y - self._drag_offset)

    def _on_track_drag(self, event):
        if not self._thumb_box:
            return
        self._scroll_to_thumb_top(event.y - self._drag_offset)

    def _on_track_release(self, _event):
        self._drag_offset = 0

    def _scroll_to_thumb_top(self, top):
        height = self.scroll_track.winfo_height()
        thumb_height = (self._thumb_box[1] - self._thumb_box[0]) if self._thumb_box else 20
        travel = height - thumb_height
        if travel <= 0:
            return
        fraction = max(0.0, min(1.0, top / travel))
        span = self._scroll_last - self._scroll_first
        self.canvas.yview_moveto(fraction * max(0.0, 1.0 - span))

    def _on_wheel(self, event):
        if self.inner.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _bind_wheel_tree(self, widget):
        """
        滚轮事件不会从子控件冒泡到 canvas，所以要把整棵控件树都绑上，
        否则鼠标停在色块上滚轮就没反应。
        """
        widget.bind("<MouseWheel>", self._on_wheel)
        for child in widget.winfo_children():
            self._bind_wheel_tree(child)

    def _sync_scroll(self):
        """按内容高度调整色板区高度；用户手动拉过大小后不再自动调整。"""
        self.inner.update_idletasks()
        needed = self.inner.winfo_reqheight()
        if not self._panel_size:
            limit = min(MAX_SCROLL_HEIGHT,
                        int(self.win.winfo_screenheight() * SCROLL_SCREEN_RATIO))
            view = min(needed, limit)
            self.canvas.configure(height=max(80, view))
        self.canvas.configure(
            scrollregion=(0, 0, self.canvas.winfo_width(), needed))
        self._draw_thumb()

    # ---------- 位置 ----------

    def _place_near_dialog(self):
        try:
            left, top, right, bottom = self.tool.dialog_rect(self.hdlg)
        except Exception:
            left, top, right, bottom = 100, 100, 600, 500
        self.win.update_idletasks()
        try:
            scale = self.win.winfo_fpixels("1i") / 96.0
        except Exception:
            scale = 1.0
        screen_w = self.win.winfo_screenwidth()
        screen_h = self.win.winfo_screenheight()

        if self._panel_size:
            # 用记住的尺寸，但不能超出屏幕可用区域
            width = min(self._panel_size[0], screen_w - 20)
            height = min(self._panel_size[1], screen_h - 60)
        else:
            width = int(PANEL_WIDTH * scale)
            height = int(self.win.winfo_reqheight() * scale)

        x = right + 8
        if x + width > screen_w:
            x = left - width - 8
        if x < 0:
            x = 8
        y = top
        if y + height > screen_h:
            y = max(8, screen_h - height - 40)
        self.win.geometry("%dx%d+%d+%d" % (width, height, x, y))

    # ---------- 交互 ----------

    def _on_entry_click(self, _event):
        set_noactivate(self.win, False)
        self.entry.focus_set()
        # 打字期间暂停从对话框回灌，否则每 250 毫秒就会被覆盖
        self._typing = True

    def _on_entry_blur(self, _event):
        self._typing = False
        set_noactivate(self.win, True)

    def _on_entry_return(self, _event):
        value = cc.parse_color(self.entry.get())
        if value:
            self.apply(value, cc.rgb_to_hex(*value))
        else:
            self.status.config(text="无法识别的颜色：%s" % self.entry.get())
        return "break"

    def apply(self, rgb, hex_value="", alpha=None):
        """
        写入颜色。alpha 为 None 表示保持对话框当前透明度不变；
        为 0-100 时一并设置透明度（自定义色板里带了 @ 的项会这样传）。
        """
        ok, actual = dio.write_rgb(self.hdlg, *rgb)
        if not ok:
            self.status.config(text="写入失败，对话框可能已关闭")
            return
        if alpha is not None:
            dio.write_transparency(self.hdlg, alpha)
            self.last_alpha = alpha
            self._set_alpha_display(alpha)
        self.last_rgb = actual
        self._set_preview(actual)
        self.entry.delete(0, "end")
        self.entry.insert(0, hex_value or cc.rgb_to_hex(*actual))
        self._refresh_alpha_text()

    def _set_preview(self, rgb):
        self.preview.configure(bg=cc.rgb_to_hex(*rgb))

    def _refresh_alpha_text(self):
        """透明度与当前颜色分两条状态显示，互不覆盖。"""
        parts = []
        if self.last_rgb:
            parts.append(cc.rgb_to_hex(*self.last_rgb))
        if self.last_alpha is not None:
            parts.append("透明度 %d%%（不透明 %d%%）"
                         % (self.last_alpha, 100 - self.last_alpha))
            if self.last_alpha > 0:
                parts.append("颜色会被稀释")
        self.status.config(text="　".join(parts))

    def refresh_from_dialog(self):
        """首次显示时调用一次，把对话框当前状态搬过来。"""
        self.sync_from_dialog()

    def sync_from_dialog(self):
        """
        把对话框里的当前颜色与透明度同步到面板。
        由主循环每 250 毫秒调用一次，因此用户在 Prism 里点色块、拖色轮、
        改数值，面板都会跟着变。
        """
        rgb = dio.read_rgb(self.hdlg)
        if rgb and rgb != self.last_rgb:
            self.last_rgb = rgb
            self._set_preview(rgb)
            if not self._typing:
                self.entry.delete(0, "end")
                self.entry.insert(0, cc.rgb_to_hex(*rgb))
            self._refresh_alpha_text()

        alpha = dio.read_transparency(self.hdlg)
        if alpha != self.last_alpha:
            self.last_alpha = alpha
            if alpha is not None:
                # 同步滑块与输入框，但不要让 set() 反过来再写一次对话框
                self._set_alpha_display(alpha)
            self._refresh_alpha_text()

    # ---------- 吸管 ----------

    def start_picker(self):
        from picker_v1 import Picker
        self.win.withdraw()
        self.picker = Picker(self.tool.root, self._on_picked, self._on_pick_cancel)

    def _restore_panel(self):
        try:
            self.win.deiconify()
            self.attach_owner()
            set_noactivate(self.win, True)
        except Exception:
            pass

    def _on_picked(self, rgb, spread):
        self._restore_panel()
        self.apply(rgb, cc.rgb_to_hex(*rgb))
        if spread > 12:
            self.status.config(text="写入 %s（邻域色差 %d）"
                                    % (cc.rgb_to_hex(*rgb), spread))

    def _on_pick_cancel(self):
        self._restore_panel()
        self.status.config(text="已取消取色")

    def destroy(self):
        try:
            self.win.destroy()
        except Exception:
            pass


class Tool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.panel = None
        self.hdlg = 0
        self.tray = None
        self._quit_requested = False
        self._retry_after = 0.0   # 面板创建连续失败时的冷却时间戳

    def request_quit(self):
        """由托盘线程调用。只设标志，真正的退出交给主线程执行。"""
        self._quit_requested = True

    def quit_app(self):
        try:
            if self.panel is not None:
                self.panel.destroy()
                self.panel = None
            if self.tray is not None:
                self.tray.stop()
                self.tray = None
        finally:
            try:
                self.root.quit()
                self.root.destroy()
            except Exception:
                pass

    def dialog_rect(self, hdlg):
        import win32gui
        return win32gui.GetWindowRect(hdlg)

    def poll(self):
        if self._quit_requested:
            self.quit_app()
            return

        # 创建失败后冷却 5 秒再试。没有这个冷却，一旦某台机器上初始化出错，
        # 每 250 毫秒就会建一个窗口又销毁，表现为面板疯狂闪现或堆叠。
        if time.time() < self._retry_after:
            self.root.after(POLL_MS, self.poll)
            return

        try:
            if self.panel is not None:
                import win32gui
                if (not win32gui.IsWindow(self.hdlg)
                        or not win32gui.IsWindowVisible(self.hdlg)):
                    self.hide_panel()
                else:
                    self.panel.sync_position()
                    self.panel.sync_from_dialog()
            else:
                hdlg = dio.find_color_dialog()
                if hdlg:
                    self.hdlg = hdlg
                    try:
                        self.panel = Panel(self, hdlg)
                    except Exception:
                        self.panel = None
                        self.hdlg = 0
                        self._retry_after = time.time() + 5.0
                        raise
        except Exception:
            self.hide_panel()
        self.root.after(POLL_MS, self.poll)

    def hide_panel(self):
        if self.panel is not None:
            self.panel.destroy()
            self.panel = None
            self.hdlg = 0

    def run(self):
        print("Prism 颜色工具已启动。")
        print("现在去 Prism 里打开任意一个颜色对话框，面板会自动出现在旁边。")
        print("自定义色板文件：%s" % st.CONFIG_FILE)
        self.tray = tr.TrayIcon(self.request_quit)
        if self.tray.start():
            print("托盘图标已就绪：右下角托盘区，右键可退出。")
        else:
            print("托盘图标注册失败，退出请用任务管理器结束进程。")
        self.root.after(100, self.poll)
        self.root.mainloop()


def enable_dpi_awareness():
    """让进程按物理像素工作，否则鼠标坐标与抓屏像素对不上。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return "shcore(1)"
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        return "user32"
    except Exception:
        return "失败"


def acquire_single_instance():
    """已有实例在跑时返回 False，避免开出多个面板。"""
    try:
        global _MUTEX_HANDLE
        _MUTEX_HANDLE = win32event.CreateMutex(None, False, MUTEX_NAME)
        return win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS
    except Exception:
        return True


def main():
    if sys.platform != "win32":
        print("本工具仅支持 Windows。")
        return 1
    if not acquire_single_instance():
        try:
            ctypes.windll.user32.MessageBoxW(
                0,
                "Prism 取色工具已经在运行了。\n\n"
                "面板会在你打开 Prism 颜色对话框时自动出现。",
                "已在运行", 0x40)
        except Exception:
            pass
        return 0
    mode = enable_dpi_awareness()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("DPI 感知：%s" % mode)
    print("进程视角的屏幕尺寸：%d × %d"
          % (win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)))
    Tool().run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
