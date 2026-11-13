# Prism 取色工具

给 GraphPad Prism 的颜色选择对话框加一个取色面板：点色块、输入色号、屏幕吸管、自动生成渐变，不再逐个输入 RGB 三个数字。

适用于 Windows 版 Prism，无需修改 Prism 安装文件。

---

## 解决的问题

在 Prism 里给数据点或柱子设一个非预设颜色，原本要走六步：

```
Format Graph → 颜色下拉 → More Colors → 先点一个基本颜色 → 逐个改 R、G、B → Add to Custom Colors → OK
```

其中「逐个改三个十进制数」是纯机械劳动。人想到的是"这个蓝"或 `#2E5A88`，不是 (46, 90, 136)。

这个工具把这六步压缩成一次点击或一次吸色。

---

## 功能

| 功能 | 说明 |
|---|---|
| **色板点选** | 自定义色板，点一下即写入 Prism |
| **色号输入** | 支持 `#RRGGBB`、`RRGGBB`、`#RGB`、`rgb(r,g,b)` |
| **屏幕吸管** | 全屏取色，带 10 倍放大镜，可从论文插图直接吸色 |
| **渐变生成** | 选两个端点色 + 步数，按 CIELAB 空间插值 |
| **透明度调节** | 滑块或直接输入数字，双向联动 |
| **实时同步** | 在 Prism 里改颜色或透明度，面板跟着变 |
| **色板文件** | 纯文本配置，记事本可编辑，支持分组与透明度标记 |
| **尺寸记忆** | 拖动调整面板大小，下次打开保持 |

---

## 快速开始

### 方式一：直接运行 exe（推荐给使用者）

1. 下载 [`PrismColorTool.exe`](https://github.com/a86837802/prism-color-tool/releases/latest/download/PrismColorTool.exe)
2. 双击运行，右下角托盘区会出现图标
3. 在 Prism 里打开任意颜色对话框，面板自动出现在旁边
4. 退出：右键托盘图标 → 退出

无需安装 Python 或任何依赖。首次启动约 2–4 秒（单文件模式需要解压）。

### 方式二：从源码运行（推荐给开发者）

```bash
pip install pywin32
python src/prism_color_tool_v1.py
```

或双击 `启动Prism取色工具.bat`（无控制台窗口）。

需要 Python 3.8 以上，带 tkinter（官方安装包默认包含）。

---

## 使用说明

### 面板结构

```
┌────────────────────────────────┐
│ ● #2E5A88        吸管     ×    │  当前色预览 / 色号输入
├────────────────────────────────┤
│ 我的收藏                  +    │  自定义色板
│ ■ ■ ■ ■                        │
│ Nature配色                +    │
│ ■ ■ ■ ■ ■ ■ ■ ■                │
│ ──────────────────────────     │
│ 渐变   A → B   [5] 色   生成   │  两端点 + 步数
│ ■ ■ ■ ■ ■                      │
├────────────────────────────────┤
│ 透明度 ────●──── [ 0 ] %       │
├────────────────────────────────┤
│ #2E5A88  透明度 0%（不透明 100%）│
└────────────────────────────────┘
```

### 常用操作

| 操作 | 结果 |
|---|---|
| 点色块 | 立即写入 Prism |
| 色号框输入后回车 | 写入该颜色 |
| 点「吸管」 | 进入全屏取色，左键确认，Esc 或右键取消 |
| 色板标题旁的 `+` | 把当前颜色（含透明度）存入该色板 |
| 色块右键 | 从色板中移除 |
| 渐变区点 A / B | 把当前色设为端点 |
| 拖动右下角三角 | 调整面板大小，松手即记住 |

### 关于透明度

Prism 的颜色对话框有一个透明度控件。**透明度不为 0 时，实际绘制的颜色会比你看到的纯色浅**——这是 Prism 的设定，不是工具造成的。面板底部会提示当前透明度。

---

## 配置文件

两个文件都在 exe（或源码根目录）同级，首次运行自动生成。

### palettes_v1.txt —— 自定义色板

```
; 分号开头是注释
>Nature配色
#1B9E77,#D95F02,#7570B3

>半透明标注
#2E5A88@50,#D55E00@30
```

规则：

- `>` 开头是一个色板的名字
- 紧随其后的行是颜色，逗号分隔，`#` 号可省略
- 一个色板可写多行，自动接续
- 颜色后缀 `@透明度`（0–100），省略表示不透明

示例文件见 `palettes.example.txt`。改完需要重启工具生效。

### settings_v1.txt —— 面板尺寸

```
width=380
height=520
```

删掉这个文件即恢复按内容自动调整大小。

---

## 兼容性

| 项目 | 支持情况 |
|---|---|
| 操作系统 | **仅 Windows**（依赖 Win32 对话框，macOS 版 Prism 不适用） |
| Prism 版本 | 7 / 8 / 9 / 10 / 11（均使用系统颜色对话框） |
| Prism 语言 | 英文版（`Choose Color`）、中文版（`选择颜色`） |
| 32/64 位 | 均可 |

官方文档自 Prism 7 起描述一致：颜色对话框 "is part of Windows and was not created by GraphPad"。更早的 Prism 4–6 官方 FAQ 称同样适用，但未实测。

---

## 技术原理

### 为什么不写插件

GraphPad 官方文档明确：

> "Does Prism have an API that can be called from other programs? No."

Prism 脚本只能开/关项目、导入数据、导出图表，无法控制绘图属性。所以不存在插件机制。

### 为什么改的是系统对话框

官方文档另一句：

> "This brings up the custom color dialog, which is part of Windows and was not created by GraphPad."

那个要输三个数字的对话框是 Windows 通用颜色对话框（`comdlg32.dll` 的 `ChooseColor`），窗口类名 `#32770`。它的控件 ID 定义在 Windows SDK 的 `dlgs.h`：

| 常量 | 值 | 用途 |
|---|---|---|
| `COLOR_RED` | 706 | 红通道输入框 |
| `COLOR_GREEN` | 707 | 绿通道输入框 |
| `COLOR_BLUE` | 708 | 蓝通道输入框 |
| `COLOR_ADD` | 712 | 添加到自定义颜色 |
| `IDOK` | 1 | 确定 |

工具做的事只有两件：给这三个输入框发 `WM_SETTEXT` 写值，再用 `WM_GETTEXT` 读回校验。不注入进程、不挂钩子、不改任何文件。

### 透明度为什么要走滑块

Prism 在标准对话框上追加了透明度控件：

- 1116 = `msctls_trackbar32`，位置 0–100
- 1117 = `Edit`，显示 `80%`

实测发现**只写 1117 文本框时滑块不动**，说明 Prism 读的是滑块位置。正确写法：

```
TBM_SETPOS(滑块, 1, 目标值)                     ← 设置位置
WM_HSCROLL(对话框, (SB_THUMBPOSITION<<16)|值, 滑块)  ← 通知对话框
```

### 吸管的两处坑

**抓屏缓冲区的行序**：`win32ui.GetBitmapBits()` 返回的像素数据是自上而下排列，与 Windows DIB 惯例（自下而上）相反。用 `height-1-y` 会导致取色点上下镜像。

**Tk 的坐标单位**：Tk 8.6 在 DPI 感知进程里，事件回调的 `x_root` 可能仍是逻辑坐标（如 1536×864），而抓屏缓冲是物理像素（1920×1080）。两者混用会让取色点系统性偏移，需改用 `win32api.GetCursorPos()`。

---

## 已知限制

1. **仅支持 Windows**。macOS 版 Prism 使用系统调色板，机制完全不同。
2. **只能改颜色与透明度**。线宽、字体、标记形状不在范围内。
3. **不解析 Prism 项目文件**。工具操作的是对话框，不是数据。
4. **单实例**。重复启动会提示"已在运行"。
5. **若 GraphPad 未来改为自绘颜色选择器**，本工具将失效。

---

## 目录结构

```
.
├── src/                        主程序
│   ├── prism_color_tool_v1.py  入口：轮询、面板调度、托盘
│   ├── dialog_io_v1.py         对话框识别与读写
│   ├── picker_v1.py            屏幕吸管
│   ├── color_convert_v1.py     颜色格式转换与 LAB 插值
│   ├── store_v1.py             色板与设置文件读写
│   ├── tray_v1.py              系统托盘
│   └── palette_data_v1.py      内置色板数据（当前未启用）
├── tools/                      调试与探测脚本
│   ├── dialog_probe_v1.py      探测颜色对话框控件结构
│   ├── write_probe_v1.py       验证颜色写入
│   ├── exp_input_v1.py         对比三种输入方式
│   ├── diag_edit_v1.py         诊断输入框样式
│   └── probe_ppt_ole_v1.py     分析 PPT 中嵌入的 Prism 对象
├── tests/                      自测脚本
├── docs/
│   ├── 设计方案.md
│   └── PPT嵌入对象结构分析.md
└── 启动Prism取色工具.bat
```

---

## 开发说明

### 打包

```bash
pyinstaller --onefile --windowed --name PrismColorTool \
  --distpath release --workpath .build --specpath .build \
  --paths src src/prism_color_tool_v1.py
```

**如果使用 conda 环境，务必先隔离 PATH**：

```bash
PATH="/path/to/env:/path/to/env/Library/bin:/c/Windows/System32:/c/Windows" \
  pyinstaller ...
```

否则 PyInstaller 会沿着 PATH 抓到其他 conda 环境的 Tcl/Tk 运行时，打出的 exe 启动时报版本冲突（`have 8.6.14, need exactly 8.6.13`）。

### 排查

工具没反应时，先跑 `tools/dialog_probe_v1.py`，它会把当前所有 `#32770` 窗口的控件树打印出来，据此判断控件 ID 是否与预期一致。

---

## 许可证

MIT License，详见 `LICENSE`。
