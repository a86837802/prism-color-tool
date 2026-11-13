# Prism 取色工具

给 GraphPad Prism 的颜色对话框加一个取色面板。点色块、输色号、屏幕吸管、自动生成渐变，不用再逐个输入 RGB 三个数字。

Windows 专用，适用于 Prism 7–11。

## 功能

- 点色块直接写入 Prism
- 支持 `#RRGGBB`、`rgb(r,g,b)` 等格式
- 屏幕吸管，带放大镜，可从论文插图吸色
- 两个端点色加步数，生成渐变
- 透明度滑块，与 Prism 双向同步
- 色板存纯文本，记事本可改

## 下载使用

[下载 PrismColorTool.exe](https://github.com/a86837802/prism-color-tool/releases/latest/download/PrismColorTool.exe)

双击运行，托盘出现图标，面板会在打开 Prism 颜色对话框时自动出现。退出：右键托盘图标 → 退出。

从源码运行：

```bash
pip install pywin32
python src/prism_color_tool_v1.py
```

## 用法

| 操作 | 结果 |
|---|---|
| 点色块 | 写入 Prism |
| 色号框回车 | 写入该颜色 |
| 点「吸管」 | 全屏取色，左键确认，Esc 取消 |
| 色板标题旁的 `+` | 收藏当前颜色 |
| 色块右键 | 从色板移除 |
| 拖右下角 | 调整面板大小，自动记住 |

## 配置

`palettes_v1.txt` 与程序同目录，首次运行自动生成：

```
>Nature配色
#1B9E77,#D95F02,#7570B3

>半透明标注
#2E5A88@50
```

`>` 开头是色板名，下一行是颜色，逗号分隔，`@` 后面是透明度。格式详见 `palettes.example.txt`。

## 兼容性

Windows 专用，不支持 macOS。Prism 7–11 均可，中英文版都认——工具操作的是 Windows 系统颜色对话框，不是 Prism 自己画的界面。

## 文档

- [设计方案](docs/设计方案.md)
- [技术细节](docs/技术细节.md)
- [PPT 嵌入对象结构分析](docs/PPT嵌入对象结构分析.md)

## 许可证

MIT
