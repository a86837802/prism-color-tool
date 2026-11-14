# -*- coding: utf-8 -*-
"""store_v1.py —— 自定义色板的 txt 配置读写"""

import sys
from pathlib import Path


# ============================================================================
# 【使用说明】
#
# 配置文件：palettes_v1.txt，与程序同目录（exe 放哪就跟着哪）。
# 格式极简，直接用记事本改：
#
#   ; 分号开头是注释，不参与解析
#   >Nature配色
#   #1B9E77,#D95F02,#7570B3
#   >我的常用
#   #2E5A88,#D55E00
#
# 规则：
#   > 开头的行是一个色板的名字
#   紧跟的、不以 > 开头的行是颜色，用逗号分隔，# 号可省略
#   一个色板可以写多行颜色，会自动接续
#   空行忽略
#
# 单独查看解析结果：
#   python store_v1.py
# ============================================================================


# ============ 变量区 ============


def _base_dir():
    """
    配置目录的基准位置。
    打包成 exe（PyInstaller onefile）后，__file__ 指向临时解压目录，
    退出即删；此时必须改用 exe 自身所在目录，保持"配置跟着 exe 走"。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = _base_dir()
CONFIG_FILE = BASE_DIR / "palettes_v1.txt"
SETTINGS_FILE = BASE_DIR / "settings_v1.txt"
DEFAULT_PALETTE_NAME = "我的收藏"

SETTINGS_HEADER = ("; 面板设置。删掉这个文件即可恢复默认尺寸。\n"
                   "; 也可以手动改这里的数字，下次打开面板时生效。\n")

DEFAULT_TEXT = """; 色板配置。用记事本直接改，改完重启工具生效。
; 以 > 开头的是一个色板的名字，下一行写颜色，用逗号分隔，# 号可省略。
; 一个色板可以写多行，会自动接续。分号开头的行是注释。
;
; 颜色后面可以加 @透明度（0-100），例如：
;   #2E5A88        完全不透明
;   #2E5A88@50     50% 透明
;
; 例：
; >Nature配色
; #1B9E77,#D95F02,#7570B3
; >半透明标注
; #2E5A88@60,#D55E00@30

>我的收藏
#1B9E77
"""


def _normalize(token):
    """
    规范化成 '#RRGGBB' 或 '#RRGGBB@透明度'。
    透明度为 0 时省略 @ 部分（等价于完全不透明）。无效返回 None。
    """
    value = token.strip().strip(",").strip()
    if not value:
        return None
    alpha = None
    if "@" in value:
        value, _, raw_alpha = value.partition("@")
        value = value.strip()
        try:
            alpha = max(0, min(100, int(raw_alpha.strip())))
        except ValueError:
            alpha = None
    if not value.startswith("#"):
        value = "#" + value
    if len(value) != 7:
        return None
    body = value[1:]
    try:
        int(body, 16)
    except ValueError:
        return None
    result = "#" + body.upper()
    if alpha:
        result += "@%d" % alpha
    return result


def split_color(item):
    """把 '#RRGGBB@50' 拆成 ('#RRGGBB', 50)。没有 @ 时透明度为 0。"""
    if "@" in item:
        hex_value, _, raw = item.partition("@")
        try:
            return hex_value, max(0, min(100, int(raw)))
        except ValueError:
            return hex_value, 0
    return item, 0


def ensure_file():
    if CONFIG_FILE.exists():
        return
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(DEFAULT_TEXT, encoding="utf-8")


def _read_lines():
    ensure_file()
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return CONFIG_FILE.read_text(encoding=encoding).splitlines()
        except UnicodeDecodeError:
            continue
    return []


def load_palettes():
    """
    解析成 [{"name": 名称, "colors": ["#RRGGBB", ...]}, ...]。
    顺序保持文件中的顺序。解析失败的行跳过，不抛异常。
    """
    palettes = []
    current = None
    for raw in _read_lines():
        line = raw.strip()
        if not line or line.startswith(";"):
            continue
        if line.startswith(">"):
            name = line[1:].strip() or "未命名"
            current = {"name": name, "colors": []}
            palettes.append(current)
            continue
        if current is None:
            current = {"name": DEFAULT_PALETTE_NAME, "colors": []}
            palettes.append(current)
        for token in line.split(","):
            value = _normalize(token)
            if value and value not in current["colors"]:
                current["colors"].append(value)
    return palettes


def save_palettes(palettes):
    """按当前内存状态整体写回，保留注释头方便下次编辑。"""
    ensure_file()
    lines = [ln for ln in DEFAULT_TEXT.splitlines()
             if ln.startswith(";") or not ln.strip()]
    while lines and not lines[-1].strip():
        lines.pop()
    lines.append("")
    for palette in palettes:
        lines.append(">%s" % palette["name"])
        if palette["colors"]:
            lines.append(",".join(palette["colors"]))
        lines.append("")
    CONFIG_FILE.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def create_palette(name):
    """
    新建一个空色板，追加到列表末尾。
    返回 (是否创建成功, 当前色板列表)。重名或名称为空时返回 False。
    """
    name = (name or "").strip()
    if not name:
        return False, load_palettes()
    palettes = load_palettes()
    if any(palette["name"] == name for palette in palettes):
        return False, palettes
    palettes.append({"name": name, "colors": []})
    save_palettes(palettes)
    return True, palettes


def add_color(hex_value, palette_name=DEFAULT_PALETTE_NAME):
    """
    把颜色加入指定色板，色板不存在则创建在末尾。
    返回 (是否新增, 当前色板列表)。
    """
    value = _normalize(hex_value)
    if not value:
        return False, load_palettes()
    palettes = load_palettes()
    target = None
    for palette in palettes:
        if palette["name"] == palette_name:
            target = palette
            break
    if target is None:
        target = {"name": palette_name, "colors": []}
        palettes.append(target)
    if value in target["colors"]:
        return False, palettes
    target["colors"].append(value)
    save_palettes(palettes)
    return True, palettes


def remove_color(hex_value):
    """从所有色板里移除该颜色。返回 (是否移除, 当前色板列表)。"""
    value = _normalize(hex_value)
    if not value:
        return False, load_palettes()
    palettes = load_palettes()
    removed = False
    for palette in palettes:
        if value in palette["colors"]:
            palette["colors"].remove(value)
            removed = True
    if removed:
        palettes = [p for p in palettes if p["colors"] or p["name"]]
        save_palettes(palettes)
    return removed, palettes


# ============ 面板尺寸记忆 ============

def load_size():
    """
    读取上次调整的面板尺寸，返回 (width, height)。
    没有记录或内容无效时返回 None，让面板按内容自动决定大小。
    """
    if not SETTINGS_FILE.exists():
        return None
    try:
        text = SETTINGS_FILE.read_text(encoding="utf-8-sig")
    except Exception:
        return None
    width = height = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(";") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip().lower()
        try:
            number = int(float(value.strip()))
        except ValueError:
            continue
        if key == "width":
            width = number
        elif key == "height":
            height = number
    if width and height:
        return (max(240, width), max(240, height))
    return None


def save_size(width, height):
    """把当前面板尺寸写进设置文件，供下次启动恢复。"""
    try:
        SETTINGS_FILE.write_text(
            "%swidth=%d\nheight=%d\n" % (SETTINGS_HEADER, int(width), int(height)),
            encoding="utf-8")
        return True
    except Exception:
        return False


def clear_size():
    """清除尺寸记忆，恢复自动大小。"""
    try:
        if SETTINGS_FILE.exists():
            SETTINGS_FILE.unlink()
        return True
    except Exception:
        return False


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ensure_file()
    print("配置文件：%s" % CONFIG_FILE)
    print("")
    for palette in load_palettes():
        print(">%s（%d 色）" % (palette["name"], len(palette["colors"])))
        print("   " + ",".join(palette["colors"]))
