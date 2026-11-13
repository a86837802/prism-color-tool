# -*- coding: utf-8 -*-
"""color_convert_v1.py —— 颜色格式转换与解析"""

import re


# ============================================================================
# 【使用说明】
#
# 提供三种能力：
#   1. 把用户可能粘贴的颜色字符串解析成 (R, G, B)
#      支持 #RRGGBB、RRGGBB、#RGB、rgb(46,90,136)、rgba(46,90,136,0.5)
#   2. (R, G, B) 与 hex 互转
#   3. (R, G, B) 与 HSL 互转（标准尺度：H 0-360，S/L 0-100%）
#
# 单独测试：
#   python color_convert_v1.py "#2E5A88"
#   python color_convert_v1.py "rgb(46,90,136)"
# ============================================================================


# ============ 变量区 ============
HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")
# 三位缩写必须带 # 号，否则 "bad"、"abc" 这类普通单词会被误解析成颜色
HEX_SHORT_RE = re.compile(r"^#([0-9a-fA-F]{3})$")
RGB_RE = re.compile(
    r"^rgba?\(\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})", re.I)
CSS_NAMED = {
    "black": (0, 0, 0), "white": (255, 255, 255), "red": (255, 0, 0),
    "green": (0, 128, 0), "blue": (0, 0, 255), "gray": (128, 128, 128),
    "grey": (128, 128, 128), "orange": (255, 165, 0), "purple": (128, 0, 128),
}


def clamp(value, low=0, high=255):
    return max(low, min(high, int(round(value))))


def hex_to_rgb(text):
    """'#2E5A88' 或 '2E5A88' → (46, 90, 136)。失败返回 None。"""
    m = HEX_RE.match(text.strip())
    if not m:
        return None
    h = m.group(1)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def short_hex_to_rgb(text):
    """'#ABC' → (170, 187, 204)。"""
    m = HEX_SHORT_RE.match(text.strip())
    if not m:
        return None
    h = m.group(1)
    return (int(h[0] * 2, 16), int(h[1] * 2, 16), int(h[2] * 2, 16))


def rgb_to_hex(r, g, b, upper=True):
    s = "#%02X%02X%02X" if upper else "#%02x%02x%02x"
    return s % (clamp(r), clamp(g), clamp(b))


def parse_color(text):
    """
    解析任意常见颜色写法。返回 (r, g, b) 或 None。
    顺序：六位 hex → 三位 hex → rgb()/rgba() → 英文颜色名。
    """
    if not text:
        return None
    raw = text.strip()

    value = hex_to_rgb(raw)
    if value:
        return value
    value = short_hex_to_rgb(raw)
    if value:
        return value

    m = RGB_RE.match(raw)
    if m:
        return (clamp(int(m.group(1))), clamp(int(m.group(2))), clamp(int(m.group(3))))

    named = CSS_NAMED.get(raw.lower())
    if named:
        return named
    return None


def rgb_to_hsl(r, g, b):
    """返回 (h 0-360, s 0-100, l 0-100)。"""
    r, g, b = clamp(r) / 255.0, clamp(g) / 255.0, clamp(b) / 255.0
    mx, mn = max(r, g, b), min(r, g, b)
    delta = mx - mn
    lightness = (mx + mn) / 2.0

    if delta == 0:
        hue = 0.0
        saturation = 0.0
    else:
        saturation = delta / (1 - abs(2 * lightness - 1)) if lightness not in (0, 1) else 0.0
        if mx == r:
            hue = 60 * (((g - b) / delta) % 6)
        elif mx == g:
            hue = 60 * (((b - r) / delta) + 2)
        else:
            hue = 60 * (((r - g) / delta) + 4)
    return (round(hue, 1), round(saturation * 100, 1), round(lightness * 100, 1))


def hsl_to_rgb(h, s, l):
    """输入 H 0-360、S/L 0-100，返回 (r, g, b)。"""
    h = h % 360
    s = max(0.0, min(100.0, s)) / 100.0
    l = max(0.0, min(100.0, l)) / 100.0
    c = (1 - abs(2 * l - 1)) * s
    x = c * (1 - abs((h / 60) % 2 - 1))
    m = l - c / 2
    if h < 60:
        rp, gp, bp = c, x, 0
    elif h < 120:
        rp, gp, bp = x, c, 0
    elif h < 180:
        rp, gp, bp = 0, c, x
    elif h < 240:
        rp, gp, bp = 0, x, c
    elif h < 300:
        rp, gp, bp = x, 0, c
    else:
        rp, gp, bp = c, 0, x
    return (clamp((rp + m) * 255), clamp((gp + m) * 255), clamp((bp + m) * 255))


def relative_luminance(r, g, b):
    """WCAG 相对亮度，用于判断文字该用黑还是白。"""
    def channel(v):
        v = v / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def readable_text_color(r, g, b):
    """在给定底色上，返回对比度更高的黑白文字色。"""
    return "#000000" if relative_luminance(r, g, b) > 0.45 else "#FFFFFF"


# ============ CIELAB 转换与渐变 ============
# 为什么渐变不直接在 RGB 里插值：
#   RGB 是设备编码，不是感知均匀空间。从蓝到黄直接插值，中段会经过灰，
#   得到一串发闷的颜色。LAB 里两个色点之间走直线，中段亮度与饱和度
#   变化更接近人眼预期，适合做科研配色的连续梯度。

D65_WHITE = (0.95047, 1.0, 1.08883)


def _srgb_to_linear(value):
    value = value / 255.0
    return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4


def _linear_to_srgb(value):
    value = max(0.0, min(1.0, value))
    return 12.92 * value if value <= 0.0031308 else 1.055 * (value ** (1 / 2.4)) - 0.055


def rgb_to_lab(r, g, b):
    """sRGB → CIELAB（D65 白点）。"""
    rl, gl, bl = _srgb_to_linear(r), _srgb_to_linear(g), _srgb_to_linear(b)
    x = rl * 0.4124564 + gl * 0.3575761 + bl * 0.1804375
    y = rl * 0.2126729 + gl * 0.7151522 + bl * 0.0721750
    z = rl * 0.0193339 + gl * 0.1191920 + bl * 0.9503041

    xr, yr, zr = x / D65_WHITE[0], y / D65_WHITE[1], z / D65_WHITE[2]

    def f(t):
        return t ** (1 / 3) if t > 0.008856 else (7.787 * t + 16 / 116)

    fx, fy, fz = f(xr), f(yr), f(zr)
    return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz))


def lab_to_rgb(lightness, a_star, b_star):
    """CIELAB → sRGB。超出色域的分量会被裁剪到 0-255。"""
    fy = (lightness + 16) / 116
    fx = fy + a_star / 500
    fz = fy - b_star / 200

    def inv(t):
        return t ** 3 if t ** 3 > 0.008856 else (t - 16 / 116) / 7.787

    x = D65_WHITE[0] * inv(fx)
    y = D65_WHITE[1] * inv(fy)
    z = D65_WHITE[2] * inv(fz)

    rl = x * 3.2404542 + y * -1.5371385 + z * -0.4985314
    gl = x * -0.9692660 + y * 1.8760108 + z * 0.0415560
    bl = x * 0.0556434 + y * -0.2040259 + z * 1.0572252
    return (clamp(_linear_to_srgb(rl) * 255),
            clamp(_linear_to_srgb(gl) * 255),
            clamp(_linear_to_srgb(bl) * 255))


def mix_rgb(c1, c2, t):
    """按比例 t 混合，0 得 c1，1 得 c2。直接对通道插值。"""
    return (clamp(c1[0] + (c2[0] - c1[0]) * t),
            clamp(c1[1] + (c2[1] - c1[1]) * t),
            clamp(c1[2] + (c2[2] - c1[2]) * t))


def mix_lab(c1, c2, t):
    """在 CIELAB 空间插值。"""
    l1, a1, b1 = rgb_to_lab(*c1)
    l2, a2, b2 = rgb_to_lab(*c2)
    return lab_to_rgb(l1 + (l2 - l1) * t,
                      a1 + (a2 - a1) * t,
                      b1 + (b2 - b1) * t)


def make_gradient(c1, c2, steps, space="lab"):
    """
    生成从 c1 到 c2 的渐变，步数包含两端。
    steps=5 返回 5 个颜色：[c1, 中间3个, c2]。
    space 取 "lab"（感知均匀，默认）或 "rgb"（直接通道插值）。
    """
    steps = max(2, int(steps))
    mix = mix_lab if space.lower() == "lab" else mix_rgb
    return [mix(c1, c2, i / (steps - 1)) for i in range(steps)]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        samples = ["#2E5A88", "2E5A88", "#ABC", "rgb(46,90,136)", "red", "bad"]
        for s in samples:
            print("%-16s → %s" % (s, parse_color(s)))
        print("")
        src = (46, 90, 136)
        print("LAB 往返：%s → %s → %s"
              % (src, rgb_to_lab(*src), lab_to_rgb(*rgb_to_lab(*src))))
        print("")
        print("渐变 蓝→黄，5 步（LAB）：")
        for c in make_gradient((46, 90, 136), (240, 228, 66), 5):
            print("   %s  %s" % (rgb_to_hex(*c), c))
        print("渐变 蓝→黄，5 步（RGB，作对照）：")
        for c in make_gradient((46, 90, 136), (240, 228, 66), 5, "rgb"):
            print("   %s  %s" % (rgb_to_hex(*c), c))
    else:
        print(parse_color(sys.argv[1]))
