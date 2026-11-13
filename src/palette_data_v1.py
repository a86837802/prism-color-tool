# -*- coding: utf-8 -*-
"""palette_data_v1.py —— 内置科研配色色板"""


# ============================================================================
# 【使用说明】
#
# 存放内置色板数据。每套色板是一个字典：
#   name   显示名称
#   note   用途说明（会显示在面板的提示行）
#   colors [(色名, #RRGGBB), ...]
#
# 单独查看：python palette_data_v1.py
# ============================================================================


# ============ 变量区 ============
OKABE_ITO = {
    "name": "Okabe-Ito",
    "note": "色盲友好，Nature/Science 系期刊推荐",
    "colors": [
        ("黑", "#000000"),
        ("橙", "#E69F00"),
        ("天蓝", "#56B4E9"),
        ("蓝绿", "#009E73"),
        ("黄", "#F0E442"),
        ("蓝", "#0072B2"),
        ("朱红", "#D55E00"),
        ("红紫", "#CC79A7"),
    ],
}

BREWER_DARK2 = {
    "name": "ColorBrewer Dark2",
    "note": "分类数据，八类别区分度高",
    "colors": [
        ("1", "#1B9E77"),
        ("2", "#D95F02"),
        ("3", "#7570B3"),
        ("4", "#E7298A"),
        ("5", "#66A61E"),
        ("6", "#E6AB02"),
        ("7", "#A6761D"),
        ("8", "#666666"),
    ],
}

GRAYS = {
    "name": "灰度梯度",
    "note": "黑白打印，六阶均匀",
    "colors": [
        ("0", "#000000"),
        ("20", "#333333"),
        ("40", "#666666"),
        ("60", "#999999"),
        ("80", "#CCCCCC"),
        ("100", "#FFFFFF"),
    ],
}

PALETTES = [OKABE_ITO, BREWER_DARK2, GRAYS]


def get_palette(name):
    for p in PALETTES:
        if p["name"] == name:
            return p
    return PALETTES[0]


def palette_names():
    return [p["name"] for p in PALETTES]


if __name__ == "__main__":
    for palette in PALETTES:
        print("%s —— %s" % (palette["name"], palette["note"]))
        for label, hex_value in palette["colors"]:
            print("   %-4s %s" % (label, hex_value))
        print("")
