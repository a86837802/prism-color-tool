# -*- coding: utf-8 -*-
"""
dialog_probe_v1.py —— Prism 颜色对话框控件树探测器（只读）

在 Prism 弹出颜色对话框时，抓取该窗口的类名、标题、所有者链、
全部子控件（ID / 类名 / 文本 / 坐标），并按 Windows SDK 的标准控件 ID
逐一试探，判定它是否就是系统通用颜色对话框。

本脚本只读取窗口信息，不发送任何消息、不修改任何控件、不写入 Prism。
"""

import argparse
import ctypes
import json
import sys
import time
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

import win32api
import win32con
import win32gui
import win32process


# ============================================================================
# 【使用说明】
#
# 目的
#     回答三个问题，它们决定后续实现能否成立：
#       1. Prism 10.6 的"更多颜色"弹的是不是 Windows 标准对话框（类名 #32770）
#       2. 如果是，R/G/B 三个输入框的控件 ID 是不是 dlgs.h 的标准值
#       3. 从 PPT 双击打开（OLE 就地激活）时，所有者链长什么样
#
# 环境
#     D:\\Conda\\envs\\dock\\python.exe   （已含 pywin32）
#
# 用法
#     默认：等待模式，最多等 180 秒
#       python dialog_probe_v1.py
#     立即快照：不等，直接扫描当前已打开的窗口然后退出
#       python dialog_probe_v1.py --now
#     自定义等待时长
#       python dialog_probe_v1.py --timeout 300
#
# 操作步骤
#     1. 运行脚本（在终端里）
#     2. 切到 Prism，打开任意一张图
#     3. Format Graph → 点某个数据集的颜色下拉 → 选 More Colors
#     4. 等脚本提示"已捕获"，回到终端看结果
#
# 输出
#     控制台摘要 + 输出/dialog_probe_v1_<时间戳>.md + .json
#
# 安全
#     全程只调用查询类 Win32 API（EnumWindows / GetWindowText / GetClassName
#     / GetDlgItem / GetWindowRect），不发送 WM_* 消息，不写任何控件。
# ============================================================================


# ============ 变量区（命令行参数会覆盖这里的默认值） ============
BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "输出"

POLL_INTERVAL = 0.3          # 轮询间隔（秒）
DEFAULT_TIMEOUT = 180        # 默认等待时长（秒）
SETTLE_SECONDS = 3.0         # 捕获首个候选后继续观察的时间，用于收齐同批窗口

# 候选标题。2026-11-13 实测：Prism 10.6 英文版的实际标题是 "Choose Color"。
# 中文版可能为"颜色"，一并列入。判定用子串匹配，因此 "Color" 也能覆盖
# "Choose Color"，但显式写出更利于阅读配置。
TITLE_CANDIDATES = ("Choose Color", "颜色", "颜色和透明度",
                    "Color", "Colors", "More Colors", "更多颜色")

# 标准颜色对话框的控件 ID，取自 Windows SDK 的 dlgs.h
STD_CONTROL_IDS = {
    700: "COLOR_HUESCROLL",
    701: "COLOR_SATSCROLL",
    702: "COLOR_LUMSCROLL",
    703: "COLOR_HUE",
    704: "COLOR_SAT",
    705: "COLOR_LUM",
    706: "COLOR_RED",
    707: "COLOR_GREEN",
    708: "COLOR_BLUE",
    709: "COLOR_CURRENT",
    710: "COLOR_RAINBOW",
    711: "COLOR_SAVE",
    712: "COLOR_ADD",
    713: "COLOR_SOLID",
    714: "COLOR_TUNE",
    715: "COLOR_SCHEMES",
    716: "COLOR_ELEMENT",
    717: "COLOR_SAMPLES",
    718: "COLOR_PALETTE",
    719: "COLOR_MIX",
    720: "COLOR_SELECT",
    1: "IDOK",
    2: "IDCANCEL",
}

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


# ============ 进程与窗口信息 ============

def get_process_name(pid):
    """取进程可执行文件名。失败返回空串。"""
    # 方式一：GetModuleFileNameEx（需要 PROCESS_VM_READ）
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
            False, pid)
        try:
            return Path(win32process.GetModuleFileNameEx(handle, 0)).name
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        pass
    # 方式二：QueryFullProcessImageNameW（权限要求更低）
    try:
        handle = win32api.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = wintypes.DWORD(1024)
            if ctypes.windll.kernel32.QueryFullProcessImageNameW(
                    handle, 0, buf, ctypes.byref(size)):
                return Path(buf.value).name
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        pass
    return ""


def get_window_basic(hwnd):
    """取窗口的基础信息。任何一项取不到都不抛异常。"""
    info = {"hwnd": hwnd, "hwnd_hex": "%08X" % hwnd}
    try:
        info["class"] = win32gui.GetClassName(hwnd)
    except Exception:
        info["class"] = ""
    try:
        info["title"] = win32gui.GetWindowText(hwnd)
    except Exception:
        info["title"] = ""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        info["pid"] = pid
        info["process"] = get_process_name(pid)
    except Exception:
        info["pid"] = 0
        info["process"] = ""
    try:
        info["rect"] = win32gui.GetWindowRect(hwnd)
    except Exception:
        info["rect"] = (0, 0, 0, 0)
    info["visible"] = bool(win32gui.IsWindowVisible(hwnd))
    info["enabled"] = bool(win32gui.IsWindowEnabled(hwnd))
    return info


def owner_chain(hwnd, depth=6):
    """沿 GW_OWNER 向上追溯。用于识别 OLE 就地激活场景。"""
    chain = []
    cur = hwnd
    for _ in range(depth):
        try:
            cur = win32gui.GetWindow(cur, win32con.GW_OWNER)
        except Exception:
            break
        if not cur:
            break
        chain.append(get_window_basic(cur))
    return chain


def enum_top_windows():
    result = []

    def callback(hwnd, _):
        result.append(hwnd)
        return True

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return result


def dump_control_tree(hwnd):
    """枚举对话框的全部后代控件（EnumChildWindows 会递归）。"""
    items = []

    def callback(h, _):
        try:
            ident = win32gui.GetDlgCtrlID(h)
        except Exception:
            ident = 0
        try:
            cls = win32gui.GetClassName(h)
        except Exception:
            cls = ""
        try:
            text = win32gui.GetWindowText(h)
        except Exception:
            text = ""
        try:
            rect = win32gui.GetWindowRect(h)
        except Exception:
            rect = (0, 0, 0, 0)
        items.append({
            "hwnd_hex": "%08X" % h,
            "id": ident,
            "id_const": STD_CONTROL_IDS.get(ident, ""),
            "class": cls,
            "text": text[:120],
            "rect": rect,
            "size": (rect[2] - rect[0], rect[3] - rect[1]),
            "visible": bool(win32gui.IsWindowVisible(h)),
            "enabled": bool(win32gui.IsWindowEnabled(h)),
        })
        return True

    try:
        win32gui.EnumChildWindows(hwnd, callback, None)
    except Exception:
        pass
    return items


def probe_std_ids(hwnd):
    """按 dlgs.h 的标准 ID 逐个 GetDlgItem 试探。"""
    out = []
    for cid in sorted(STD_CONTROL_IDS):
        try:
            child = win32gui.GetDlgItem(hwnd, cid)
        except Exception:
            child = 0
        if child:
            try:
                cls = win32gui.GetClassName(child)
            except Exception:
                cls = ""
            try:
                text = win32gui.GetWindowText(child)
            except Exception:
                text = ""
            out.append({
                "id": cid,
                "const": STD_CONTROL_IDS[cid],
                "class": cls,
                "text": text[:60],
                "visible": bool(win32gui.IsWindowVisible(child)),
            })
    return out


# ============ 判定 ============

def is_likely_target(info, chain):
    """
    判定这个 #32770 窗口是否值得详细探测。
    系统上任何程序都会弹 #32770，不做过滤会在等待期间被无关窗口打断。
    返回 (是否重点, 理由)。
    """
    proc = (info.get("process") or "").lower()
    title = info.get("title") or ""
    if "prism" in proc:
        return True, "进程名含 prism"
    if any(t and t in title for t in TITLE_CANDIDATES):
        return True, "标题命中候选集"
    for w in chain or []:
        if "prism" in (w.get("process") or "").lower():
            return True, "所有者链含 prism"
    return False, ""


def judge_dialog(info):
    """
    给候选窗口下判定。
    返回 (等级, 结论文字)。等级：A 标准对话框可用 / B 需适配 / C 不适用
    """
    cls = info.get("class", "")
    title = info.get("title", "")
    tree = info.get("tree") or []
    std = {item["id"]: item for item in (info.get("std_ids") or [])}

    if cls != "#32770":
        return "C", ("窗口类名为 `%s`，不是标准对话框类 #32770。"
                     "若该窗口确实是 Prism 的颜色选择界面，说明它是自绘的，"
                     "需另设定位策略。" % (cls or "空"))

    has_rgb = all(cid in std and std[cid]["class"].lower() == "edit"
                  for cid in (706, 707, 708))
    has_add = 712 in std
    has_ok = 1 in std

    if has_rgb and has_add and has_ok:
        return "A", ("标准系统颜色对话框：类名 #32770，R/G/B 输入框位于 706/707/708，"
                     "Add to Custom Colors 位于 712，OK 位于 1。§3.2 与 §3.3 的假设成立。")

    missing = []
    if not has_rgb:
        missing.append("706/707/708 未全部对应到 Edit 控件")
    if not has_add:
        missing.append("712 不存在")
    if not has_ok:
        missing.append("1 (IDOK) 不存在")
    return "B", ("是 #32770 对话框，但标准控件 ID 不完全匹配：%s。"
                 "需按下方控件树重新映射，并把新 ID 写入配置。"
                 % "；".join(missing))


# ============ 报告 ============

def build_report(captured, elapsed, all_dialogs, others):
    lines = []
    lines.append("# Prism 颜色对话框探测报告 v1")
    lines.append("")
    lines.append("- 生成时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("- 等待时长：%.1f 秒" % elapsed)
    lines.append("- 候选对话框数量：%d" % len(captured))
    lines.append("- 本次探测未发送任何消息、未修改任何控件")
    lines.append("- Python %s" % sys.version.split()[0])
    lines.append("")

    if not captured:
        lines.append("## 未捕获到候选对话框")
        lines.append("")
        lines.append("可能原因与排查顺序：")
        lines.append("")
        lines.append("1. 等待期间没有在 Prism 里打开颜色对话框。重跑脚本，")
        lines.append("   在 180 秒内完成 Format Graph → 颜色下拉 → More Colors。")
        lines.append("2. 颜色对话框属于其他窗口类，未被判定条件纳入。")
        lines.append("   本次运行期间观察到的全部对话框类窗口如下：")
        lines.append("")
        if all_dialogs:
            lines.append("| 类名 | 标题 | 进程 | PID |")
            lines.append("|---|---|---|---|")
            for w in all_dialogs:
                lines.append("| `%s` | %s | %s | %d |"
                             % (w["class"], w["title"] or "（无标题）",
                                w["process"] or "?", w["pid"]))
        else:
            lines.append("（无）")
        lines.append("")
        return "\n".join(lines)

    for idx, item in enumerate(captured, 1):
        info = item["info"]
        grade, verdict = judge_dialog(info)
        lines.append("## 候选 %d：%s" % (idx, info["title"] or "（无标题）"))
        lines.append("")
        lines.append("| 项目 | 值 |")
        lines.append("|---|---|")
        lines.append("| 句柄 | `%s` |" % info["hwnd_hex"])
        lines.append("| 类名 | `%s` |" % info["class"])
        lines.append("| 标题 | %s |" % (info["title"] or "（无标题）"))
        lines.append("| 进程 | %s |" % (info["process"] or "（取不到）"))
        lines.append("| PID | %d |" % info["pid"])
        lines.append("| 矩形 | %s |" % (info["rect"],))
        lines.append("| 可见 / 可用 | %s / %s |" % (info["visible"], info["enabled"]))
        lines.append("| **判定** | **%s 级** |" % grade)
        lines.append("| 结论 | %s |" % verdict)
        lines.append("")

        chain = info.get("owner_chain") or []
        lines.append("所有者链（自下而上）：")
        lines.append("")
        if chain:
            lines.append("| 层级 | 类名 | 标题 | 进程 | PID |")
            lines.append("|---|---|---|---|---|")
            for i, w in enumerate(chain, 1):
                lines.append("| %d | `%s` | %s | %s | %d |"
                             % (i, w["class"], w["title"] or "（无）",
                                w["process"] or "?", w["pid"]))
        else:
            lines.append("无所有者（顶层窗口，或所有者链为空）")
        lines.append("")

        std = info.get("std_ids") or []
        lines.append("按 dlgs.h 标准 ID 的试探结果：")
        lines.append("")
        if std:
            lines.append("| ID | 常量名 | 类名 | 文本 | 可见 |")
            lines.append("|---|---|---|---|---|")
            for s in std:
                lines.append("| %d | `%s` | `%s` | %s | %s |"
                             % (s["id"], s["const"], s["class"],
                                s["text"] or "（空）", s["visible"]))
        else:
            lines.append("**一个标准 ID 都没有命中。** 该对话框的控件编号与 dlgs.h 完全不同。")
        lines.append("")

        tree = info.get("tree") or []
        lines.append("完整控件树（%d 个）：" % len(tree))
        lines.append("")
        lines.append("| 句柄 | ID | 常量名 | 类名 | 文本 | 尺寸 | 可见 |")
        lines.append("|---|---|---|---|---|---|---|")
        for c in tree:
            lines.append("| `%s` | %d | %s | `%s` | %s | %dx%d | %s |"
                         % (c["hwnd_hex"], c["id"],
                            ("`%s`" % c["id_const"]) if c["id_const"] else "",
                            c["class"], (c["text"] or "（空）")[:40],
                            c["size"][0], c["size"][1], c["visible"]))
        lines.append("")

    # 总结
    lines.append("## 总结")
    lines.append("")
    grades = [judge_dialog(i["info"])[0] for i in captured]
    if "A" in grades:
        lines.append("命中 A 级：Prism 弹的就是标准系统颜色对话框，控件 ID 与 dlgs.h 一致。")
        lines.append("§3.2 的识别条件与 §3.3 的控件定位可以直接使用，无需降级方案。")
    elif "B" in grades:
        lines.append("命中 B 级：是标准对话框类，但控件编号与 dlgs.h 不同。")
        lines.append("按上面控件树里 `Edit` 类控件的实际 ID 建立映射，写入配置文件即可。")
    else:
        lines.append("全部为 C 级：Prism 没有使用标准对话框。")
        lines.append("需要重新设计定位策略（按控件树结构启发式查找，或改用坐标定位）。")
    lines.append("")

    if others:
        lines.append("## 附：本次扫描中出现的其他对话框类窗口（已排除，非 Prism）")
        lines.append("")
        lines.append("| 类名 | 标题 | 进程 | PID |")
        lines.append("|---|---|---|---|")
        for w in others:
            lines.append("| `%s` | %s | %s | %d |"
                         % (w["class"], w["title"] or "（无标题）",
                            w["process"] or "?", w["pid"]))
        lines.append("")
    return "\n".join(lines)


# ============ 主流程 ============

def main():
    parser = argparse.ArgumentParser(
        description="探测 Prism 颜色对话框的控件结构（只读）")
    parser.add_argument("--now", action="store_true",
                        help="不等待，立即扫描当前已打开的窗口")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT,
                        help="等待秒数，默认 %d" % DEFAULT_TIMEOUT)
    parser.add_argument("-o", "--output", default=str(OUTPUT_DIR),
                        help="报告输出目录")
    args = parser.parse_args()

    out_dir = Path(args.output).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 62)
    print("Prism 颜色对话框探测器（只读）")
    print("=" * 62)
    if args.now:
        print("模式：立即快照")
    else:
        print("模式：等待捕获（最多 %.0f 秒）" % args.timeout)
        print("")
        print("请在 Prism 里打开一个颜色对话框：")
        print("  选一张图 → Format Graph → 某个数据集的颜色下拉 → More Colors")
        print("")
    print("扫描中……（按 Ctrl+C 可提前结束）")
    print("")

    start = time.time()
    captured = []
    others = []
    seen = {}
    all_dialogs = []
    settle_until = None

    try:
        while True:
            now = time.time()
            if not args.now and now - start > args.timeout:
                break
            for hwnd in enum_top_windows():
                if hwnd in seen:
                    continue
                info = get_window_basic(hwnd)
                seen[hwnd] = info
                is_dialog_class = (info["class"] == "#32770")
                is_prism = (info["process"] or "").lower() == "prism.exe"
                if not (is_dialog_class or is_prism):
                    continue
                all_dialogs.append(info)
                if not is_dialog_class:
                    continue
                info["owner_chain"] = owner_chain(hwnd)
                likely, reason = is_likely_target(info, info["owner_chain"])
                if not likely:
                    # 系统里任何程序都会弹 #32770，无关的只记基本信息，
                    # 不深入探测，也不触发"观察倒计时"
                    others.append(info)
                    continue
                # 重点候选：补齐全套信息
                info["match_reason"] = reason
                info["std_ids"] = probe_std_ids(hwnd)
                info["tree"] = dump_control_tree(hwnd)
                captured.append({"info": info})
                grade, verdict = judge_dialog(info)
                print("[捕获] 类名=%s 标题=%s 进程=%s（%s）" %
                      (info["class"], info["title"] or "（无）",
                       info["process"] or "?", reason))
                print("[判定] %s 级 —— %s" % (grade, verdict))
                print("[详情] 控件 %d 个，标准 ID 命中 %d 个"
                      % (len(info["tree"]), len(info["std_ids"])))
                print("")
                if settle_until is None and not args.now:
                    settle_until = time.time() + SETTLE_SECONDS
                    print("继续观察 %.0f 秒以收齐同批窗口……" % SETTLE_SECONDS)
                    print("")
            if args.now:
                break
            if settle_until and time.time() > settle_until:
                break
            time.sleep(POLL_INTERVAL)
    except KeyboardInterrupt:
        print("")
        print("已手动中断。")

    elapsed = time.time() - start
    report = build_report(captured, elapsed, all_dialogs, others)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    md_path = out_dir / ("Dialog探测报告_v1_%s.md" % stamp)
    json_path = out_dir / ("Dialog探测结果_v1_%s.json" % stamp)
    md_path.write_text(report, encoding="utf-8")
    json_path.write_text(
        json.dumps({"elapsed": elapsed, "captured": captured,
                    "observed_dialogs": all_dialogs},
                   ensure_ascii=False, indent=2, default=str),
        encoding="utf-8")

    print("=" * 62)
    print("扫描结束，用时 %.1f 秒，捕获候选对话框 %d 个" % (elapsed, len(captured)))
    print("报告：%s" % md_path)
    print("原始数据：%s" % json_path)
    if not captured:
        print("")
        print("没有捕获到与 Prism 相关的对话框。")
        print("请确认两件事：1) Prism 已经启动；2) 等待期间确实点开过颜色对话框。")
        print("若 Prism 没启动，可先加 --timeout 300 延长等待，再操作。")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
