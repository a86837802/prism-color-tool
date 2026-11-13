# -*- coding: utf-8 -*-
"""_quit_test_v1.py —— 验证退出链路：设置标志后进程能否真正结束"""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import prism_color_tool_v1 as t


def main():
    tool = t.Tool()
    tool.tray = t.tr.TrayIcon(tool.request_quit, "退出测试")
    ready = tool.tray.start()
    print("托盘启动:", ready, flush=True)

    def trigger():
        print("→ 设置退出标志", flush=True)
        tool._quit_requested = True

    tool.root.after(3000, trigger)
    tool.root.after(100, tool.poll)

    start = time.time()
    tool.root.mainloop()
    print("mainloop 已退出，耗时 %.1f 秒" % (time.time() - start), flush=True)

    # 主动停托盘线程，观察是否会卡住
    print("→ 停止托盘线程", flush=True)
    if tool.tray:
        tool.tray.stop()
    print("→ 测试结束，若此处无输出说明进程卡住", flush=True)
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
