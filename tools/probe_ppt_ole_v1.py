# -*- coding: utf-8 -*-
"""
probe_ppt_ole_v1.py —— PPTX 内嵌 OLE 对象结构探测器（全程只读）

探测 PowerPoint 文件中由 GraphPad Prism 粘贴生成的对象，
判定其存储方式，并给出"能否解析出 Prism 数据"的技术结论。
"""

import argparse
import io
import json
import posixpath
import re
import struct
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

try:
    import pythoncom
    import win32com.storagecon as storagecon
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False


# ============================================================================
# 【使用说明】
#
# 用途
#     Prism 的"粘贴为 Prism 对象"会把图表及其关联数据嵌入 PPT。本脚本打开
#     该 PPT，提取其中每一个内嵌对象，判断它属于以下哪一种：
#       1. 打包方式（存在 \x01Ole10Native 流）
#          → 流内是完整的原始文件字节，可原样还原为 .pzfx，能直接解析
#       2. OLE 服务器方式（存在 CONTENTS 等流，无 Ole10Native）
#          → 存的是 Prism 私有序列化格式，需逆向才能解析
#       3. 未发现内嵌对象
#          → 该图是粘贴为图片，不含数据
#
#     .pptx / .pptm 本质是 zip 包，路径约定：
#       ppt/slides/slideN.xml            幻灯片内容，内嵌对象以 <p:oleObj> 标记
#       ppt/slides/_rels/slideN.xml.rels 关系映射，r:id → ../embeddings/xxx.bin
#       ppt/embeddings/oleObjectN.bin    OLE 结构化存储（CFBF）容器
#       ppt/media/*                      OLE 对象旁边的预览图（EMF/PNG）
#
# 环境
#     D:\Conda\envs\dock\python.exe      （已含 pywin32，无需安装任何东西）
#     仅依赖标准库与 pywin32 的 pythoncom。
#
# 用法
#     python probe_ppt_ole_v1.py                      扫描 samples/ 下全部文件
#     python probe_ppt_ole_v1.py -i 某个.pptx         只查一个文件
#     python probe_ppt_ole_v1.py -i 某个目录          查一个目录
#     python probe_ppt_ole_v1.py -o 输出目录          指定报告输出位置
#     python probe_ppt_ole_v1.py --keep-temp          保留解出的中间文件（调试用）
#
# 安全说明
#     本脚本只做读取与解析：不解压到磁盘（唯一的例外是 OLE 容器，因为
#     StgOpenStorage 需要文件路径，会写到项目内的 .tmp/ 目录并在结束后删除）。
#     不修改、不移动、不覆盖任何原始 PPT。
# ============================================================================


# ============ 变量区（命令行参数会覆盖这里的默认值） ============
BASE_DIR = Path(__file__).resolve().parent.parent
SAMPLE_DIR = BASE_DIR / "samples"          # 默认扫描目录
OUTPUT_DIR = BASE_DIR / "输出"              # 默认报告输出目录
TEMP_DIR = BASE_DIR / ".tmp"               # OLE 容器临时落盘目录

CFBF_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # OLE 结构化存储文件头
SUPPORTED_SUFFIX = {".pptx", ".pptm", ".potx", ".potm"}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ============ 基础工具 ============

def sniff_bytes(data, limit=32):
    """根据文件头猜测数据类型，用于判断流内容是什么。"""
    if not data:
        return "空"
    head = data[:limit]
    if head[:8] == CFBF_MAGIC:
        return "OLE 结构化存储（CFBF）"
    if head[:4] == b"PK\x03\x04":
        return "ZIP 容器"
    if head[:2] == b"\x1f\x8b":
        return "gzip 压缩流"
    if head[:5] == b"%PDF-":
        return "PDF 文档"
    if head[:4] == b"\x89PNG":
        return "PNG 图片"
    if head[:2] == b"\xff\xd8":
        return "JPEG 图片"
    stripped = head.lstrip(b"\xef\xbb\xbf").lstrip()
    if stripped[:5] in (b"<?xml",):
        return "XML 文本"
    if stripped[:1] in (b"{", b"["):
        return "JSON 文本"
    if stripped[:1] == b"<":
        return "XML/HTML 文本"
    # 老版 Prism 的 .pzf / 新版 .pzfx 的私有头部无法凭魔数直接判定，
    # 这里只报告可打印字符比例，供人工比对
    printable = sum(1 for b in head if 32 <= b < 127)
    if printable / len(head) > 0.85:
        return "疑似纯文本（可打印字符占比 %.0f%%）" % (printable * 100.0 / len(head))
    return "未知二进制（无已知魔数）"


def hex_preview(data, n=64):
    """返回前 n 字节的十六进制字符串，便于人工比对格式。"""
    return " ".join("%02X" % b for b in data[:n])


def decode_mbcs(raw):
    """按 Windows 本地编码解码字符串（Ole10Native 里是 ANSI）。"""
    for enc in ("utf-8", "gbk", "mbcs"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("latin-1", errors="replace")


# ============ OLE 容器解析 ============

def read_stream(storage, name):
    """从 IStorage 里完整读出一个流。失败返回 None。"""
    try:
        # 注意：pywin32 要求第 2 个参数（reserved）必须是 None，传 0 会抛异常
        stream = storage.OpenStream(
            name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE
        )
    except Exception:
        return None
    chunks = []
    while True:
        try:
            chunk = stream.Read(65536)
        except Exception:
            break
        if not chunk:
            break
        chunks.append(bytes(chunk))
    stream = None          # pywin32 用引用计数释放，无 Release 方法
    return b"".join(chunks)


def dump_ole_container(bin_bytes, tag):
    """把一个 oleObjectN.bin 的字节流写成临时文件，用 StgOpenStorage 枚举内部流。"""
    if not HAS_PYWIN32:
        return {"error": "缺少 pywin32，无法打开 OLE 容器"}

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    tmp_path = TEMP_DIR / ("ole_probe_%s.bin" % tag)
    tmp_path.write_bytes(bin_bytes)

    result = {"container": sniff_bytes(bin_bytes), "streams": [], "storages": []}
    storage = None
    try:
        storage = pythoncom.StgOpenStorage(
            str(tmp_path), None,
            storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE,
        )
        for elem in storage.EnumElements():
            name = elem[0]
            data = read_stream(storage, name)
            if data is None:
                # OpenStream 失败说明是子存储（目录），不是流
                result["storages"].append(name)
                continue
            readable = name.replace("\x01", "^A").replace("\x05", "^E")
            result["streams"].append({
                "name": readable,
                "raw_name": name,
                "size": len(data),
                "sniff": sniff_bytes(data),
                "hex": hex_preview(data, 48),
            })
            if name.lstrip("\x01\x05") == "Ole10Native":
                result["ole10native_raw"] = data
            elif name.lstrip("\x01\x05") == "CompObj":
                result["compobj_raw"] = data
            elif name == "CONTENTS":
                result["contents_raw"] = data
    except Exception as exc:
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
    finally:
        storage = None      # 触发引用释放
        try:
            tmp_path.unlink()
        except OSError:
            pass

    # CONTENTS 流单独深挖：它是 Prism 10 存放项目本体的地方
    raw = result.pop("contents_raw", None)
    if raw is not None:
        result["contents"] = inspect_contents_stream(raw)
    return result


def inspect_contents_stream(data):
    """
    Prism 10 的 CONTENTS 流结构：4 字节小端长度字段 + ZIP 数据。
    ZIP 内包含 JSON 元数据、CSV 原始数据，以及 graphs/ 下的私有二进制。
    """
    out = {"size": len(data)}
    if len(data) <= 8:
        out["note"] = "流过短，无法判定"
        return out
    declared = struct.unpack("<I", data[:4])[0]
    out["declared_len"] = declared
    out["len_matches"] = (declared == len(data))
    body = data[4:]
    if body[:4] != b"PK\x03\x04":
        out["is_zip"] = False
        out["body_sniff"] = sniff_bytes(body)
        return out

    out["is_zip"] = True
    try:
        with zipfile.ZipFile(io.BytesIO(body)) as zf:
            entries = [{"name": n, "size": zf.getinfo(n).file_size}
                       for n in zf.namelist()]
            out["entries"] = entries
            names = {e["name"] for e in entries}
            out["graph_bins"] = [e for e in entries
                                 if e["name"].startswith("graphs/")
                                 and e["name"].endswith("data.bin")]
            out["csv_tables"] = [e["name"] for e in entries
                                 if e["name"].endswith("data.csv")]
            out["json_count"] = sum(1 for n in names if n.endswith(".json"))
            if "document.json" in names:
                try:
                    doc = json.loads(zf.read("document.json").decode("utf-8"))
                    out["document"] = {
                        "formatVersion": doc.get("formatVersion"),
                        "minPrismVersion": doc.get("minPrismVersion"),
                        "createdBy": (doc.get("createdBy") or {}).get("version"),
                        "modifiedBy": (doc.get("modifiedBy") or {}).get("version"),
                        "sheetCounts": {k: len(v) for k, v in
                                        (doc.get("sheets") or {}).items()},
                    }
                except Exception as exc:
                    out["document_error"] = str(exc)
    except Exception as exc:
        out["zip_error"] = "%s: %s" % (type(exc).__name__, exc)
    return out


def parse_ole10native(data):
    """
    解析 Ole10Native 流。结构（微软公开的打包格式）：
        0..5    固定 6 字节头
        6..     原始文件名（ANSI，\\0 结尾）
        ..      原始完整路径（ANSI，\\0 结尾）
        4 字节  未使用
        4 字节  临时路径长度（小端）
        ..      临时路径（ANSI）
        4 字节  内容长度（小端）
        ..      内容字节（即原始文件本体）
    """
    out = {}
    try:
        pos = 6
        end = data.index(b"\x00", pos)
        out["filename"] = decode_mbcs(data[pos:end])
        pos = end + 1
        end = data.index(b"\x00", pos)
        out["original_path"] = decode_mbcs(data[pos:end])
        pos = end + 1
        pos += 4
        temp_len = struct.unpack("<I", data[pos:pos + 4])[0]
        pos += 4
        out["temp_path"] = decode_mbcs(data[pos:pos + temp_len])
        pos += temp_len
        size = struct.unpack("<I", data[pos:pos + 4])[0]
        pos += 4
        content = data[pos:pos + size]
        out["content_size"] = size
        out["content_actual"] = len(content)
        out["content_sniff"] = sniff_bytes(content)
        out["content_hex"] = hex_preview(content, 48)
        out["content_truncated"] = len(content) < size

        # 若内容本身就是 zip（新版 Office 文件），列出其内部条目供比对
        if content[:4] == b"PK\x03\x04":
            tmp = TEMP_DIR / "_content_probe.zip"
            try:
                tmp.write_bytes(content)
                with zipfile.ZipFile(tmp) as zf:
                    out["zip_entries"] = zf.namelist()[:40]
            except Exception as exc:
                out["zip_error"] = str(exc)
            finally:
                try:
                    tmp.unlink()
                except OSError:
                    pass
    except Exception as exc:
        out["parse_error"] = "%s: %s" % (type(exc).__name__, exc)
    return out


def parse_compobj(data):
    """从 \\x01CompObj 流里提取可读字符串（主要是 ProgID 与显示名）。"""
    if not data:
        return []
    text = "".join(chr(b) if 32 <= b < 127 else "\n" for b in data)
    parts = [p.strip() for p in text.split("\n") if len(p.strip()) >= 3]
    return parts[:10]


# ============ PPTX 包结构解析 ============

def resolve_target(slide_path, target):
    """把 rels 里的相对路径 ../embeddings/x.bin 规范化成 zip 内绝对路径。"""
    base = posixpath.dirname(slide_path)
    return posixpath.normpath(posixpath.join(base, target))


def parse_slide_ole(slide_xml, rels_xml, slide_path):
    """从一张幻灯片的 XML 里找出全部 <p:oleObj>，并关联到 embeddings 文件。"""
    rid_map = {}
    if rels_xml:
        try:
            rroot = ET.fromstring(rels_xml)
            for rel in rroot:
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    rid_map[rid] = resolve_target(slide_path, target)
        except ET.ParseError:
            pass

    found = []
    seen_rid = set()
    if not slide_xml:
        return found
    try:
        root = ET.fromstring(slide_xml)
    except ET.ParseError:
        return found

    for elem in root.iter():
        if not elem.tag.endswith("}oleObj"):
            continue
        rid = elem.get("{%s}id" % REL_NS)
        # PowerPoint 会把同一个对象在 mc:Choice 与 mc:Fallback 里各写一份，
        # 表现为同一 r:id 出现两次。按 r:id 去重，避免重复计数。
        if rid and rid in seen_rid:
            continue
        if rid:
            seen_rid.add(rid)
        found.append({
            "slide": posixpath.basename(slide_path),
            "name": elem.get("name", ""),
            "prog_id": elem.get("progId", ""),
            "show_as_icon": elem.get("showAsIcon", ""),
            "r_id": rid or "",
            "embedding_path": rid_map.get(rid, ""),
        })
    return found


def inspect_package(pptx_path):
    """解析一个 pptx/pptm 包，返回结构化结果。"""
    info = {
        "file": str(pptx_path),
        "size_kb": round(pptx_path.stat().st_size / 1024.0, 1),
        "ok": False,
        "slides": 0,
        "media": [],
        "ole_objects": [],
        "other_embeddings": [],
        "notes": [],
    }
    try:
        with zipfile.ZipFile(pptx_path) as zf:
            names = zf.namelist()
            info["ok"] = True
            slides = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
            info["slides"] = len(slides)
            info["media"] = [
                {"name": n, "size_kb": round(zf.getinfo(n).file_size / 1024.0, 1)}
                for n in names if n.startswith("ppt/media/")
            ]
            embeddings = [n for n in names if n.startswith("ppt/embeddings/")]

            # 逐张幻灯片找 oleObj
            for slide in slides:
                rels_path = posixpath.join(
                    posixpath.dirname(slide), "_rels",
                    posixpath.basename(slide) + ".rels"
                )
                slide_xml = zf.read(slide) if slide in names else b""
                rels_xml = zf.read(rels_path) if rels_path in names else b""
                for obj in parse_slide_ole(slide_xml, rels_xml, slide):
                    info["ole_objects"].append(obj)

            # 与 oleObj 无关的其他 embeddings（例如纯附件）
            referenced = {o["embedding_path"] for o in info["ole_objects"] if o["embedding_path"]}
            info["other_embeddings"] = [n for n in embeddings if n not in referenced]

            # 逐个打开 OLE 容器
            for idx, obj in enumerate(info["ole_objects"], 1):
                path = obj["embedding_path"]
                if not path or path not in names:
                    obj["container_error"] = "关系中指向的 embeddings 文件不存在"
                    continue
                raw = zf.read(path)
                obj["bin_size_kb"] = round(len(raw) / 1024.0, 1)
                obj["bin_magic"] = sniff_bytes(raw)
                if raw[:8] != CFBF_MAGIC:
                    obj["container_error"] = "不是 CFBF 容器，可能是旧式打包"
                    continue
                dump = dump_ole_container(raw, "o%d" % idx)
                obj["streams"] = dump.get("streams", [])
                obj["sub_storages"] = dump.get("storages", [])
                if "error" in dump:
                    obj["container_error"] = dump["error"]
                if dump.get("compobj_raw"):
                    obj["compobj_strings"] = parse_compobj(dump["compobj_raw"])
                if dump.get("ole10native_raw"):
                    obj["ole10native"] = parse_ole10native(dump["ole10native_raw"])
                if dump.get("contents"):
                    obj["contents"] = dump["contents"]
    except zipfile.BadZipFile:
        info["notes"].append("不是有效的 OOXML 包（可能是 .ppt 旧格式，本脚本不支持）")
    except Exception as exc:
        info["notes"].append("%s: %s" % (type(exc).__name__, exc))
    return info


# ============ 结论判定 ============

def judge(obj):
    """根据流结构给出该对象的技术判定。"""
    if obj.get("container_error"):
        return "异常", obj["container_error"]
    if obj.get("ole10native"):
        o10 = obj["ole10native"]
        if o10.get("parse_error"):
            return "打包方式（解析出错）", o10["parse_error"]
        return "打包方式", (
            "内含完整原始文件 %s（%d 字节，类型：%s）"
            % (o10.get("filename", "?"), o10.get("content_actual", 0),
               o10.get("content_sniff", "?"))
        )
    contents = obj.get("contents") or {}
    if contents.get("is_zip"):
        return "OLE 服务器方式（内含 ZIP 项目容器）", (
            "CONTENTS 去掉 4 字节长度前缀后是 ZIP，共 %d 个条目："
            "%d 个图二进制、%d 个 CSV 数据表、%d 个 JSON"
            % (len(contents.get("entries", [])), len(contents.get("graph_bins", [])),
               len(contents.get("csv_tables", [])), contents.get("json_count", 0)))
    stream_names = [s["name"] for s in obj.get("streams", [])]
    if "CONTENTS" in stream_names:
        return "OLE 服务器方式", "存在 CONTENTS 流，但未识别出已知容器结构，需逆向"
    if stream_names:
        return "未知 OLE 结构", "内部流：%s" % "、".join(stream_names)
    return "空容器或未解析", "未发现任何流"


# ============ 报告生成 ============

def build_report(results, target_desc):
    lines = []
    lines.append("# PPT 内嵌 Prism 对象探测报告 v1")
    lines.append("")
    lines.append("- 生成时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("- 扫描目标：`%s`" % target_desc)
    lines.append("- 解释器：Python %s（pywin32 %s）"
                 % (sys.version.split()[0], "可用" if HAS_PYWIN32 else "缺失"))
    lines.append("- 本次探测未修改任何原始文件")
    lines.append("")

    total_obj = sum(len(r["ole_objects"]) for r in results)
    lines.append("## 一、文件总览")
    lines.append("")
    lines.append("| 文件 | 大小(KB) | 幻灯片 | OLE对象 | 媒体文件 | 状态 |")
    lines.append("|---|---|---|---|---|---|")
    for r in results:
        status = "正常" if r["ok"] else "；".join(r["notes"]) or "解析失败"
        lines.append("| %s | %s | %d | %d | %d | %s |" % (
            Path(r["file"]).name, r["size_kb"], r["slides"],
            len(r["ole_objects"]), len(r["media"]), status))
    lines.append("")
    lines.append("共扫描 %d 个文件，发现 %d 个内嵌对象。" % (len(results), total_obj))
    lines.append("")

    if total_obj == 0:
        lines.append("## 二、结论")
        lines.append("")
        lines.append("未发现任何 OLE 内嵌对象。这意味着这些图是**粘贴为图片**，")
        lines.append("不含 Prism 数据，无法双击进入 Prism 编辑。")
        lines.append("")
        lines.append("若你确认双击能打开 Prism，请检查：")
        lines.append("1. 粘贴时是否用了「选择性粘贴 → GraphPad Prism 项目对象」；")
        lines.append("2. 是否误扫了导出后的 PPT（另存为 PDF 或图片后对象会丢失）。")
        lines.append("")
        if results and results[0]["media"]:
            lines.append("该文件中现有媒体文件（供比对粘贴格式）：")
            lines.append("")
            for m in results[0]["media"][:20]:
                lines.append("- `%s`（%.1f KB）" % (m["name"], m["size_kb"]))
            lines.append("")

    idx = 0
    for r in results:
        for obj in r["ole_objects"]:
            idx += 1
            verdict, detail = judge(obj)
            lines.append("## 对象 %d：%s / %s" % (idx, Path(r["file"]).name, obj["slide"]))
            lines.append("")
            lines.append("| 项目 | 值 |")
            lines.append("|---|---|")
            lines.append("| 显示名 | `%s` |" % (obj.get("name") or "（无）"))
            lines.append("| ProgID | `%s` |" % (obj.get("prog_id") or "（无）"))
            lines.append("| 关系项 | `%s` |" % (obj.get("r_id") or "（无）"))
            lines.append("| 嵌入路径 | `%s` |" % (obj.get("embedding_path") or "（无）"))
            lines.append("| 容器大小 | %s KB |" % obj.get("bin_size_kb", "?"))
            lines.append("| 容器类型 | %s |" % obj.get("bin_magic", "?"))
            lines.append("| **判定** | **%s** |" % verdict)
            lines.append("| 说明 | %s |" % detail)
            lines.append("")

            if obj.get("compobj_strings"):
                lines.append("CompObj 流内可读字符串：%s" %
                             " / ".join("`%s`" % s for s in obj["compobj_strings"]))
                lines.append("")

            if obj.get("streams"):
                lines.append("内部流清单：")
                lines.append("")
                lines.append("| 流名 | 大小(字节) | 内容嗅探 | 前 48 字节 |")
                lines.append("|---|---|---|---|")
                for s in obj["streams"]:
                    lines.append("| `%s` | %d | %s | `%s` |"
                                 % (s["name"], s["size"], s["sniff"], s["hex"]))
                lines.append("")

            if obj.get("sub_storages"):
                lines.append("子存储：%s" %
                             "、".join("`%s`" % s for s in obj["sub_storages"]))
                lines.append("")

            # CONTENTS 流深挖结论
            c = obj.get("contents")
            if c and c.get("is_zip"):
                lines.append("**CONTENTS 流已识别为 ZIP 项目容器**")
                lines.append("")
                lines.append("- 前 4 字节长度字段：%s（与实际长度%s）"
                             % (c.get("declared_len"),
                                "一致" if c.get("len_matches") else "不一致"))
                lines.append("- 条目总数：%d（JSON %d 个，CSV 数据表 %d 个）"
                             % (len(c.get("entries", [])), c.get("json_count", 0),
                                len(c.get("csv_tables", []))))
                doc = c.get("document") or {}
                if doc:
                    lines.append("- formatVersion：`%s`，minPrismVersion：`%s`"
                                 % (doc.get("formatVersion"), doc.get("minPrismVersion")))
                    lines.append("- 创建版本：%s，最后修改版本：%s"
                                 % (doc.get("createdBy"), doc.get("modifiedBy")))
                    counts = doc.get("sheetCounts") or {}
                    lines.append("- 工作表数量：%s" % "、".join(
                        "%s %d" % (k, v) for k, v in counts.items()))
                lines.append("")
                if c.get("graph_bins"):
                    lines.append("图形私有二进制（绘图属性应存于此）：")
                    lines.append("")
                    for g in c["graph_bins"]:
                        lines.append("- `%s`（%d 字节）" % (g["name"], g["size"]))
                    lines.append("")
            elif c:
                lines.append("CONTENTS 流未识别为 ZIP：%s" % c.get("body_sniff", ""))
                lines.append("")

            o10 = obj.get("ole10native")
            if o10 and not o10.get("parse_error"):
                lines.append("Ole10Native 解析结果：")
                lines.append("")
                lines.append("| 字段 | 值 |")
                lines.append("|---|---|")
                lines.append("| 原始文件名 | `%s` |" % o10.get("filename", ""))
                lines.append("| 原始路径 | `%s` |" % o10.get("original_path", ""))
                lines.append("| 内容声明大小 | %s 字节 |" % o10.get("content_size", "?"))
                lines.append("| 内容实际读取 | %s 字节 |" % o10.get("content_actual", "?"))
                lines.append("| 内容类型 | %s |" % o10.get("content_sniff", "?"))
                lines.append("| 内容前 48 字节 | `%s` |" % o10.get("content_hex", ""))
                lines.append("")
                if o10.get("zip_entries"):
                    lines.append("内容为 ZIP，内部条目（前 40 个）：")
                    lines.append("")
                    for name in o10["zip_entries"]:
                        lines.append("- `%s`" % name)
                    lines.append("")

    # 总体结论
    lines.append("## 三、总体结论")
    lines.append("")
    verdicts = [judge(o)[0] for r in results for o in r["ole_objects"]]
    if not verdicts:
        lines.append("无对象可供判断。")
    elif all(v == "打包方式" for v in verdicts):
        lines.append("全部对象均为**打包方式**。这意味着可以从 PPT 中完整还原出")
        lines.append("原始 Prism 文件，解析这条路在技术上成立。")
    elif any(v.startswith("OLE 服务器方式") for v in verdicts):
        if any("ZIP" in v for v in verdicts):
            lines.append("存在 OLE 服务器方式的嵌入对象，且其 CONTENTS 流是 **ZIP 容器**。")
            lines.append("")
            lines.append("可读取：项目元数据、原始数据表（CSV）、分析参数与结果（JSON）、")
            lines.append("        工作表标题与轴标题。")
            lines.append("")
            lines.append("不可直读：绘图属性（含颜色、线型、字体），存放于 `graphs/` 下的")
            lines.append("          `data.bin`，魔数为 `PCFFGRA4` 的私有二进制，无字段名。")
            lines.append("")
            lines.append("另需注意：PPT 中显示的图形是 `OlePres000` 等流内的呈现缓存，")
            lines.append("修改数据后若不同步重建缓存，显示不会改变。")
        else:
            lines.append("存在 **OLE 服务器方式**的对象。其内部是 Prism 私有序列化格式，")
            lines.append("不含标准文件结构。若要解析，只能逆向其二进制布局；")
            lines.append("这在工程上可行但成本高，且写回极易破坏文件。")
    else:
        lines.append("对象结构异常，需人工检查上面的流清单。")
    lines.append("")
    return "\n".join(lines)


# ============ 主流程 ============

def collect_targets(target: Path):
    """收齐待扫描的 pptx 文件。"""
    if target.is_file():
        return [target]
    if target.is_dir():
        files = []
        for suffix in SUPPORTED_SUFFIX:
            files.extend(target.rglob("*" + suffix))
        return sorted(set(files))
    return []


def main():
    parser = argparse.ArgumentParser(
        description="探测 PPTX 内嵌 OLE 对象的结构（只读）")
    parser.add_argument("-i", "--input", default=str(SAMPLE_DIR),
                        help="待探测的 pptx 文件或目录，默认 samples/")
    parser.add_argument("-o", "--output", default=str(OUTPUT_DIR),
                        help="报告输出目录，默认 输出/")
    parser.add_argument("--keep-temp", action="store_true",
                        help="保留 .tmp 下的中间文件")
    args = parser.parse_args()

    target = Path(args.input).expanduser().resolve()
    out_dir = Path(args.output).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    files = collect_targets(target)
    if not files:
        print("未找到待探测的文件。")
        print("扫描位置：%s" % target)
        print("支持的扩展名：%s" % "、".join(sorted(SUPPORTED_SUFFIX)))
        print("请把 PPT 放进 samples/ 目录，或用 -i 指定路径。")
        return 1

    if not HAS_PYWIN32:
        print("警告：当前解释器没有 pywin32，无法打开 OLE 容器。")
        print("请改用 D:\\Conda\\envs\\dock\\python.exe 运行本脚本。")

    print("待探测文件 %d 个：" % len(files))
    for f in files:
        print("  - %s" % f)
    print("")

    results = []
    for f in files:
        print("正在解析：%s" % f.name)
        results.append(inspect_package(f))

    report = build_report(results, str(target))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    md_path = out_dir / ("PPT内嵌对象探测报告_v1_%s.md" % stamp)
    json_path = out_dir / ("PPT内嵌对象探测结果_v1_%s.json" % stamp)
    md_path.write_text(report, encoding="utf-8")
    json_path.write_text(
        json.dumps({"target": str(target), "results": results},
                   ensure_ascii=False, indent=2, default=str),
        encoding="utf-8")

    # 控制台摘要
    total = sum(len(r["ole_objects"]) for r in results)
    print("")
    print("=" * 60)
    for r in results:
        print("%s：幻灯片 %d 张，OLE 对象 %d 个，媒体 %d 个"
              % (Path(r["file"]).name, r["slides"],
                 len(r["ole_objects"]), len(r["media"])))
        for obj in r["ole_objects"]:
            verdict, detail = judge(obj)
            print("  [%s] ProgID=%s → %s" %
                  (obj["slide"], obj.get("prog_id") or "?", verdict))
            print("      %s" % detail)
        for note in r["notes"]:
            print("  注意：%s" % note)
    print("=" * 60)
    print("共 %d 个内嵌对象" % total)
    print("报告：%s" % md_path)
    print("原始数据：%s" % json_path)

    if not args.keep_temp:
        for leftover in TEMP_DIR.glob("*"):
            try:
                leftover.unlink()
            except OSError:
                pass
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
