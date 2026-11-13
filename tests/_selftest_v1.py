# -*- coding: utf-8 -*-
"""
_selftest_v1.py —— 探测脚本的自测夹具生成器

造两个最小可解析的 .pptx，用来验证 probe_ppt_ole_v1.py 的两条判定分支：
  1. fixture_package.pptx   打包方式：oleObject1.bin 内含 \x01Ole10Native 流
  2. fixture_ole_server.pptx OLE 服务器方式：oleObject1.bin 内只有 CONTENTS 流

用法：
    D:\\Conda\\envs\\dock\\python.exe tests\\_selftest_v1.py
然后：
    D:\\Conda\\envs\\dock\\python.exe probe\\probe_ppt_ole_v1.py -i tests\\_fixture

说明：
    夹具只写在 tests/_fixture/ 下，不触碰 samples/ 与任何真实文件。
    CFBF 容器用 pythoncom 真实创建，与 Office 生成的结构一致。
"""

import io
import struct
import sys
import zipfile
from pathlib import Path

import pythoncom
import win32com.storagecon as storagecon


# ============ 变量区 ============
BASE_DIR = Path(__file__).resolve().parent.parent
FIXTURE_DIR = Path(__file__).resolve().parent / "_fixture"

SLIDE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:oleObj spid="_x0000_s1026" name="Object 1" progId="Prism.Graph"
              r:id="rId2" imgW="3600000" imgH="2400000">
      <p:embed/>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="Object 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="2400000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:oleObj>
  </p:spTree></p:cSld>
</p:sld>
"""

RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/>
</Relationships>
"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>
"""


# ============ 构造工具 ============

def make_fake_pzfx():
    """造一个假的项目内容。用 zip 封装，模拟"内容本身是容器文件"的情形。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("project.json", '{"name":"fake","graphs":[{"colors":["#2E5A88"]}]}')
        zf.writestr("data/sheet1.csv", "x,y\n1,2\n3,4\n")
    return buf.getvalue()


def make_ole10native(filename, orig_path, content):
    """按打包方式构造 \\x01Ole10Native 流。"""
    buf = bytearray()
    buf += b"\x02\x00\x00\x00\x02\x00"                 # 6 字节固定头
    buf += filename.encode("mbcs") + b"\x00"
    buf += orig_path.encode("mbcs") + b"\x00"
    buf += b"\x00\x00\x00\x00"                          # 4 字节未使用
    temp = r"C:\Temp\ole_tmp.tmp"
    buf += struct.pack("<I", len(temp)) + temp.encode("mbcs")
    buf += struct.pack("<I", len(content)) + content
    buf += struct.pack("<I", 0)                         # 尾部 unicode 段留空
    return bytes(buf)


def build_cfbf(path, streams):
    """用 pythoncom 真实创建一个 CFBF 容器并写入若干流。"""
    storage = pythoncom.StgCreateDocfile(
        str(path),
        storagecon.STGM_READWRITE | storagecon.STGM_CREATE | storagecon.STGM_SHARE_EXCLUSIVE,
        0,
    )
    for name, data in streams.items():
        # IStorage::CreateStream 的签名是 (name, grfMode, reserved1, reserved2)
        stream = storage.CreateStream(
            name,
            storagecon.STGM_CREATE | storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE,
            0, 0)
        stream.Write(data)
        stream.Commit(0)
        stream = None          # pywin32 用引用计数释放，无 Release 方法
    storage.Commit(0)
    storage = None


def pack_pptx(pptx_path, ole_bin_path, media_bytes):
    """把幻灯片、关系、嵌入对象打成一个最小 pptx。"""
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", RELS_XML)
        zf.writestr("ppt/embeddings/oleObject1.bin", ole_bin_path.read_bytes())
        zf.writestr("ppt/media/image1.emf", media_bytes)


# ============ 主流程 ============

def main():
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    payload = make_fake_pzfx()

    # 夹具一：打包方式
    bin1 = FIXTURE_DIR / "_ole_package.bin"
    build_cfbf(bin1, {
        "\x01Ole10Native": make_ole10native("test.pzfx", r"D:\data\test.pzfx", payload),
        "\x01CompObj": b"\x01\x00\xfe\xff" + b"GraphPad Prism Project\x00",
    })
    out1 = FIXTURE_DIR / "fixture_package.pptx"
    pack_pptx(out1, bin1, b"\x01\x00\x00\x00" + b"\x00" * 32)
    bin1.unlink()

    # 夹具二：OLE 服务器方式（只有 CONTENTS）
    bin2 = FIXTURE_DIR / "_ole_server.bin"
    build_cfbf(bin2, {
        "CONTENTS": b"\xde\xad\xbe\xef" + bytes(range(64)),
        "\x01CompObj": b"\x01\x00\xfe\xff" + b"Prism.Graph\x00",
    })
    out2 = FIXTURE_DIR / "fixture_ole_server.pptx"
    pack_pptx(out2, bin2, b"\x01\x00\x00\x00" + b"\x00" * 32)
    bin2.unlink()

    print("夹具已生成：")
    for p in (out1, out2):
        print("  %s（%.1f KB）" % (p, p.stat().st_size / 1024.0))
    print("")
    print("下一步运行探测脚本验证：")
    print('  python probe\\probe_ppt_ole_v1.py -i tests\\_fixture')
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    sys.exit(main())
