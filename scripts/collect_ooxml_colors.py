"""从 OOXML 直接采集所有"可控色"。

可控色 = 可以通过修改 OOXML 主题色或硬编码色实现换色的颜色，包括：
- shape 填充 srgbClr (在 spPr/solidFill 内)
- shape 线条 srgbClr (在 ln/solidFill 内)
- 文字 srgbClr (在 rPr/solidFill 内)
- 渐变色标 srgbClr (在 gradFill/gsLst/gs 内)
- 主题色引用 schemeClr (单独记录槽位名)

明确排除：
- blipFill 内任何颜色（图片像素，不可被 OOXML 主题色控制）
- chart/table/SmartArt 子树暂时也采，但单独标记 in_chart/in_tbl/in_dgm

同时采集每页的图片 bbox（p:pic 和 spPr/blipFill 形状），用于 render 阶段叠加 mask。

CLI:
    python collect_ooxml_colors.py <pptx> --out palette.json
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}

A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
DGM = "{%s}" % NS["dgm"]
C = "{%s}" % NS["c"]


def _classify_carrier(node: etree._Element) -> str:
    """根据 srgbClr/schemeClr 的祖先链推断它担当什么角色。"""
    in_chart = False
    in_tbl = False
    in_dgm = False

    cur = node
    nearest_fill_kind = None
    nearest_container = None

    while cur is not None:
        tag = cur.tag
        if tag == f"{C}chart":
            in_chart = True
        elif tag == f"{A}tbl":
            in_tbl = True
        elif tag == f"{DGM}graphic" or tag.startswith(DGM):
            in_dgm = True
        elif nearest_fill_kind is None:
            if tag == f"{A}solidFill":
                nearest_fill_kind = "solidFill"
            elif tag == f"{A}gradFill":
                nearest_fill_kind = "gradFill"

        if nearest_container is None:
            if tag == f"{A}rPr":
                nearest_container = "text"
            elif tag == f"{A}ln":
                nearest_container = "line"
            elif tag == f"{P}spPr" or tag == f"{A}spPr":
                # spPr 同时出现在 ln 之外的位置 -> 是形状填充
                if nearest_container is None:
                    nearest_container = "shape_fill"
        cur = cur.getparent()

    if nearest_container == "text":
        carrier = "text_color"
    elif nearest_container == "line":
        carrier = "line_color"
    elif nearest_container == "shape_fill":
        carrier = "shape_fill"
    else:
        carrier = "other"

    if nearest_fill_kind == "gradFill":
        carrier = "gradient_" + carrier

    flags = []
    if in_chart:
        flags.append("in_chart")
    if in_tbl:
        flags.append("in_tbl")
    if in_dgm:
        flags.append("in_dgm")

    return carrier + ("|" + ",".join(flags) if flags else "")


def _is_in_blipfill(node: etree._Element) -> bool:
    cur = node
    while cur is not None:
        tag = cur.tag
        if tag == f"{A}blipFill" or tag == f"{P}blipFill":
            return True
        cur = cur.getparent()
    return False


def _has_background_image(xml_bytes: bytes) -> bool:
    """检查 cSld/bg/bgPr/blipFill 是否存在（slide/layout/master 通用）。"""
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return False
    for bg in root.iter(f"{P}bg"):
        bgPr = bg.find(f"{P}bgPr")
        if bgPr is None:
            continue
        if bgPr.find(f"{A}blipFill") is not None or bgPr.find(f"{P}blipFill") is not None:
            return True
    return False


def _read_slide_size(zf: zipfile.ZipFile) -> tuple[int, int]:
    try:
        root = etree.fromstring(zf.read("ppt/presentation.xml"))
        sldSz = root.find(f"{P}sldSz")
        if sldSz is not None:
            return int(sldSz.get("cx", "9144000")), int(sldSz.get("cy", "6858000"))
    except Exception:
        pass
    return 9144000, 6858000  # 默认 16:9 早期常见值


def _read_rel_target(zf: zipfile.ZipFile, rels_path: str, type_suffix: str) -> str | None:
    try:
        root = etree.fromstring(zf.read(rels_path))
        for rel in root:
            t = rel.get("Type", "")
            if t.endswith(type_suffix):
                return rel.get("Target")
    except KeyError:
        return None
    except Exception:
        return None
    return None


def _walk_xml(xml_bytes: bytes, source_label: str, color_records: list, image_regions: list, slide_idx: int | None):
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return

    # 1) srgbClr 采集
    for node in root.iter(f"{A}srgbClr"):
        if _is_in_blipfill(node):
            continue
        val = node.get("val")
        if not val:
            continue
        carrier = _classify_carrier(node)
        color_records.append({
            "hex": val.upper(),
            "carrier": carrier,
            "source_xml": source_label,
            "kind": "srgbClr",
        })

    # 2) schemeClr 引用记录（不参与色板提取，但记录哪些色已主题化）
    for node in root.iter(f"{A}schemeClr"):
        if _is_in_blipfill(node):
            continue
        val = node.get("val")
        if not val:
            continue
        carrier = _classify_carrier(node)
        color_records.append({
            "hex": None,
            "scheme_slot": val,
            "carrier": carrier,
            "source_xml": source_label,
            "kind": "schemeClr",
        })

    # 3) 图片 bbox 采集（仅 slide 级，layout/master 暂不采）
    if slide_idx is not None:
        for pic in root.iter(f"{P}pic"):
            spPr = pic.find(f"{P}spPr")
            if spPr is None:
                spPr = pic.find(f"{A}spPr")
            if spPr is None:
                continue
            xfrm = spPr.find(f"{A}xfrm")
            if xfrm is None:
                continue
            off = xfrm.find(f"{A}off")
            ext = xfrm.find(f"{A}ext")
            if off is None or ext is None:
                continue
            try:
                box = {
                    "x": int(off.get("x", "0")),
                    "y": int(off.get("y", "0")),
                    "w": int(ext.get("cx", "0")),
                    "h": int(ext.get("cy", "0")),
                    "image": pic.find(f".//{P}cNvPr").get("descr", "") if pic.find(f".//{P}cNvPr") is not None else "",
                    "kind": "p:pic",
                }
                image_regions.append({"slide": slide_idx, "box": box})
            except (TypeError, ValueError):
                continue

        # blipFill 在 sp 上的（图片填充形状）
        for sp in root.iter(f"{P}sp"):
            sp_spPr = sp.find(f"{P}spPr")
            if sp_spPr is None:
                continue
            blip = sp_spPr.find(f"{A}blipFill")
            if blip is None:
                blip = sp_spPr.find(f"{P}blipFill")
            if blip is None:
                continue
            xfrm = sp_spPr.find(f"{A}xfrm")
            if xfrm is None:
                continue
            off = xfrm.find(f"{A}off")
            ext = xfrm.find(f"{A}ext")
            if off is None or ext is None:
                continue
            try:
                box = {
                    "x": int(off.get("x", "0")),
                    "y": int(off.get("y", "0")),
                    "w": int(ext.get("cx", "0")),
                    "h": int(ext.get("cy", "0")),
                    "image": "shape_blipFill",
                    "kind": "sp:blipFill",
                }
                image_regions.append({"slide": slide_idx, "box": box})
            except (TypeError, ValueError):
                continue


def collect(pptx_path: Path) -> dict:
    color_records: list = []
    image_regions: list = []
    chart_table_smartart: list = []

    with zipfile.ZipFile(pptx_path, "r") as z:
        names = z.namelist()
        slide_w, slide_h = _read_slide_size(z)

        # 解析 slide -> layout -> master 的引用链
        layouts_with_bg: dict[str, bool] = {}
        masters_with_bg: dict[str, bool] = {}
        for name in names:
            if name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml"):
                layouts_with_bg[name] = _has_background_image(z.read(name))
            elif name.startswith("ppt/slideMasters/slideMaster") and name.endswith(".xml"):
                masters_with_bg[name] = _has_background_image(z.read(name))

        def _slide_inherited_bg(slide_name: str) -> list[dict]:
            """看 slide 关联的 layout/master 是否有背景图，返回需 mask 的全页 box。"""
            inherited = []
            base = slide_name.split("/")[-1]
            rels = f"ppt/slides/_rels/{base}.rels"
            layout_target = _read_rel_target(z, rels, "/slideLayout")
            if layout_target:
                layout_path = ("ppt/slides/" + layout_target).replace("/../", "/")
                # 规范化路径
                parts = []
                for seg in layout_path.split("/"):
                    if seg == "..":
                        if parts: parts.pop()
                    elif seg and seg != ".":
                        parts.append(seg)
                layout_path = "/".join(parts)
                if layouts_with_bg.get(layout_path):
                    inherited.append({"x": 0, "y": 0, "w": slide_w, "h": slide_h,
                                       "image": "layout_background", "kind": "layout:bg"})
                # layout -> master
                layout_base = layout_path.split("/")[-1]
                layout_rels = f"ppt/slideLayouts/_rels/{layout_base}.rels"
                master_target = _read_rel_target(z, layout_rels, "/slideMaster")
                if master_target:
                    master_path = ("ppt/slideLayouts/" + master_target)
                    parts = []
                    for seg in master_path.split("/"):
                        if seg == "..":
                            if parts: parts.pop()
                        elif seg and seg != ".":
                            parts.append(seg)
                    master_path = "/".join(parts)
                    if masters_with_bg.get(master_path):
                        inherited.append({"x": 0, "y": 0, "w": slide_w, "h": slide_h,
                                           "image": "master_background", "kind": "master:bg"})
            return inherited

        # theme
        for name in sorted(n for n in names if n.startswith("ppt/theme/") and n.endswith(".xml")):
            _walk_xml(z.read(name), name, color_records, image_regions, slide_idx=None)

        # slide masters
        for name in sorted(n for n in names if n.startswith("ppt/slideMasters/") and n.endswith(".xml")):
            _walk_xml(z.read(name), name, color_records, image_regions, slide_idx=None)

        # slide layouts
        for name in sorted(n for n in names if n.startswith("ppt/slideLayouts/") and n.endswith(".xml")):
            _walk_xml(z.read(name), name, color_records, image_regions, slide_idx=None)

        # slides
        slide_files = sorted(
            (n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
            key=lambda s: int(s.rsplit("slide", 1)[1].split(".")[0]),
        )
        for name in slide_files:
            idx = int(name.rsplit("slide", 1)[1].split(".")[0])
            xml_bytes = z.read(name)
            _walk_xml(xml_bytes, name, color_records, image_regions, slide_idx=idx)

            # slide 自身的背景图
            if _has_background_image(xml_bytes):
                image_regions.append({
                    "slide": idx,
                    "box": {"x": 0, "y": 0, "w": slide_w, "h": slide_h,
                            "image": "slide_background", "kind": "slide:bg"},
                })
            # 继承自 layout/master 的背景图
            for box in _slide_inherited_bg(name):
                image_regions.append({"slide": idx, "box": box})

            # chart/table/smartart 区域标记
            try:
                root = etree.fromstring(xml_bytes)
                for tag, kind in [(f"{C}chart", "chart"), (f"{A}tbl", "table"), (f"{DGM}graphic", "smartart")]:
                    for node in root.iter(tag):
                        # 找 graphicFrame 的 xfrm
                        gf = node
                        while gf is not None and gf.tag != f"{P}graphicFrame":
                            gf = gf.getparent()
                        if gf is None:
                            continue
                        xfrm = gf.find(f"{P}xfrm")
                        if xfrm is None:
                            continue
                        off = xfrm.find(f"{A}off")
                        ext = xfrm.find(f"{A}ext")
                        if off is None or ext is None:
                            continue
                        chart_table_smartart.append({
                            "slide": idx,
                            "kind": kind,
                            "bbox": {
                                "x": int(off.get("x", "0")),
                                "y": int(off.get("y", "0")),
                                "w": int(ext.get("cx", "0")),
                                "h": int(ext.get("cy", "0")),
                            },
                        })
            except etree.XMLSyntaxError:
                pass

    # 聚合 srgbClr：按 hex 统计频次和承载类型
    srgb_records = [r for r in color_records if r["kind"] == "srgbClr"]
    by_hex: dict[str, dict] = defaultdict(lambda: {"frequency": 0, "carriers": [], "in_chart_count": 0, "in_tbl_count": 0, "in_dgm_count": 0})
    for r in srgb_records:
        e = by_hex[r["hex"]]
        e["frequency"] += 1
        e["carriers"].append(r["carrier"].split("|")[0])
        if "in_chart" in r["carrier"]:
            e["in_chart_count"] += 1
        if "in_tbl" in r["carrier"]:
            e["in_tbl_count"] += 1
        if "in_dgm" in r["carrier"]:
            e["in_dgm_count"] += 1

    # 排序：频次降序
    colors = []
    for hx, data in by_hex.items():
        carriers_dedup = sorted(set(data["carriers"]))
        colors.append({
            "hex": hx,
            "frequency": data["frequency"],
            "carriers": carriers_dedup,
            "in_chart_count": data["in_chart_count"],
            "in_tbl_count": data["in_tbl_count"],
            "in_dgm_count": data["in_dgm_count"],
            "schemeClr_refs": [],  # 留空，schemeClr 的引用单独记录
        })
    colors.sort(key=lambda c: c["frequency"], reverse=True)
    if colors:
        max_f = colors[0]["frequency"]
        for c in colors:
            c["is_dominant"] = c["frequency"] >= max(3, max_f * 0.3)

    # schemeClr 槽位使用统计
    scheme_usage: dict[str, int] = defaultdict(int)
    for r in color_records:
        if r["kind"] == "schemeClr":
            scheme_usage[r["scheme_slot"]] += 1

    # 图片 bbox 按 slide 聚合
    by_slide: dict[int, list] = defaultdict(list)
    for entry in image_regions:
        by_slide[entry["slide"]].append(entry["box"])
    image_regions_per_slide = [
        {"slide": idx, "boxes": boxes}
        for idx, boxes in sorted(by_slide.items())
    ]

    return {
        "colors": colors,
        "scheme_slot_usage": dict(scheme_usage),
        "slide_size_emu": {"w": slide_w, "h": slide_h},
        "image_regions_per_slide": image_regions_per_slide,
        "chart_table_smartart_zones": chart_table_smartart,
        "summary": {
            "total_srgb_color_uses": sum(c["frequency"] for c in colors),
            "unique_srgb_colors": len(colors),
            "image_regions_count": sum(len(r["boxes"]) for r in image_regions_per_slide),
        },
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--out", required=True, type=Path)
    args = ap.parse_args()

    if not args.pptx.exists():
        print(f"ERROR: pptx not found: {args.pptx}", file=sys.stderr)
        return 1

    result = collect(args.pptx)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "out": str(args.out), **result["summary"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
