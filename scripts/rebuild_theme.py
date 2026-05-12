"""阶段 A 核心：把 PPT 主题色化。

输入 decision.json（来自 Agent 的 OOXML × 视觉双向印证）：
{
  "theme_colors": {
    "dk1": {"hex":"...","source":"ooxml_confirmed"},
    "lt1": {"hex":"...","source":"ooxml_confirmed"},
    "dk2": {"hex":"..."}, "lt2": {"hex":"..."},
    "accent1..6": {"hex":"...","source":"vision_corrected","ooxml_origin":"..."},
    "hlink": {"hex":"..."}, "folHlink": {"hex":"..."}
  }
}

做两件事：
1. 重写 ppt/theme/theme1.xml 的 <a:clrScheme>，按 OOXML 规范：
   - dk1 / lt1 用 <a:sysClr val="windowText|window" lastClr="..."/>
   - 其余用 <a:srgbClr val="..."/>
2. 扫描 slides/layouts/masters 的所有 <a:srgbClr>：
   - 跳过 blipFill 内
   - 计算与 12 色的 CIEDE2000：
     * < 10 → 替换为 <a:schemeClr val="<slot>"/>
     * 10~25 → 记录到 unmapped_warning（保留原色）
     * > 25 → 保留
   - 双值机制：匹配源用 ooxml_origin（如有），赋值已写入 theme

输出 themed.pptx + 内嵌 docProps/custom.xml 审计信息。

CLI:
    python rebuild_theme.py <pptx> --decision decision.json --out themed.pptx
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from color_utils import hex_delta_e, normalize_hex  # noqa: E402

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]

THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2",
               "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink"]

DEFAULT_DELTA_E_MAP_THRESHOLD = 10.0
DEFAULT_DELTA_E_WARN_THRESHOLD = 25.0


def _build_clr_scheme(theme_colors: dict, original_name: str = "Custom") -> etree._Element:
    """按 OOXML 规范构建 <a:clrScheme> 元素。"""
    nsmap = {"a": NS["a"]}
    clr = etree.Element(A + "clrScheme", nsmap=nsmap)
    clr.set("name", original_name)

    def _add_slot(slot: str, hex_val: str, sys_val: str | None = None) -> None:
        slot_el = etree.SubElement(clr, A + slot)
        if sys_val is not None:
            sys_el = etree.SubElement(slot_el, A + "sysClr")
            sys_el.set("val", sys_val)
            sys_el.set("lastClr", hex_val.upper())
        else:
            srgb = etree.SubElement(slot_el, A + "srgbClr")
            srgb.set("val", hex_val.upper())

    for slot in THEME_SLOTS:
        info = theme_colors.get(slot)
        if not info:
            raise ValueError(f"theme_colors 缺少 {slot}")
        hex_val = normalize_hex(info["hex"])
        if slot == "dk1":
            _add_slot(slot, hex_val, sys_val="windowText")
        elif slot == "lt1":
            _add_slot(slot, hex_val, sys_val="window")
        else:
            _add_slot(slot, hex_val)
    return clr


def _replace_theme_clr_scheme(theme_xml: bytes, theme_colors: dict) -> bytes:
    root = etree.fromstring(theme_xml)
    # 找原有 clrScheme（在 themeElements 内）
    nsmap = {"a": NS["a"]}
    theme_elements = root.find(A + "themeElements")
    if theme_elements is None:
        raise RuntimeError("theme1.xml 中找不到 themeElements")
    old_scheme = theme_elements.find(A + "clrScheme")
    original_name = old_scheme.get("name", "Custom") if old_scheme is not None else "Custom"

    new_scheme = _build_clr_scheme(theme_colors, original_name)

    if old_scheme is not None:
        # 保持位置：插入到 old 的位置
        old_scheme.addprevious(new_scheme)
        theme_elements.remove(old_scheme)
    else:
        theme_elements.insert(0, new_scheme)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _is_in_blipfill(node: etree._Element) -> bool:
    cur = node
    while cur is not None:
        tag = cur.tag
        if tag == f"{A}blipFill" or tag == f"{P}blipFill":
            return True
        cur = cur.getparent()
    return False


def _build_match_table(theme_colors: dict) -> list[tuple[str, str, str]]:
    """构建 (匹配源 hex, 目标槽位, 赋值 hex) 列表。

    匹配源优先用 ooxml_origin（视觉修正过的色），兜底用 hex 自己。
    赋值用 theme 里写入的 hex（视觉修正后）。
    """
    out = []
    for slot, info in theme_colors.items():
        target_hex = normalize_hex(info["hex"])
        match_src = info.get("ooxml_origin") or target_hex
        match_src = normalize_hex(match_src)
        out.append((match_src, slot, target_hex))
    return out


def _remap_srgb_in_xml(
    xml_bytes: bytes,
    match_table: list[tuple[str, str, str]],
    explicit_map: dict[str, str],
    map_threshold: float,
    warn_threshold: float,
) -> tuple[bytes, list, list]:
    """扫描 srgbClr 并重映射，返回 (new_xml, mapped_records, unmapped_warnings)。

    explicit_map: 直接 hex -> slot 强制映射，优先级最高。
    """
    root = etree.fromstring(xml_bytes)
    mapped: list[dict] = []
    warnings: list[dict] = []

    for srgb in root.iter(f"{A}srgbClr"):
        if _is_in_blipfill(srgb):
            continue
        val = srgb.get("val")
        if not val:
            continue
        val_norm = normalize_hex(val)

        forced = explicit_map.get(val_norm)
        if forced:
            scheme = etree.Element(f"{A}schemeClr", nsmap={"a": NS["a"]})
            scheme.set("val", forced)
            for child in list(srgb):
                scheme.append(child)
            srgb.addprevious(scheme)
            srgb.getparent().remove(srgb)
            mapped.append({"original": val_norm, "slot": forced, "delta_e": 0.0,
                           "via": "slot_mapping"})
            continue

        best_slot, best_de = None, float("inf")
        for match_src, slot, _target in match_table:
            try:
                de = hex_delta_e(val_norm, match_src)
            except ValueError:
                continue
            if de < best_de:
                best_de, best_slot = de, slot

        if best_slot is None:
            continue

        if best_de < map_threshold:
            scheme = etree.Element(f"{A}schemeClr", nsmap={"a": NS["a"]})
            scheme.set("val", best_slot)
            for child in list(srgb):
                scheme.append(child)
            srgb.addprevious(scheme)
            srgb.getparent().remove(srgb)
            mapped.append({"original": val_norm, "slot": best_slot,
                           "delta_e": round(best_de, 2), "via": "delta_e"})
        elif best_de < warn_threshold:
            warnings.append({"original": val_norm, "nearest_slot": best_slot,
                             "delta_e": round(best_de, 2)})
        # 否则保留

    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return new_xml, mapped, warnings


def rebuild(
    pptx_in: Path, decision: dict, pptx_out: Path,
    *, slot_mapping: dict[str, str] | None = None,
    map_threshold: float = DEFAULT_DELTA_E_MAP_THRESHOLD,
    warn_threshold: float = DEFAULT_DELTA_E_WARN_THRESHOLD,
) -> dict:
    theme_colors = decision["theme_colors"]
    # 标准化所有 hex
    for slot in THEME_SLOTS:
        if slot not in theme_colors:
            raise ValueError(f"decision.theme_colors 缺少 {slot}")
        info = theme_colors[slot]
        info["hex"] = normalize_hex(info["hex"])
        if "ooxml_origin" in info and info["ooxml_origin"]:
            info["ooxml_origin"] = normalize_hex(info["ooxml_origin"])

    match_table = _build_match_table(theme_colors)
    explicit_map: dict[str, str] = {}
    if slot_mapping:
        for hx, slot in slot_mapping.items():
            if slot not in THEME_SLOTS:
                raise ValueError(f"slot_mapping 含未知槽位: {slot}")
            explicit_map[normalize_hex(hx)] = slot

    pptx_out.parent.mkdir(parents=True, exist_ok=True)
    if pptx_out.exists():
        pptx_out.unlink()

    all_mapped: list[dict] = []
    all_warnings: list[dict] = []

    with zipfile.ZipFile(pptx_in, "r") as zin:
        names = zin.namelist()
        with zipfile.ZipFile(pptx_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name == "ppt/theme/theme1.xml":
                    data = _replace_theme_clr_scheme(data, theme_colors)
                elif (
                    name.startswith("ppt/slides/slide") and name.endswith(".xml")
                    or name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml")
                    or name.startswith("ppt/slideMasters/slideMaster") and name.endswith(".xml")
                ):
                    try:
                        data, mapped, warnings = _remap_srgb_in_xml(
                            data, match_table, explicit_map,
                            map_threshold, warn_threshold,
                        )
                        for m in mapped:
                            m["file"] = name
                        for w in warnings:
                            w["file"] = name
                        all_mapped.extend(mapped)
                        all_warnings.extend(warnings)
                    except etree.XMLSyntaxError:
                        pass
                zout.writestr(name, data)

            audit = {
                "theme_colors": theme_colors,
                "slot_mapping_used": explicit_map,
                "thresholds": {"map": map_threshold, "warn": warn_threshold},
                "remap_count": len(all_mapped),
                "remap_via_slot_mapping": sum(1 for m in all_mapped if m.get("via") == "slot_mapping"),
                "remap_via_delta_e": sum(1 for m in all_mapped if m.get("via") == "delta_e"),
                "unmapped_warning_count": len(all_warnings),
                "unmapped_warnings": all_warnings[:50],
            }
            zout.writestr(
                "docProps/ppt-clone-audit.json",
                json.dumps(audit, ensure_ascii=False, indent=2),
            )

    return {
        "ok": True,
        "out": str(pptx_out),
        "remap_count": len(all_mapped),
        "remap_via_slot_mapping": sum(1 for m in all_mapped if m.get("via") == "slot_mapping"),
        "remap_via_delta_e": sum(1 for m in all_mapped if m.get("via") == "delta_e"),
        "warning_count": len(all_warnings),
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--decision", required=True, type=Path,
                    help="Agent 双向印证产物 JSON")
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--slot-mapping", type=Path, default=None,
                    help="可选：显式 hex->slot 映射 JSON，优先级最高")
    ap.add_argument("--map-threshold", type=float, default=DEFAULT_DELTA_E_MAP_THRESHOLD,
                    help=f"ΔE 替换阈值，默认 {DEFAULT_DELTA_E_MAP_THRESHOLD}")
    ap.add_argument("--warn-threshold", type=float, default=DEFAULT_DELTA_E_WARN_THRESHOLD,
                    help=f"ΔE 警告阈值，默认 {DEFAULT_DELTA_E_WARN_THRESHOLD}")
    args = ap.parse_args()

    if not args.pptx.exists():
        print(f"ERROR: pptx not found: {args.pptx}", file=sys.stderr)
        return 1
    if not args.decision.exists():
        print(f"ERROR: decision json not found: {args.decision}", file=sys.stderr)
        return 1

    decision = json.loads(args.decision.read_text(encoding="utf-8"))
    slot_mapping = None
    if args.slot_mapping:
        if not args.slot_mapping.exists():
            print(f"ERROR: slot_mapping not found: {args.slot_mapping}", file=sys.stderr)
            return 1
        slot_mapping = json.loads(args.slot_mapping.read_text(encoding="utf-8"))

    result = rebuild(
        args.pptx, decision, args.out,
        slot_mapping=slot_mapping,
        map_threshold=args.map_threshold,
        warn_threshold=args.warn_threshold,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
