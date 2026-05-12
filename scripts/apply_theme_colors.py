"""仅修改 theme1.xml 的 12 个色槽（阶段 B 的换色操作）。

前提：PPT 已被 rebuild_theme.py 主题色化（所有可控色 srgbClr 已改为 schemeClr）。
之后改 theme 12 色就会全局生效。

输入 new_colors.json：
{
  "dk1":"RRGGBB","lt1":"RRGGBB","dk2":"RRGGBB","lt2":"RRGGBB",
  "accent1":"RRGGBB",...,"accent6":"RRGGBB",
  "hlink":"RRGGBB","folHlink":"RRGGBB"
}
（值可以是 {"hex": "RRGGBB"} 形式，也可直接是字符串）

WCAG 校验：dk1 vs lt1 对比度 < 4.5:1 时拒绝（除非 --force）。

CLI:
    python apply_theme_colors.py <pptx> --colors new_colors.json --out final.pptx [--force]
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
from color_utils import normalize_hex, wcag_contrast  # noqa: E402

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A = "{%s}" % NS_A

THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2",
               "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink"]


def _normalize_input(colors_in: dict) -> dict[str, str]:
    out = {}
    for slot in THEME_SLOTS:
        v = colors_in.get(slot)
        if v is None:
            raise ValueError(f"new_colors 缺少 {slot}")
        if isinstance(v, dict):
            v = v.get("hex")
        if not isinstance(v, str):
            raise ValueError(f"{slot} 值非法: {v!r}")
        out[slot] = normalize_hex(v)
    return out


def _build_clr_scheme(colors: dict[str, str], original_name: str) -> etree._Element:
    nsmap = {"a": NS_A}
    clr = etree.Element(A + "clrScheme", nsmap=nsmap)
    clr.set("name", original_name)

    def _add(slot: str, hex_val: str, sys_val: str | None = None) -> None:
        slot_el = etree.SubElement(clr, A + slot)
        if sys_val:
            sys_el = etree.SubElement(slot_el, A + "sysClr")
            sys_el.set("val", sys_val)
            sys_el.set("lastClr", hex_val)
        else:
            srgb = etree.SubElement(slot_el, A + "srgbClr")
            srgb.set("val", hex_val)

    for slot in THEME_SLOTS:
        if slot == "dk1":
            _add(slot, colors[slot], "windowText")
        elif slot == "lt1":
            _add(slot, colors[slot], "window")
        else:
            _add(slot, colors[slot])
    return clr


def _patch_theme_xml(theme_xml: bytes, colors: dict[str, str]) -> bytes:
    root = etree.fromstring(theme_xml)
    theme_elements = root.find(A + "themeElements")
    if theme_elements is None:
        raise RuntimeError("theme1.xml 缺少 themeElements")
    old = theme_elements.find(A + "clrScheme")
    name = old.get("name", "Custom") if old is not None else "Custom"
    new = _build_clr_scheme(colors, name)
    if old is not None:
        old.addprevious(new)
        theme_elements.remove(old)
    else:
        theme_elements.insert(0, new)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def apply(pptx_in: Path, colors_in: dict, pptx_out: Path, force: bool = False) -> dict:
    colors = _normalize_input(colors_in)

    # WCAG 校验
    contrast = wcag_contrast(colors["dk1"], colors["lt1"])
    if contrast < 4.5 and not force:
        raise ValueError(
            f"WCAG 对比度不足: dk1={colors['dk1']} vs lt1={colors['lt1']} "
            f"contrast={contrast:.2f} < 4.5。--force 可强制写入"
        )

    pptx_out.parent.mkdir(parents=True, exist_ok=True)
    if pptx_out.exists():
        pptx_out.unlink()

    with zipfile.ZipFile(pptx_in, "r") as zin:
        with zipfile.ZipFile(pptx_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == "ppt/theme/theme1.xml":
                    data = _patch_theme_xml(data, colors)
                zout.writestr(name, data)

    return {
        "ok": True,
        "out": str(pptx_out),
        "wcag_dk1_lt1_contrast": round(contrast, 2),
        "applied_colors": colors,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--colors", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--force", action="store_true",
                    help="忽略 WCAG 对比度不足的校验")
    args = ap.parse_args()

    if not args.pptx.exists():
        print(f"ERROR: pptx not found: {args.pptx}", file=sys.stderr)
        return 1
    if not args.colors.exists():
        print(f"ERROR: colors json not found: {args.colors}", file=sys.stderr)
        return 1

    data = json.loads(args.colors.read_text(encoding="utf-8"))
    # 兼容两种结构：{"theme_colors": {...}} 或 直接 {...}
    if "theme_colors" in data:
        data = data["theme_colors"]

    try:
        result = apply(args.pptx, data, args.out, force=args.force)
    except ValueError as e:
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
