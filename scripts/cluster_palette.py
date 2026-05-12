"""把 collect_ooxml_colors 输出的可控色板做 ΔE 聚类，并给出槽位建议。

- 输入：collect_ooxml_colors.py 的 palette.json
- 输出：color_clusters.json
  {
    "schema_version": 1,
    "threshold_de": 8.0,
    "clusters": [
      {
        "id": "C1",
        "representative_hex": "404450",
        "members": [{"hex": "404450", "frequency": 146}, ...],
        "total_frequency": 240,
        "carriers": ["shape_fill","text_color"],
        "in_chart_count": 0, "in_tbl_count": 27,
        "is_dominant": true,
        "lab_L": 28.4,
        "is_grayscale": false,
        "suggested_slot": "dk1"
      }, ...
    ],
    "slot_suggestion_rationale": "..."
  }

启发式槽位建议（仅作为 Agent 参考，最终决策仍由 Agent 给 slot_mapping）：
  - 最深的非彩色 -> dk1
  - 最浅的非彩色 -> lt1
  - 第二深 -> dk2
  - 第二浅 -> lt2
  - 频次最高的若干彩色 -> accent1..6（按色相分散排序）
  - hlink/folHlink 不主动建议（多数模板未硬编码）

CLI:
    python cluster_palette.py palette.json --out color_clusters.json [--threshold 8.0]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from color_utils import hex_delta_e, hex_to_lab, hex_to_rgb, normalize_hex  # noqa: E402

THEME_SLOTS_ORDERED = ["dk1", "lt1", "dk2", "lt2",
                       "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                       "hlink", "folHlink"]


def _is_grayscale(hex_color: str, sat_thresh: float = 8.0) -> bool:
    r, g, b = hex_to_rgb(hex_color)
    return (max(r, g, b) - min(r, g, b)) <= sat_thresh


def _hue(hex_color: str) -> float:
    """返回 0..360 色相，灰阶返回 -1。"""
    r, g, b = (c / 255.0 for c in hex_to_rgb(hex_color))
    mx, mn = max(r, g, b), min(r, g, b)
    d = mx - mn
    if d < 0.03:
        return -1.0
    if mx == r:
        h = 60 * (((g - b) / d) % 6)
    elif mx == g:
        h = 60 * (((b - r) / d) + 2)
    else:
        h = 60 * (((r - g) / d) + 4)
    return h % 360


def cluster(palette: dict, threshold: float) -> list[dict]:
    colors = sorted(palette.get("colors", []),
                    key=lambda c: c.get("frequency", 0), reverse=True)
    clusters: list[dict] = []
    for c in colors:
        hx = normalize_hex(c["hex"])
        placed = False
        for cl in clusters:
            try:
                de = hex_delta_e(hx, cl["representative_hex"])
            except ValueError:
                continue
            if de <= threshold:
                cl["members"].append({
                    "hex": hx, "frequency": c.get("frequency", 0),
                    "carriers": c.get("carriers", []),
                })
                cl["total_frequency"] += c.get("frequency", 0)
                # 重新选 representative：取频次最高的成员
                best_member = max(cl["members"], key=lambda m: m["frequency"])
                cl["representative_hex"] = best_member["hex"]
                cl["carriers"] = sorted(set(
                    sum([m.get("carriers", []) for m in cl["members"]], [])
                ))
                cl["in_chart_count"] = cl.get("in_chart_count", 0) + c.get("in_chart_count", 0)
                cl["in_tbl_count"] = cl.get("in_tbl_count", 0) + c.get("in_tbl_count", 0)
                cl["in_dgm_count"] = cl.get("in_dgm_count", 0) + c.get("in_dgm_count", 0)
                placed = True
                break
        if not placed:
            clusters.append({
                "representative_hex": hx,
                "members": [{
                    "hex": hx, "frequency": c.get("frequency", 0),
                    "carriers": c.get("carriers", []),
                }],
                "total_frequency": c.get("frequency", 0),
                "carriers": list(c.get("carriers", [])),
                "in_chart_count": c.get("in_chart_count", 0),
                "in_tbl_count": c.get("in_tbl_count", 0),
                "in_dgm_count": c.get("in_dgm_count", 0),
            })

    # 排 id 与基础属性
    for i, cl in enumerate(clusters, start=1):
        cl["id"] = f"C{i}"
        L, _, _ = hex_to_lab(cl["representative_hex"])
        cl["lab_L"] = round(L, 2)
        cl["is_grayscale"] = _is_grayscale(cl["representative_hex"])
        cl["hue"] = round(_hue(cl["representative_hex"]), 2)
    return clusters


def suggest_slots(clusters: list[dict]) -> tuple[list[dict], str]:
    """给每个 cluster 建议 suggested_slot。返回 (clusters, rationale)。"""
    grays = sorted([c for c in clusters if c["is_grayscale"]], key=lambda c: c["lab_L"])
    chroma = [c for c in clusters if not c["is_grayscale"]]
    chroma.sort(key=lambda c: c["total_frequency"], reverse=True)

    # 重置 suggested_slot
    for c in clusters:
        c["suggested_slot"] = None

    if grays:
        grays[0]["suggested_slot"] = "dk1"
        grays[-1]["suggested_slot"] = "lt1"
        if len(grays) >= 3:
            grays[1]["suggested_slot"] = "dk2"
            grays[-2]["suggested_slot"] = "lt2"
        elif len(grays) == 2:
            pass  # dk2/lt2 留空
    accent_idx = 1
    for cl in chroma:
        if accent_idx > 6:
            break
        cl["suggested_slot"] = f"accent{accent_idx}"
        accent_idx += 1

    rationale_bits = [
        f"grayscale={len(grays)} (assigned dk1/lt1/dk2/lt2 by Lab L*)",
        f"chromatic={len(chroma)} (accent1..accent6 by frequency)",
        "dk2/lt2 may be left blank when grayscale clusters < 3; Agent should derive.",
        "hlink/folHlink left to Agent: most templates don't hardcode them.",
    ]
    return clusters, " | ".join(rationale_bits)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("palette", type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--threshold", type=float, default=8.0,
                    help="ΔE 合并阈值，默认 8.0（CIEDE2000）")
    args = ap.parse_args()

    if not args.palette.exists():
        print(f"ERROR: palette not found: {args.palette}", file=sys.stderr)
        return 1

    palette = json.loads(args.palette.read_text(encoding="utf-8"))
    clusters = cluster(palette, args.threshold)
    clusters, rationale = suggest_slots(clusters)

    payload = {
        "schema_version": 1,
        "threshold_de": args.threshold,
        "clusters": clusters,
        "slot_suggestion_rationale": rationale,
        "summary": {
            "input_unique_colors": len(palette.get("colors", [])),
            "cluster_count": len(clusters),
            "grayscale_clusters": sum(1 for c in clusters if c["is_grayscale"]),
        },
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "out": str(args.out),
                      **payload["summary"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
