# -*- coding: utf-8 -*-
"""抽样验收：比对 原模板 vs _clone.pptx 的字号/字数一致性。

对每个 content_shape，从原文件读出每段的 dominant font size + char count，
然后从 _clone 对应 shape_id 读出同样信息。报告：
- paragraph_count_mismatch 统计
- font_size_delta_distribution：|clone - src| 绝对值分布
- char_ratio_distribution：clone_chars / src_chars 分布
- 前 5 条大字号段对照样例

用法：
  python ppt-clone-skill/scripts/verify_font_and_chars.py
"""
from __future__ import annotations

import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
SCRIPTS = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import parse_template_story as pts  # noqa: E402


PAIRS = [
    ("20607102 简约机械设备产品介绍宣传推广.pptx",
     "20607102 简约机械设备产品介绍宣传推广_clone.pptx"),
    ("50129835 欢度六一儿童节主题活动通用模板.pptx",
     "50129835 欢度六一儿童节主题活动通用模板_clone.pptx"),
    ("50329450 黑色科技风高端新品发布会通用PPT模板.pptx",
     "50329450 黑色科技风高端新品发布会通用PPT模板_clone.pptx"),
    ("总结汇报-图表.pptx",
     "总结汇报-图表_clone.pptx"),
]


def _index_shapes(story: dict) -> dict:
    """shape_id -> {para_sizes, para_chars, char_count, font_size_pt}"""
    idx = {}
    for s in story.get("slides", []):
        for cs in s.get("content_shapes", []):
            idx[cs["shape_id"]] = {
                "slide_index": s["slide_index"],
                "para_sizes": cs.get("per_paragraph_font_size_pt", []) or [],
                "para_chars": cs.get("per_paragraph_char_count", []) or [],
                "char_count": int(cs.get("char_count", 0) or 0),
                "font_size_pt": cs.get("font_size_pt"),
                "paragraph_count": int(cs.get("paragraph_count", 1) or 1),
            }
    return idx


def verify_pair(src_name: str, clone_name: str) -> dict:
    src = ROOT / src_name
    clone = ROOT / clone_name
    if not src.exists() or not clone.exists():
        return {"src": src_name, "ok": False, "error": "file_missing"}

    story_src = pts.parse(src, thresholds=pts.DEFAULT_THRESHOLDS)
    story_clone = pts.parse(clone, thresholds=pts.DEFAULT_THRESHOLDS)
    idx_src = _index_shapes(story_src)
    idx_clone = _index_shapes(story_clone)

    para_count_mismatch = 0
    font_deltas: list[float] = []
    char_ratios: list[float] = []
    large_font_samples: list[dict] = []
    missing_in_clone = 0

    for sid, s in idx_src.items():
        c = idx_clone.get(sid)
        if not c:
            missing_in_clone += 1
            continue
        if s["paragraph_count"] != c["paragraph_count"]:
            para_count_mismatch += 1

        # 段级比较
        n = min(len(s["para_sizes"]), len(c["para_sizes"]))
        for i in range(n):
            src_sz = float(s["para_sizes"][i]) if s["para_sizes"][i] else 0.0
            clo_sz = float(c["para_sizes"][i]) if c["para_sizes"][i] else 0.0
            if src_sz > 0 and clo_sz > 0:
                delta = abs(clo_sz - src_sz)
                font_deltas.append(delta)
                # 大字号样例抽取
                if src_sz >= 28 and len(large_font_samples) < 8:
                    src_chars = (s["para_chars"][i]
                                 if i < len(s["para_chars"]) else 0)
                    clo_chars = (c["para_chars"][i]
                                 if i < len(c["para_chars"]) else 0)
                    large_font_samples.append({
                        "slide": s["slide_index"],
                        "shape_id": sid.split("::")[-1],
                        "para_idx": i,
                        "src_size": src_sz,
                        "clone_size": clo_sz,
                        "src_chars": src_chars,
                        "clone_chars": clo_chars,
                    })
            # 字数比
            src_cc = (s["para_chars"][i]
                      if i < len(s["para_chars"]) else 0)
            clo_cc = (c["para_chars"][i]
                      if i < len(c["para_chars"]) else 0)
            if src_cc > 0:
                char_ratios.append(clo_cc / src_cc)

    def _bucket(values, edges):
        if not values:
            return {}
        out = {f"≤{e}": 0 for e in edges}
        out[f">{edges[-1]}"] = 0
        for v in values:
            placed = False
            for e in edges:
                if v <= e:
                    out[f"≤{e}"] += 1
                    placed = True
                    break
            if not placed:
                out[f">{edges[-1]}"] += 1
        return out

    return {
        "src": src_name,
        "clone": clone_name,
        "ok": True,
        "shapes_src": len(idx_src),
        "shapes_clone": len(idx_clone),
        "missing_in_clone": missing_in_clone,
        "paragraph_count_mismatch": para_count_mismatch,
        "font_delta_stats": {
            "count": len(font_deltas),
            "mean": round(statistics.fmean(font_deltas), 2) if font_deltas else 0,
            "p95": (round(sorted(font_deltas)[int(len(font_deltas) * 0.95)], 2)
                    if len(font_deltas) >= 5 else None),
            "max": round(max(font_deltas), 2) if font_deltas else 0,
            "buckets": _bucket(font_deltas, [0.5, 1.0, 2.0, 4.0]),
        },
        "char_ratio_stats": {
            "count": len(char_ratios),
            "mean": round(statistics.fmean(char_ratios), 3) if char_ratios else 0,
            "buckets": _bucket(char_ratios, [0.5, 0.8, 1.0, 1.1]),
        },
        "large_font_samples": large_font_samples,
    }


def main() -> int:
    import json
    any_fail = False
    for src_name, clone_name in PAIRS:
        r = verify_pair(src_name, clone_name)
        print(f"\n=== {src_name} ===")
        print(json.dumps(r, ensure_ascii=False, indent=2))
        if not r.get("ok"):
            any_fail = True
    return 0 if not any_fail else 1


if __name__ == "__main__":
    sys.exit(main())
