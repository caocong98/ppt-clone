"""把 content_blueprint.json (v2) 投射到 template_story.json (v2) 的 shape 上。

不调 LLM。唯一会触发 LLM 的情况：某 shape 文本超过 char_limit.max → 写入
pending_shrink_requests，由 Agent 调 Prompt D 压缩后通过 `--resume` 回灌。

content_mapping v4 schema：
{
  "schema_version": 4,
  "slides": {
    "1": {
      "story_role": "cover",
      "page_theme": "...",
      "assignments": [
        {"key": "slide_1::sp_5", "value": "新标题", "shape_id": "slide_1::sp_5",
         "source": "blueprint"},
        {"key": "slide_1::sp_8", "value": ["第一段", "第二段"], "shape_id": "...",
         "source": "blueprint_paragraphs"},
        {"key": "slide_1::sp_9", "value": "__skip__", "shape_id": "...",
         "source": "blueprint_preserve"},
        {"key": "slide_1::sp_99", "value": "__clear__", "shape_id": "...",
         "source": "non_content", "preserve_action": "clear_to_empty"}
      ]
    }
  },
  "pending_shrink_requests": [
    {"slide_index": N, "shape_id": "...", "text": "...",
     "char_limit": {"min": 8, "max": 14}, "previous_attempts": [],
     "reason": "exceed_max_chars"}
  ]
}

CLI:
  python map_blueprint_to_template.py --story story.json --blueprint bp.json --out mapping.json
  python map_blueprint_to_template.py --story story.json --blueprint bp.json \\
      --resume shrink_results.json --previous-mapping content_mapping.json --out mapping.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCHEMA_VERSION = 4

PRESERVE_TOKEN = "__preserve__"
SKIP_TOKEN = "__skip__"
CLEAR_TOKEN = "__clear__"


def _count_chars(text: str) -> int:
    if not text:
        return 0
    return sum(1 for c in text if c not in (" ", "\n", "\t", "\r", "\u3000"))


def _resolve_shape_value(value):
    """把 Blueprint 的 string|string[]|__preserve__ 转为 apply 层认得的形态。

    返回 (apply_value, is_preserve, total_chars)。
    """
    if isinstance(value, str) and value.strip() == PRESERVE_TOKEN:
        return SKIP_TOKEN, True, 0
    if isinstance(value, str):
        text = value
        return text, False, _count_chars(text)
    if isinstance(value, list):
        cleaned = [str(p) for p in value]
        total = sum(_count_chars(p) for p in cleaned)
        return cleaned, False, total
    # 兜底
    return str(value), False, _count_chars(str(value))


def map_blueprint_to_mapping(story: dict, blueprint: dict) -> dict:
    bp_slides = {s["slide_index"]: s for s in blueprint.get("slides", [])}
    slides_out: dict[str, dict] = {}
    pending_shrinks: list[dict] = []

    for ss in story.get("slides", []):
        idx = ss["slide_index"]
        bp_slide = bp_slides.get(idx, {})
        bp_texts = bp_slide.get("shape_texts", {}) or {}
        page_theme = bp_slide.get("page_theme", "")
        assignments: list[dict] = []

        # 1) content_shapes：直接按 shape_id 翻译
        for cs in ss.get("content_shapes", []):
            sid = cs["shape_id"]
            char_limit = cs.get("char_limit", {"min": 1, "max": 50})
            max_chars = int(char_limit.get("max", 50))
            hard_ceiling = int(cs.get("hard_ceiling_chars") or max_chars)
            per_paragraph = cs.get("per_paragraph") or []
            per_para_limits = [
                {
                    "idx": pp.get("idx"),
                    "char_limit": pp.get("char_limit"),
                    "hard_ceiling_chars": pp.get("hard_ceiling_chars"),
                    "is_emphasis": bool(pp.get("is_emphasis")),
                    "font_size_pt": pp.get("font_size_pt"),
                }
                for pp in per_paragraph
            ]

            if sid not in bp_texts:
                # Blueprint 漏了 → 默认 keep_original
                assignments.append({
                    "key": sid,
                    "value": SKIP_TOKEN,
                    "shape_id": sid,
                    "source": "blueprint_missing_skip",
                })
                continue

            apply_value, is_preserve, total_chars = _resolve_shape_value(
                bp_texts[sid])

            if is_preserve:
                assignments.append({
                    "key": sid,
                    "value": SKIP_TOKEN,
                    "shape_id": sid,
                    "source": "blueprint_preserve",
                })
                continue

            # 字数硬上限：超过 hard_ceiling 才写入 pending_shrink_requests
            # （>max 不强制压缩，apply 层会按 fill_ratio 截断）
            shrink_triggered = False
            # 列表值：逐段检查段级上限
            if isinstance(apply_value, list) and per_para_limits:
                for i, seg in enumerate(apply_value):
                    if i >= len(per_para_limits):
                        break
                    pl = per_para_limits[i].get("char_limit") or {}
                    p_max = int(pl.get("max", 0))
                    p_hard = int(per_para_limits[i].get(
                        "hard_ceiling_chars") or p_max)
                    seg_chars = _count_chars(seg)
                    if p_hard and seg_chars > p_hard:
                        pending_shrinks.append({
                            "slide_index": idx,
                            "shape_id": sid,
                            "paragraph_idx": i,
                            "text": seg,
                            "char_limit": per_para_limits[i].get("char_limit"),
                            "hard_ceiling_chars": p_hard,
                            "current_chars": seg_chars,
                            "is_emphasis": per_para_limits[i].get("is_emphasis"),
                            "previous_attempts": [],
                            "reason": "paragraph_exceed_hard_ceiling",
                        })
                        shrink_triggered = True
            # 整 shape 总字数超过 hard_ceiling 也入列（主要给单段 string 用）
            if not shrink_triggered and hard_ceiling and \
                    total_chars > hard_ceiling:
                pending_shrinks.append({
                    "slide_index": idx,
                    "shape_id": sid,
                    "text": apply_value if isinstance(apply_value, str) else
                            "\n".join(apply_value),
                    "char_limit": char_limit,
                    "hard_ceiling_chars": hard_ceiling,
                    "current_chars": total_chars,
                    "previous_attempts": [],
                    "reason": "exceed_hard_ceiling",
                })

            source = "blueprint_paragraphs" if isinstance(apply_value, list) \
                else "blueprint"
            assignment = {
                "key": sid,
                "value": apply_value,
                "shape_id": sid,
                "source": source,
                "char_count": total_chars,
                "char_limit": char_limit,
                "hard_ceiling_chars": hard_ceiling,
                "is_placeholder_text": bool(cs.get("is_placeholder_text")),
                # shape 级 dominant 字号（apply 层做 font-pinning 用）
                "font_size_pt": cs.get("font_size_pt"),
                # 兼容：有些 shape 只有单段，这里再导一次段级字号，便于
                # apply 层在克隆段落 / 新增段落时有明确目标字号。
                "per_paragraph_font_size_pt":
                    cs.get("per_paragraph_font_size_pt") or [],
            }
            if per_para_limits:
                assignment["per_paragraph_char_limits"] = per_para_limits
            assignments.append(assignment)

        # 2) non_content_shapes：按 preserve_action 生成 __skip__ / __clear__
        for nc in ss.get("non_content_shapes", []):
            sid = nc["shape_id"]
            action = nc.get("preserve_action", "keep_original")
            value = CLEAR_TOKEN if action == "clear_to_empty" else SKIP_TOKEN
            assignments.append({
                "key": sid,
                "value": value,
                "shape_id": sid,
                "source": "non_content",
                "role": nc.get("role"),
                "preserve_action": action,
            })

        slides_out[str(idx)] = {
            "story_role": ss.get("story_role"),
            "page_theme": page_theme,
            "assignments": assignments,
        }

    return {
        "schema_version": SCHEMA_VERSION,
        "slides": slides_out,
        "pending_shrink_requests": pending_shrinks,
        "global_style_guide": blueprint.get("global_style_guide"),
        "meta": {
            "assignment_total": sum(len(v["assignments"]) for v in slides_out.values()),
            "pending_shrink_total": len(pending_shrinks),
        },
    }


def _dedup_vs_siblings(shape_id: str, new_text: str, mapping: dict) -> bool:
    """检查新文本与同 slide 其他 shape 是否重复（避免压缩后撞车）。"""
    target_slide_idx = None
    for s_key, page in mapping.get("slides", {}).items():
        for a in page.get("assignments", []):
            if a.get("shape_id") == shape_id:
                target_slide_idx = s_key
                break
        if target_slide_idx is not None:
            break
    if target_slide_idx is None:
        return False
    page = mapping["slides"][target_slide_idx]
    for a in page.get("assignments", []):
        if a.get("shape_id") == shape_id:
            continue
        v = a.get("value")
        if v in (SKIP_TOKEN, CLEAR_TOKEN):
            continue
        cmp_text = v if isinstance(v, str) else "\n".join(v) \
            if isinstance(v, list) else None
        if cmp_text and cmp_text.strip() == new_text.strip():
            return True
    return False


def apply_shrink_resume(mapping: dict, shrink_results: dict) -> dict:
    """把 Prompt D 产出的 shrink 结果回填到 mapping。

    shrink_results 格式：
    {
      "results": [
        {"slide_index": N, "shape_id": "slide_N::sp_X",
         "text": "压缩后", "previous_attempts": [...]}
      ]
    }
    """
    fallback_required: list[dict] = []
    for r in shrink_results.get("results", []):
        idx = int(r["slide_index"])
        sid = r["shape_id"]
        new_text = (r.get("text") or "").strip()
        page = mapping.get("slides", {}).get(str(idx))
        if not page:
            fallback_required.append({**r, "reason": "slide_not_found"})
            continue

        matched = False
        for a in page["assignments"]:
            if a.get("shape_id") != sid:
                continue
            if isinstance(a["value"], list):
                # 多段：把整个值替换为单段（除非 LLM 给的是带换行的多段文本）
                if "\n" in new_text:
                    a["value"] = [p for p in new_text.split("\n") if p.strip()]
                else:
                    a["value"] = new_text
            else:
                a["value"] = new_text
            a["source"] = "shrink_resumed"
            matched = True
            break

        if not matched:
            fallback_required.append({**r, "reason": "assignment_not_found"})
            continue

        if _dedup_vs_siblings(sid, new_text, mapping):
            fallback_required.append({
                **r, "reason": "dedup_conflict_after_shrink",
            })

    # 从 pending_shrink_requests 移除已处理的 (slide_index, shape_id)
    result_keys = {(int(r["slide_index"]), r["shape_id"])
                   for r in shrink_results.get("results", [])}
    pending = [p for p in mapping.get("pending_shrink_requests", [])
               if (p["slide_index"], p["shape_id"]) not in result_keys]
    mapping["pending_shrink_requests"] = pending

    if fallback_required:
        mapping.setdefault("fallback_required", []).extend(fallback_required)

    mapping["meta"]["pending_shrink_total"] = len(pending)
    mapping["meta"]["shrink_applied"] = len(shrink_results.get("results", []))
    mapping["meta"]["shrink_fallback_count"] = len(fallback_required)
    return mapping


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--story", required=True, type=Path)
    ap.add_argument("--blueprint", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--resume", type=Path, default=None,
                    help="shrink_results.json；提供则回灌 Prompt D 结果")
    ap.add_argument("--previous-mapping", type=Path, default=None,
                    help="--resume 模式下已有 content_mapping.json")
    args = ap.parse_args()

    story = json.loads(args.story.read_text(encoding="utf-8"))
    blueprint = json.loads(args.blueprint.read_text(encoding="utf-8"))

    if args.resume is not None:
        prev = args.previous_mapping or args.out
        if not prev.exists():
            print(f"ERROR: --resume 需要 --previous-mapping 指向已有 mapping: "
                  f"{prev}", file=sys.stderr)
            return 1
        mapping = json.loads(prev.read_text(encoding="utf-8"))
        shrink_results = json.loads(args.resume.read_text(encoding="utf-8"))
        mapping = apply_shrink_resume(mapping, shrink_results)
    else:
        mapping = map_blueprint_to_mapping(story, blueprint)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(mapping, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    print(json.dumps({
        "ok": True, "out": str(args.out),
        "assignment_total": mapping["meta"]["assignment_total"],
        "pending_shrink_total": mapping["meta"]["pending_shrink_total"],
        "fallback_required": mapping["meta"].get("shrink_fallback_count", 0),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
