"""template_story.json (v2) + 新大纲 → content_blueprint.json (v2) 工作流。

本脚本不调 LLM，只做三件事：
- scaffold  ：生成 blueprint_scaffold.json，按 shape 粒度暴露 char_limit / shape_groups /
              story_role / global_style_guide
- validate  ：校验 Agent 产出的 content_blueprint.json 是否满足 shape 级硬约束
- rollback  ：Prompt C-retry 用：把未被 errors 指向的 shape 回退为上一版

Blueprint v2 数据结构：
{
  "schema_version": 2,
  "title": "新主题",
  "global_style_guide": {...},
  "slides": [
    {
      "slide_index": 1,
      "page_theme": "...",
      "story_role": "cover",
      "shape_texts": {
        "slide_1::sp_5": "...",                         # 单段
        "slide_1::sp_8": ["第一段", "第二段", "第三段"], # 多段落
        "slide_1::sp_99": "__preserve__"                # 保留模板原文
      }
    }
  ]
}

CLI:
  python build_content_blueprint.py scaffold \\
      --story template_story.json [--outline outline.txt | --topic-json topic.json] \\
      --out blueprint_scaffold.json

  python build_content_blueprint.py validate \\
      --story template_story.json --blueprint content_blueprint.json [--strict]

  python build_content_blueprint.py rollback \\
      --prev previous_blueprint.json --new retry_blueprint.json \\
      --errors error_shapes.json --out merged_blueprint.json
"""

from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path

SCHEMA_VERSION = 2

PRESERVE_TOKEN = "__preserve__"


# === scaffold ===

def build_scaffold(story: dict, *, user_outline: str = "",
                   topic_meta: dict | None = None) -> dict:
    """产 blueprint_scaffold.json：LLM 填 shape_texts 之前的结构骨架。

    每页暴露：
    - story_role
    - content_shapes 列表（每个 shape 完整 char_limit / paragraph_count / hint）
    - shape_groups 弱提示（同组并列约束）
    - non_content_summary（仅统计装饰类型，告知 LLM 不必管）
    """
    if story.get("schema_version") != SCHEMA_VERSION:
        # 兼容性提示，但不阻断
        pass

    slides_out: list[dict] = []
    decoration_types_summary: dict[str, int] = {}

    for s in story.get("slides", []):
        content_shapes_brief: list[dict] = []
        for cs in s.get("content_shapes", []):
            text_hint = (cs.get("original_text") or "").replace("\n", " ").strip()
            text_preview = text_hint[:30] + ("..." if len(text_hint) > 30 else "")
            brief = {
                "shape_id": cs["shape_id"],
                "bbox_region": cs.get("bbox_region"),
                "char_limit": cs.get("char_limit", {"min": 1, "max": 50}),
                "hard_ceiling_chars": cs.get("hard_ceiling_chars"),
                "paragraph_count": cs.get("paragraph_count", 1),
                "original_char_count": int(cs.get("char_count", 0) or 0),
                "original_paragraphs_char_count":
                    cs.get("per_paragraph_char_count", []),
                "per_paragraph_char_count": cs.get("per_paragraph_char_count", []),
                "font_size_pt": cs.get("font_size_pt"),
                "original_text_preview": text_preview,
                "is_placeholder_text": bool(cs.get("is_placeholder_text")),
                "has_bullet": bool(cs.get("has_bullet")),
                "ph_type": cs.get("ph_type"),
            }
            if cs.get("parent_table_id"):
                brief["parent_table_id"] = cs["parent_table_id"]
                brief["table_cell_index"] = cs.get("table_cell_index")
            # 多段 shape 暴露 per_paragraph hint（零估算：严格等于原文字数）
            pp_list = cs.get("per_paragraph") or []
            if pp_list:
                brief["per_paragraph"] = [
                    {
                        "idx": pp["idx"],
                        "font_size_pt": pp.get("font_size_pt"),
                        "original_char_count": pp.get("original_char_count"),
                        "char_limit": pp.get("char_limit"),
                        "hard_ceiling_chars": pp.get("hard_ceiling_chars"),
                        "is_emphasis": bool(pp.get("is_emphasis")),
                        "hint": (
                            "emphasis: 字号偏大，优先用短词/名词"
                            if pp.get("is_emphasis")
                            else "body: 正文段"
                        ),
                    }
                    for pp in pp_list
                ]
            content_shapes_brief.append(brief)

        # 统计 non_content
        for nc in s.get("non_content_shapes", []):
            role = nc.get("role", "unknown")
            decoration_types_summary[role] = \
                decoration_types_summary.get(role, 0) + 1

        slides_out.append({
            "slide_index": s["slide_index"],
            "story_role": s.get("story_role"),
            "content_shapes": content_shapes_brief,
            "shape_groups": s.get("shape_groups", []),
            "non_content_count": len(s.get("non_content_shapes", [])),
        })

    global_style_guide = _default_style_guide(user_outline, topic_meta)

    total_shapes = sum(len(x["content_shapes"]) for x in slides_out)
    total_groups = sum(len(x["shape_groups"]) for x in slides_out)
    placeholder_shapes = sum(
        1 for x in slides_out for sh in x["content_shapes"]
        if sh.get("is_placeholder_text"))

    return {
        "schema_version": SCHEMA_VERSION,
        "artifact_type": "blueprint_scaffold",
        "title": (topic_meta or {}).get("title", ""),
        "user_outline": user_outline,
        "global_style_guide": global_style_guide,
        "slides": slides_out,
        "decoration_types_summary": decoration_types_summary,
        "meta": {
            "slide_count": len(slides_out),
            "content_shape_total": total_shapes,
            "shape_group_total": total_groups,
            "placeholder_shape_total": placeholder_shapes,
            "strict_char_equal": True,
            "replace_policy": "replace_all_by_default",
            "preserve_only_examples": [
                "固定品牌名 / logo 文字",
                "版权声明 / 公司全称",
                "固定日期格式",
            ],
            "preserve_soft_cap_ratio": 0.2,
            "strict_char_rule_summary": (
                "每段新文本字数必须严格等于或略少于该段 original_char_count。"
                "char_limit.max 就是原文字数，不是预估上限；"
                "段数（paragraphs）必须与 paragraph_count 一致。"
                "is_emphasis=true 仅作字体大、用词宜短的风格 hint，不放宽字数。"
            ),
            "replacement_guideline": (
                "默认对 scaffold 里所有 content_shape 都给出新主题的真实文本。"
                "non_content_shapes (LOGO / 序号 / 装饰) 已被过滤，不会出现在 "
                "scaffold。__preserve__ 仅限固定品牌/版权/固定日期等极少数例外，"
                "单页 __preserve__ 占比建议 < 20%。"
            ),
        },
    }


def _default_style_guide(user_outline: str, topic_meta: dict | None) -> dict:
    topic_meta = topic_meta or {}
    return {
        "thesis_prompt": topic_meta.get("thesis_prompt", ""),
        "terminology": topic_meta.get("terminology", []),
        "must_use_words": topic_meta.get("must_use_words", []),
        "banned_words": topic_meta.get("banned_words", []),
        "tone": topic_meta.get("tone", "professional"),
        "voice": topic_meta.get("voice", "third_person"),
        "language": topic_meta.get("language", "zh-CN"),
        "unit": topic_meta.get("unit", ""),
        "seasonal_hint": topic_meta.get("seasonal_hint", ""),
    }


# === validate ===

def _count_chars(text: str) -> int:
    if not text:
        return 0
    return sum(1 for c in text if c not in (" ", "\n", "\t", "\r", "\u3000"))


def validate_blueprint(story: dict, blueprint: dict) -> dict:
    """校验 v2 Blueprint：shape_id 齐全 + 字数硬上限 + 多段落段数匹配 + __preserve__ 语义。"""
    errors: list[dict] = []
    warnings: list[dict] = []

    if blueprint.get("schema_version") != SCHEMA_VERSION:
        warnings.append({
            "kind": "schema_version_mismatch",
            "expected": SCHEMA_VERSION,
            "actual": blueprint.get("schema_version"),
        })

    story_slides = {s["slide_index"]: s for s in story.get("slides", [])}
    bp_slides = {s["slide_index"]: s for s in blueprint.get("slides", [])}

    if set(story_slides) != set(bp_slides):
        missing = sorted(set(story_slides) - set(bp_slides))
        extra = sorted(set(bp_slides) - set(story_slides))
        if missing:
            errors.append({"kind": "missing_slide_in_blueprint",
                           "slide_indices": missing})
        if extra:
            errors.append({"kind": "extra_slide_in_blueprint",
                           "slide_indices": extra})

    for idx, ss in story_slides.items():
        if idx not in bp_slides:
            continue
        bs = bp_slides[idx]
        cs_by_id = {cs["shape_id"]: cs for cs in ss.get("content_shapes", [])}
        bp_texts = bs.get("shape_texts", {}) or {}

        # 必须覆盖每个 content_shape
        preserve_on_content_count = 0
        for sid, cs in cs_by_id.items():
            if sid not in bp_texts:
                errors.append({
                    "kind": "missing_shape_text",
                    "slide_index": idx, "shape_id": sid,
                    "msg": f"slide_{idx} 缺少 shape_id={sid} 的 shape_texts 条目",
                })
                continue
            value = bp_texts[sid]
            _check_shape_value(idx, sid, cs, value, errors, warnings)

            # 默认全替换策略：对 is_placeholder_text=false 的 shape 写 __preserve__
            # 给出 warning，引导 LLM/Agent 改为实际文本。
            if (
                isinstance(value, str)
                and value.strip() == PRESERVE_TOKEN
                and not cs.get("is_placeholder_text")
            ):
                preserve_on_content_count += 1
                warnings.append({
                    "kind": "unnecessary_preserve_on_content",
                    "slide_index": idx, "shape_id": sid,
                    "msg": (
                        f"slide_{idx}.{sid} 是 content_shape 但写了 __preserve__，"
                        f"换主题语境下大概率应替换为新文本。__preserve__ 仅限固定品牌/"
                        f"版权/固定日期等极少数例外。"
                    ),
                })

        # 单页 __preserve__ 占比 > 30% 给 warning
        total_cs = len(cs_by_id)
        if total_cs > 0 and preserve_on_content_count / total_cs > 0.3:
            warnings.append({
                "kind": "too_many_preserves_on_slide",
                "slide_index": idx,
                "preserve_count": preserve_on_content_count,
                "content_shape_total": total_cs,
                "ratio": round(preserve_on_content_count / total_cs, 2),
                "msg": (
                    f"slide_{idx} 有 {preserve_on_content_count}/{total_cs} "
                    f"个 content_shape 使用了 __preserve__，占比过高"
                    f"（>30%），换主题语义稀薄，建议改为真实文本。"
                ),
            })

        # 多余的 shape_id（不在 story 里）
        for extra_id in set(bp_texts) - set(cs_by_id):
            warnings.append({
                "kind": "unknown_shape_id",
                "slide_index": idx, "shape_id": extra_id,
                "msg": f"slide_{idx} 出现未知 shape_id={extra_id}",
            })

        # page_theme 软警告
        if not bs.get("page_theme"):
            warnings.append({
                "kind": "missing_page_theme",
                "slide_index": idx,
            })

    return {"errors": errors, "warnings": warnings}


def _check_shape_value(slide_idx: int, shape_id: str, cs: dict,
                       value, errors: list, warnings: list) -> None:
    """单 shape 校验：__preserve__ / string / string[] 三种合法形态。"""
    char_limit = cs.get("char_limit", {"min": 1, "max": 50})
    max_chars = int(char_limit.get("max", 50))
    min_chars = int(char_limit.get("min", 1))
    expected_para = int(cs.get("paragraph_count", 1) or 1)
    per_paragraph = cs.get("per_paragraph") or []

    # __preserve__ 合法
    if isinstance(value, str) and value.strip() == PRESERVE_TOKEN:
        # 占位文本 shape 不允许 __preserve__（必须替换）
        if cs.get("is_placeholder_text"):
            errors.append({
                "kind": "preserve_on_placeholder",
                "slide_index": slide_idx, "shape_id": shape_id,
                "msg": f"slide_{slide_idx}.{shape_id} 是占位文本，不可 __preserve__",
            })
        return

    if isinstance(value, str):
        # 单段
        text = value
        chars = _count_chars(text)
        if expected_para > 1:
            errors.append({
                "kind": "paragraph_count_strict_mismatch",
                "slide_index": slide_idx, "shape_id": shape_id,
                "expected_paragraphs": expected_para, "actual_paragraphs": 1,
                "msg": (f"slide_{slide_idx}.{shape_id} 原 {expected_para} 段，"
                        f"blueprint 只给了单字符串；段数必须严格与原模板一致，"
                        f"请拆成长度为 {expected_para} 的数组。"),
            })
        if max_chars > 0 and chars > max_chars:
            errors.append({
                "kind": "char_count_exceeds_limit",
                "slide_index": slide_idx, "shape_id": shape_id,
                "chars": chars, "max_chars": max_chars,
                "msg": (f"slide_{slide_idx}.{shape_id} 字数 {chars} > "
                        f"原文字数 {max_chars}（char_limit.max 等于原文）"),
            })
        elif chars > 0 and chars < min_chars:
            warnings.append({
                "kind": "char_count_below_min",
                "slide_index": slide_idx, "shape_id": shape_id,
                "chars": chars, "min_chars": min_chars,
                "msg": (f"slide_{slide_idx}.{shape_id} 字数 {chars} < "
                        f"min {min_chars}（原文基本等字数，偏短易留白）"),
            })
        return

    if isinstance(value, list):
        # 多段
        if expected_para > 1 and len(value) != expected_para:
            errors.append({
                "kind": "paragraph_count_strict_mismatch",
                "slide_index": slide_idx, "shape_id": shape_id,
                "expected": expected_para, "actual": len(value),
                "msg": (f"slide_{slide_idx}.{shape_id} 原 {expected_para} 段，"
                        f"blueprint 给 {len(value)} 段；段数必须严格一致"
                        f"（apply 克隆/清空仅作兜底，可能导致字号错乱）。"),
            })
        total_chars = sum(_count_chars(p) for p in value if isinstance(p, str))
        if max_chars > 0 and total_chars > max_chars:
            errors.append({
                "kind": "char_count_exceeds_limit",
                "slide_index": slide_idx, "shape_id": shape_id,
                "chars": total_chars, "max_chars": max_chars,
                "msg": (f"slide_{slide_idx}.{shape_id} 多段总字数 {total_chars} > "
                        f"原文总字数 {max_chars}"),
            })
        for i, item in enumerate(value):
            if not isinstance(item, str):
                errors.append({
                    "kind": "invalid_paragraph_value_type",
                    "slide_index": slide_idx, "shape_id": shape_id,
                    "paragraph_index": i,
                    "actual_type": type(item).__name__,
                })
        # per_paragraph 逐段校验（严格对齐原文字数）
        if per_paragraph:
            for i, item in enumerate(value):
                if not isinstance(item, str):
                    continue
                pp = per_paragraph[i] if i < len(per_paragraph) else None
                if not pp:
                    continue
                p_limit = pp.get("char_limit") or {}
                p_max = int(p_limit.get("max", 0))
                p_min = int(p_limit.get("min", 0))
                p_orig = int(pp.get("original_char_count", p_max) or p_max)
                p_chars = _count_chars(item)
                is_emphasis = bool(pp.get("is_emphasis"))
                if p_max > 0 and p_chars > p_max:
                    errors.append({
                        "kind": "paragraph_char_count_exceeds_limit",
                        "slide_index": slide_idx, "shape_id": shape_id,
                        "paragraph_index": i,
                        "chars": p_chars, "max_chars": p_max,
                        "original_char_count": p_orig,
                        "is_emphasis": is_emphasis,
                        "msg": (f"slide_{slide_idx}.{shape_id} 段 {i} "
                                f"字数 {p_chars} > 原文字数 {p_orig}"
                                + ("（emphasis 段，字号偏大更易溢出）"
                                   if is_emphasis else "")),
                    })
                # 段字数显著少于原文 -> 轻量 warning（允许，但提示留白）
                if p_max > 0 and p_chars > 0 and p_chars < max(1, p_orig - 3):
                    warnings.append({
                        "kind": "paragraph_char_count_below_original",
                        "slide_index": slide_idx, "shape_id": shape_id,
                        "paragraph_index": i,
                        "chars": p_chars, "original_char_count": p_orig,
                        "msg": (f"段 {i} 字数 {p_chars} 显著少于原文 {p_orig}；"
                                "若非有意留白，建议靠近原文字数以避免版面空洞"),
                    })
        return

    # 其他类型
    errors.append({
        "kind": "invalid_shape_value_type",
        "slide_index": slide_idx, "shape_id": shape_id,
        "actual_type": type(value).__name__,
        "msg": "shape_texts[shape_id] 必须为 string / string[] / '__preserve__'",
    })


# === rollback (Prompt C-retry 幂等性) ===

def rollback_unchanged(prev_bp: dict, new_bp: dict,
                       error_shapes: list[dict]) -> dict:
    """把"未出错的 shape"回退到上一版；只保留新 bp 中错误的 shape 的改动。

    error_shapes 格式：[{"slide_index": N, "shape_id": "..."}]
    """
    error_set = {(int(e["slide_index"]), e["shape_id"]) for e in error_shapes}
    merged = copy.deepcopy(new_bp)
    prev_slide_map = {s["slide_index"]: s for s in prev_bp.get("slides", [])}

    for s in merged.get("slides", []):
        idx = s["slide_index"]
        prev_s = prev_slide_map.get(idx)
        if not prev_s:
            continue
        prev_texts = prev_s.get("shape_texts", {}) or {}
        for sid, val in list(s.get("shape_texts", {}).items()):
            key = (idx, sid)
            if key in error_set:
                continue  # 保留新值
            if sid in prev_texts:
                s["shape_texts"][sid] = copy.deepcopy(prev_texts[sid])
    return merged


# === CLI ===

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    sc = sub.add_parser("scaffold", help="生成 blueprint_scaffold.json 喂给 Prompt C")
    sc.add_argument("--story", required=True, type=Path)
    sc.add_argument("--outline", type=Path, default=None,
                    help="用户大纲文本文件")
    sc.add_argument("--topic-json", type=Path, default=None,
                    help="可选：{title,thesis_prompt,terminology,tone,voice,language,...}")
    sc.add_argument("--out", required=True, type=Path)

    vv = sub.add_parser("validate", help="校验 content_blueprint.json")
    vv.add_argument("--story", required=True, type=Path)
    vv.add_argument("--blueprint", required=True, type=Path)
    vv.add_argument("--strict", action="store_true")

    rb = sub.add_parser("rollback",
                        help="Prompt C-retry：把未报错 shape 回退到上一版")
    rb.add_argument("--prev", required=True, type=Path)
    rb.add_argument("--new", required=True, type=Path)
    rb.add_argument("--errors", required=True, type=Path,
                    help="[{slide_index,shape_id},...] json")
    rb.add_argument("--out", required=True, type=Path)

    args = ap.parse_args()

    if args.cmd == "scaffold":
        story = json.loads(args.story.read_text(encoding="utf-8"))
        outline_text = args.outline.read_text(encoding="utf-8") \
            if args.outline and args.outline.exists() else ""
        topic_meta = json.loads(args.topic_json.read_text(encoding="utf-8")) \
            if args.topic_json and args.topic_json.exists() else None
        scaffold = build_scaffold(story, user_outline=outline_text,
                                  topic_meta=topic_meta)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(scaffold, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        print(json.dumps({"ok": True, "out": str(args.out),
                          "meta": scaffold["meta"]}, ensure_ascii=False))
        return 0

    if args.cmd == "validate":
        story = json.loads(args.story.read_text(encoding="utf-8"))
        blueprint = json.loads(args.blueprint.read_text(encoding="utf-8"))
        result = validate_blueprint(story, blueprint)
        if args.strict:
            result["errors"].extend(result.pop("warnings", []))
            result["warnings"] = []
        ok = not result["errors"]
        print(json.dumps({
            "ok": ok,
            "error_count": len(result["errors"]),
            "warning_count": len(result["warnings"]),
            "errors": result["errors"],
            "warnings": result["warnings"],
        }, ensure_ascii=False, indent=2))
        return 0 if ok else 2

    if args.cmd == "rollback":
        prev_bp = json.loads(args.prev.read_text(encoding="utf-8"))
        new_bp = json.loads(args.new.read_text(encoding="utf-8"))
        errs = json.loads(args.errors.read_text(encoding="utf-8"))
        if isinstance(errs, dict) and "errors" in errs:
            error_shapes = [
                {"slide_index": e["slide_index"], "shape_id": e["shape_id"]}
                for e in errs["errors"]
                if "slide_index" in e and "shape_id" in e
            ]
        elif isinstance(errs, list):
            error_shapes = errs
        else:
            error_shapes = []
        merged = rollback_unchanged(prev_bp, new_bp, error_shapes)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(merged, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        print(json.dumps({
            "ok": True, "out": str(args.out),
            "rollback_count": len(error_shapes),
        }, ensure_ascii=False))
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
