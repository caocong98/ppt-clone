"""统一校验工具：把 LLM 产物挡在落盘前。

支持的校验：
- story     : template_story.json 自身合法性（slot capacity/shape_id 非空等）
- blueprint : content_blueprint.json 是否符合 story 的 capacity / 长度约束
              （委托 build_content_blueprint.validate_blueprint）
- mapping   : content_mapping.json v3 是否符合 story（含同页重复 / list 完全一致 /
              non_content 被填 等新检查；兼容 v1 扁平 key）
- decision  : decision.json 的 12 个主题槽位是否完整、合法、对比度达标
- outline   : outline_plan.json（旧方案兼容）

退出码：0 通过 / 2 失败 / 1 输入错误
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from color_utils import normalize_hex, wcag_contrast  # noqa: E402
from build_content_blueprint import validate_blueprint as _validate_bp  # noqa: E402

KEY_RE = re.compile(r"^slide_(\d+)::(.+)$")

THEME_SLOTS = ["dk1", "lt1", "dk2", "lt2",
               "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink"]

VALID_STORY_ROLES = {"cover", "toc", "section_divider", "content",
                     "content_list", "summary", "transition", "thanks", "other"}
VALID_SLOT_KINDS = {"single", "list_slot", "enumeration_slot", "bullet_group_slot"}
VALID_OUTLINE_ROLES = {"cover", "toc", "section", "content",
                       "summary", "transition", "thanks", "other"}


# ============================================================
# story 校验
# ============================================================

def validate_story(story: dict) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    slides = story.get("slides")
    if not isinstance(slides, list):
        errors.append({"kind": "missing_slides_array"})
        return {"errors": errors, "warnings": warnings}

    for s in slides:
        idx = s.get("slide_index")
        if not isinstance(idx, int):
            errors.append({"kind": "bad_slide_index", "value": idx})
            continue
        if s.get("story_role") not in VALID_STORY_ROLES:
            warnings.append({"kind": "unknown_story_role",
                            "slide_index": idx, "role": s.get("story_role")})
        slot_ids = []
        for slot in s.get("narrative_slots", []):
            if slot.get("kind") not in VALID_SLOT_KINDS:
                errors.append({"kind": "bad_slot_kind",
                              "slide_index": idx, "slot_id": slot.get("slot_id"),
                              "value": slot.get("kind")})
            cap = slot.get("capacity")
            if not isinstance(cap, int) or cap < 1:
                errors.append({"kind": "bad_capacity",
                              "slide_index": idx, "slot_id": slot.get("slot_id"),
                              "value": cap})
            slot_ids.append(slot.get("slot_id"))
            # 结构一致性
            if slot.get("kind") == "single" and not slot.get("shapes"):
                errors.append({"kind": "single_without_shapes",
                              "slide_index": idx, "slot_id": slot.get("slot_id")})
            if slot.get("kind") in ("list_slot", "enumeration_slot") \
                    and not slot.get("shape_groups"):
                errors.append({"kind": "list_without_shape_groups",
                              "slide_index": idx, "slot_id": slot.get("slot_id")})
            if slot.get("kind") == "bullet_group_slot" and not slot.get("paragraphs"):
                errors.append({"kind": "bullet_without_paragraphs",
                              "slide_index": idx, "slot_id": slot.get("slot_id")})

        # slot_id 去重
        dups = [sid for sid, c in Counter(slot_ids).items() if c > 1 and sid]
        if dups:
            errors.append({"kind": "duplicate_slot_id",
                          "slide_index": idx, "duplicates": dups})

    # vision_ambiguous status 校验
    for amb in story.get("vision_ambiguous", []):
        if amb.get("status") not in ("pending", "resolved", "failed"):
            warnings.append({"kind": "bad_vision_status",
                            "shape_id": amb.get("shape_id"),
                            "status": amb.get("status")})

    return {"errors": errors, "warnings": warnings}


# ============================================================
# mapping v3 校验
# ============================================================

def _flatten_mapping(mapping: dict) -> tuple[dict, dict]:
    """见 apply_content._flatten_mapping 的同名函数。"""
    flat: dict[str, object] = {}
    meta: dict[str, dict] = {}
    if isinstance(mapping, dict) and "slides" in mapping \
            and isinstance(mapping["slides"], dict):
        for idx_str, page in mapping["slides"].items():
            for a in page.get("assignments", []):
                key = a["key"]
                flat[key] = a["value"]
                meta[key] = {
                    "slot_id": a.get("slot_id"),
                    "role": a.get("role"),
                    "source": a.get("source"),
                    "preserve_action": a.get("preserve_action"),
                    "item_index": a.get("item_index"),
                    "slide_index": int(idx_str),
                }
        return flat, meta
    for k, v in mapping.items():
        if k in ("schema_version", "pending_shrink_requests", "meta",
                 "fallback_required"):
            continue
        m = KEY_RE.match(k)
        if not m:
            continue
        flat[k] = v
        meta[k] = {"slide_index": int(m.group(1))}
    return flat, meta


def _build_story_shape_index(story: dict) -> tuple[dict[str, dict], dict[str, dict]]:
    """返回 (narrative_shape_info, non_content_shape_info)，key 为 'slide_N::shape_id'。"""
    narrative_idx: dict[str, dict] = {}
    non_content_idx: dict[str, dict] = {}
    for s in story.get("slides", []):
        i = s["slide_index"]
        for slot in s.get("narrative_slots", []):
            base = {
                "slot_id": slot["slot_id"],
                "slot_kind": slot["kind"],
                "role": slot["role"],
                "capacity": slot["capacity"],
                "capacity_is_hard": slot.get("capacity_is_hard", False),
                "item_template": slot.get("item_template", {}),
            }
            if slot["kind"] == "single":
                for sh in slot.get("shapes", []):
                    narrative_idx[f"slide_{i}::{sh['shape_id']}"] = {
                        **base, "shape_item_index": None,
                        "current_text": sh.get("current_text", ""),
                    }
            elif slot["kind"] in ("list_slot", "enumeration_slot"):
                for g in slot.get("shape_groups", []):
                    narrative_idx[f"slide_{i}::{g['shape_id']}"] = {
                        **base, "shape_item_index": g["item_index"],
                        "current_text": g.get("current_text", ""),
                    }
            elif slot["kind"] == "bullet_group_slot":
                for p in slot.get("paragraphs", []):
                    narrative_idx[f"slide_{i}::{p['shape_id']}"] = {
                        **base, "shape_item_index": p.get("paragraph_index"),
                        "current_text": p.get("current_text", ""),
                    }
        for nc in s.get("non_content_shapes", []):
            non_content_idx[f"slide_{i}::{nc['shape_id']}"] = {
                "role": nc.get("role"),
                "preserve_action": nc.get("preserve_action", "keep_original"),
            }
    return narrative_idx, non_content_idx


def validate_mapping(story: dict, mapping: dict) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []

    flat, meta = _flatten_mapping(mapping)
    narrative_idx, non_content_idx = _build_story_shape_index(story)

    # 按 slide 聚合文本：用于 duplicate_text_in_slide / list_items_identical
    slide_texts: dict[int, list[dict]] = {}  # slide -> [{key,value,slot_id,item_index}]

    for key, value in flat.items():
        m = KEY_RE.match(key)
        if not m:
            errors.append({"key": key, "kind": "bad_key_format"})
            continue
        slide_idx = int(m.group(1))

        # 1) 存在性
        if key not in narrative_idx and key not in non_content_idx:
            errors.append({"key": key, "kind": "shape_not_in_story",
                          "slide_index": slide_idx})
            continue

        if value == "__skip__":
            continue

        info = narrative_idx.get(key) or non_content_idx.get(key)
        assignment_meta = meta.get(key, {})

        # 2) non_content 被填非空普通值 → 视为违规
        if key in non_content_idx:
            nc = non_content_idx[key]
            if nc["preserve_action"] == "keep_original" and value != "__clear__":
                errors.append({
                    "key": key, "kind": "non_content_slot_filled",
                    "slide_index": slide_idx,
                    "role": nc["role"],
                    "preserve_action": nc["preserve_action"],
                    "offending": (value if isinstance(value, str)
                                 else f"list[{len(value)}]")[:40],
                })
                continue
            # clear_to_empty 的位置只允许 __clear__（或 __skip__）
            if nc["preserve_action"] == "clear_to_empty" \
                    and value not in ("__skip__", "__clear__"):
                warnings.append({
                    "key": key, "kind": "clear_expected_but_filled",
                    "slide_index": slide_idx,
                    "offending": (value if isinstance(value, str)
                                 else f"list[{len(value)}]")[:40],
                })
            continue

        # narrative slot 校验
        nar = narrative_idx[key]
        max_chars = int(nar.get("item_template", {}).get("text", {})
                       .get("max_chars", 0)) or 0

        if nar["slot_kind"] == "bullet_group_slot":
            if not isinstance(value, list):
                # 允许单字符串合成单段 bullet；但多段丢失 → 警告
                warnings.append({"key": key,
                                "kind": "bullet_expected_list",
                                "slide_index": slide_idx})
            else:
                if nar["capacity_is_hard"] and len(value) != nar["capacity"]:
                    errors.append({
                        "key": key, "kind": "bullet_count_mismatch",
                        "slide_index": slide_idx, "slot_id": nar["slot_id"],
                        "expected": nar["capacity"], "actual": len(value),
                    })
                for i, it in enumerate(value):
                    if not isinstance(it, str):
                        errors.append({"key": key, "kind": "bullet_item_not_string",
                                      "item_index": i})
                        continue
                    if max_chars and len(it) > max_chars:
                        errors.append({
                            "key": key, "kind": "text_exceeds_max_chars",
                            "slide_index": slide_idx, "slot_id": nar["slot_id"],
                            "item_index": i, "len": len(it), "max_chars": max_chars,
                        })
                    slide_texts.setdefault(slide_idx, []).append({
                        "key": key, "text": it.strip(),
                        "slot_id": nar["slot_id"], "item_index": i,
                    })
        else:
            if isinstance(value, str):
                if max_chars and len(value) > max_chars:
                    errors.append({
                        "key": key, "kind": "text_exceeds_max_chars",
                        "slide_index": slide_idx, "slot_id": nar["slot_id"],
                        "len": len(value), "max_chars": max_chars,
                    })
                slide_texts.setdefault(slide_idx, []).append({
                    "key": key, "text": value.strip(),
                    "slot_id": nar["slot_id"],
                    "item_index": assignment_meta.get("item_index"),
                })
            elif isinstance(value, list):
                errors.append({"key": key, "kind": "list_for_non_bullet_slot",
                              "slide_index": slide_idx,
                              "slot_kind": nar["slot_kind"]})

    # 3) 同页重复 / list 成员完全一致
    for slide_idx, items in slide_texts.items():
        # 同页整文重复（不同 slot 相同文本）
        # 注意：单 slot 的多 item（list）允许但不应全部一样 —— 下面单独判
        text_to_keys: dict[str, list[dict]] = {}
        for it in items:
            if not it["text"]:
                continue
            text_to_keys.setdefault(it["text"], []).append(it)
        for text, keys_info in text_to_keys.items():
            # 跨 slot 完全重复
            distinct_slots = {k["slot_id"] for k in keys_info}
            if len(keys_info) > 1 and len(distinct_slots) > 1:
                errors.append({
                    "kind": "duplicate_text_in_slide",
                    "slide_index": slide_idx,
                    "text": text[:40],
                    "keys": [k["key"] for k in keys_info[:5]],
                })

        # list_slot / enumeration_slot 内所有成员完全一致
        by_slot: dict[str, list[dict]] = {}
        for it in items:
            if it["item_index"] is None:
                continue
            by_slot.setdefault(it["slot_id"], []).append(it)
        for sid, members in by_slot.items():
            if len(members) < 2:
                continue
            texts = {m["text"] for m in members if m["text"]}
            if len(texts) <= 1 and texts != {""}:
                errors.append({
                    "kind": "list_items_identical",
                    "slide_index": slide_idx, "slot_id": sid,
                    "shared_text": next(iter(texts), "")[:40],
                    "item_count": len(members),
                })

    return {"errors": errors, "warnings": warnings}


# ============================================================
# decision（mode_b 色）
# ============================================================

def validate_decision(decision: dict) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    tc = decision.get("theme_colors")
    if not isinstance(tc, dict):
        errors.append({"kind": "missing_theme_colors"})
        return {"errors": errors, "warnings": warnings}
    for slot in THEME_SLOTS:
        if slot not in tc:
            errors.append({"kind": "missing_slot", "slot": slot})
            continue
        info = tc[slot]
        hx = info.get("hex") if isinstance(info, dict) else info
        if not isinstance(hx, str):
            errors.append({"kind": "slot_value_not_string", "slot": slot})
            continue
        try:
            normalize_hex(hx)
        except ValueError:
            errors.append({"kind": "invalid_hex", "slot": slot, "value": hx})
    if not errors:
        try:
            dk1 = tc["dk1"]["hex"] if isinstance(tc["dk1"], dict) else tc["dk1"]
            lt1 = tc["lt1"]["hex"] if isinstance(tc["lt1"], dict) else tc["lt1"]
            contrast = wcag_contrast(dk1, lt1)
            if contrast < 4.5:
                errors.append({"kind": "wcag_contrast_too_low",
                              "dk1": normalize_hex(dk1), "lt1": normalize_hex(lt1),
                              "contrast": round(contrast, 2)})
            elif contrast < 7.0:
                warnings.append({"kind": "wcag_contrast_aa_only",
                                "contrast": round(contrast, 2)})
        except Exception as e:  # noqa: BLE001
            errors.append({"kind": "contrast_check_failed", "msg": str(e)})
    return {"errors": errors, "warnings": warnings}


# ============================================================
# outline（旧方案兼容）
# ============================================================

def validate_outline(spec_or_story: dict, outline: dict) -> dict:
    errors: list[dict] = []
    warnings: list[dict] = []
    expected_n = spec_or_story.get("slide_count") \
        or spec_or_story.get("meta", {}).get("slide_count") \
        or len(spec_or_story.get("slides", []))
    slides = outline.get("slides")
    if not isinstance(slides, list):
        errors.append({"kind": "missing_slides_array"})
        return {"errors": errors, "warnings": warnings}
    if len(slides) != expected_n:
        errors.append({"kind": "slide_count_mismatch",
                      "outline_count": len(slides),
                      "template_count": expected_n})
    seen_idx: set[int] = set()
    for i, page in enumerate(slides, start=1):
        if not isinstance(page, dict):
            errors.append({"kind": "page_not_object", "position": i})
            continue
        idx = page.get("template_slide_index")
        if not isinstance(idx, int) or idx < 1 or idx > expected_n:
            errors.append({"kind": "bad_template_slide_index",
                          "position": i, "value": idx})
        elif idx in seen_idx:
            errors.append({"kind": "duplicate_template_slide_index", "value": idx})
        else:
            seen_idx.add(idx)
        role = page.get("role")
        if role not in VALID_OUTLINE_ROLES:
            errors.append({"kind": "invalid_role", "position": i, "role": role})
        briefing = page.get("briefing")
        if not isinstance(briefing, str) or not briefing.strip():
            errors.append({"kind": "missing_briefing", "position": i})
    missing_idx = sorted(set(range(1, expected_n + 1)) - seen_idx)
    if missing_idx:
        errors.append({"kind": "uncovered_template_slides", "indices": missing_idx})
    return {"errors": errors, "warnings": warnings}


# ============================================================
# CLI
# ============================================================

def _emit(result: dict, strict: bool) -> int:
    if strict:
        result["errors"].extend(result.pop("warnings", []))
        result["warnings"] = []
    payload = {
        "ok": not result["errors"],
        "error_count": len(result["errors"]),
        "warning_count": len(result["warnings"]),
        "errors": result["errors"],
        "warnings": result["warnings"],
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload["ok"] else 2


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="kind", required=True)

    ps = sub.add_parser("story")
    ps.add_argument("--story", required=True, type=Path)
    ps.add_argument("--strict", action="store_true")

    pb = sub.add_parser("blueprint")
    pb.add_argument("--story", required=True, type=Path)
    pb.add_argument("--blueprint", required=True, type=Path)
    pb.add_argument("--strict", action="store_true")

    pm = sub.add_parser("mapping")
    pm.add_argument("--story", required=True, type=Path)
    pm.add_argument("--mapping", required=True, type=Path)
    pm.add_argument("--strict", action="store_true")

    pd = sub.add_parser("decision")
    pd.add_argument("--decision", required=True, type=Path)
    pd.add_argument("--strict", action="store_true")

    po = sub.add_parser("outline")
    po.add_argument("--spec", required=True, type=Path,
                   help="template_spec.json 或 template_story.json")
    po.add_argument("--outline", required=True, type=Path)
    po.add_argument("--strict", action="store_true")

    args = ap.parse_args()

    try:
        if args.kind == "story":
            story = json.loads(args.story.read_text(encoding="utf-8"))
            return _emit(validate_story(story), args.strict)
        if args.kind == "blueprint":
            story = json.loads(args.story.read_text(encoding="utf-8"))
            bp = json.loads(args.blueprint.read_text(encoding="utf-8"))
            return _emit(_validate_bp(story, bp), args.strict)
        if args.kind == "mapping":
            story = json.loads(args.story.read_text(encoding="utf-8"))
            mapping = json.loads(args.mapping.read_text(encoding="utf-8"))
            return _emit(validate_mapping(story, mapping), args.strict)
        if args.kind == "decision":
            decision = json.loads(args.decision.read_text(encoding="utf-8"))
            return _emit(validate_decision(decision), args.strict)
        if args.kind == "outline":
            spec = json.loads(args.spec.read_text(encoding="utf-8"))
            outline = json.loads(args.outline.read_text(encoding="utf-8"))
            return _emit(validate_outline(spec, outline), args.strict)
    except FileNotFoundError as e:
        print(json.dumps({"ok": False, "error": f"file not found: {e.filename}"}))
        return 1
    except json.JSONDecodeError as e:
        print(json.dumps({"ok": False, "error": f"invalid json: {e}"}))
        return 1


if __name__ == "__main__":
    sys.exit(main())
