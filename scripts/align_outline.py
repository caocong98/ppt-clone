"""[DEPRECATED] 旧的 outline_plan 脚手架。

请改用：
  python build_content_blueprint.py scaffold --story template_story.json \\
      [--outline outline.txt] [--topic-json topic.json] --out blueprint_scaffold.json

新流程用 template_story.json + content_blueprint.json 代替 template_spec.json +
outline_plan.json；blueprint 由 Prompt C 直接按 story 骨架填写 beats，不再需要先
写 briefing 再压缩。

本脚本保留仅为向后兼容；执行时会打印 deprecation 警告。"""

输出 outline_plan.json：
{
  "schema_version": 1,
  "slides": [
    {
      "template_slide_index": 1,
      "role": "cover",
      "title_hint": "原标题",
      "placeholders_brief": [
        {"shape_id": "矩形 7", "role": "title", "max_chars": 10},
        {"shape_id": "矩形 9", "role": "subtitle", "max_chars": 23}
      ],
      "briefing": ""
    },
    ...
  ]
}

CLI:
  python align_outline.py scaffold --spec template_spec.json --out outline_plan.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def scaffold(spec: dict) -> dict:
    slides_out = []
    for s in spec.get("slides", []):
        title_hint = ""
        for ph in s.get("placeholders", []):
            if ph.get("role") == "title":
                title_hint = ph.get("current_text", "") or title_hint
                break
        if not title_hint and s.get("placeholders"):
            title_hint = s["placeholders"][0].get("current_text", "")

        ph_brief = []
        for ph in s.get("placeholders", []):
            ph_brief.append({
                "shape_id": ph["shape_id"],
                "role": ph.get("role", "body"),
                "max_chars": int(ph.get("max_chars", 0)),
                "max_bullets": int(ph.get("max_bullets", 1)),
                "max_chars_per_bullet": int(ph.get("max_chars_per_bullet", ph.get("max_chars", 0))),
            })

        slides_out.append({
            "template_slide_index": s["index"],
            "role": s.get("heuristic_role", "content"),
            "title_hint": title_hint,
            "placeholders_brief": ph_brief,
            "briefing": "",
        })

    return {"schema_version": 1, "slides": slides_out}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="action", required=True)

    sc = sub.add_parser("scaffold", help="基于 template_spec 生成 outline_plan 脚手架")
    sc.add_argument("--spec", required=True, type=Path)
    sc.add_argument("--out", required=True, type=Path)

    sub.add_parser("validate-help", help="校验请使用 validate.py outline ...")

    args = ap.parse_args()

    print("[DEPRECATED] align_outline.py 已被 build_content_blueprint.py 取代；"
          "本脚本仅做向后兼容，未来版本将移除。",
          file=sys.stderr)

    if args.action == "scaffold":
        spec = json.loads(args.spec.read_text(encoding="utf-8"))
        plan = scaffold(spec)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"ok": True, "out": str(args.out),
                          "slide_count": len(plan["slides"])}, ensure_ascii=False))
        return 0

    if args.action == "validate-help":
        print("请使用：python validate.py outline --spec ... --outline ...")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
