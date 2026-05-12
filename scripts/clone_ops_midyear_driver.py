# -*- coding: utf-8 -*-
"""克隆商务类总结模板 →「运营岗位年中总结汇报」正文（mode_a 保持原主题色）。

用法：
  python scripts/clone_ops_midyear_driver.py <源.pptx>
  python scripts/clone_ops_midyear_driver.py <源.pptx> --regenerate-images

交付：<源同目录>/<源stem>_clone.pptx
可选配图：<源stem>_clone_img.pptx
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

from _workspace import bundle_workspace  # noqa: E402
import parse_template_story as pts  # noqa: E402
from parse_template_story import resume_with_vision  # noqa: E402
import build_content_blueprint as bcb  # noqa: E402
from batch_clone_four import generate_blueprint  # noqa: E402


OPS_MIDYEAR_THEME: dict = {
    "title": "运营岗位年中总结汇报",
    "thesis": (
        "围绕用户增长与活跃、内容与活动闭环、转化与成本控制，复盘上半年打法，明确下半年重点。"
    ),
    "seasonal_hint": "年中总结",
    "unit": "运营部",
    "tone": "concise_data_driven",
    "vocab": {
        "nouns": [
            "DAU", "MAU", "留存", "转化", "GMV", "ROI", "CAC", "LTV",
            "活动", "内容", "社群", "裂变", "拉新", "促活", "复购",
            "漏斗", "触达", "曝光", "点击", "完播", "客诉", "NPS",
            "渠道", "投放", "自然量", "私域", "会员", "补贴", "券",
            "排期", "复盘", "AB测", "数据看板",
        ],
        "verbs_short": [
            "提升", "优化", "落地", "沉淀", "对齐", "拉通", "聚焦",
            "放大", "降本", "提效", "迭代", "验证", "闭环",
        ],
        "bullets": [
            "上半年 GMV 同比提升，ROI 稳定在目标区间上部",
            "重点活动三场破圈引流，私信与表单转化达标",
            "内容矩阵日均曝光增长，长尾话题带动自然搜索",
            "社群周活与用户留存稳中有升，客诉 SLA 全流程达标",
            "投放与自然量配比优化，单客获客成本环比下降",
            "会员与复购链路打通两段关键卡点，漏斗中段改善明显",
            "数据看板与周会机制固化，节奏与归因更清晰",
            "下半年将围绕核心品类与重点区域做资源加权与试点",
        ],
    },
    "titles": [
        "运营岗位年中总结汇报",
        "整体概览",
        "指标达成",
        "增长与获客",
        "内容与活动",
        "用户留存与社群",
        "转化与收入",
        "成本与效率",
        "亮点与复盘",
        "风险与对策",
        "下半年计划",
        "感谢与致谢",
    ],
}


def _run(cmd: list, *, ok_codes: tuple[int, ...] = (0,)) -> subprocess.CompletedProcess:
    r = subprocess.run(
        [str(c) for c in cmd],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if r.returncode not in ok_codes:
        tail = (r.stderr or r.stdout or "")[-1200:]
        raise RuntimeError(f"cmd failed rc={r.returncode}: {cmd}\n{tail}")
    return r


def main() -> int:
    ap = argparse.ArgumentParser(description="商务总结模板 → 运营年中汇报克隆")
    ap.add_argument("src", type=Path, help="源 .pptx")
    ap.add_argument(
        "--regenerate-images",
        action="store_true",
        help="克隆后为 decorative_picture 重新生成配图",
    )
    ap.add_argument(
        "--image-provider",
        choices=("pil", "openai"),
        default="pil",
        help="pil=离线示意；openai=需 OPENAI_API_KEY",
    )
    args = ap.parse_args()
    src = args.src.resolve()
    if not src.is_file():
        print(f"源文件不存在: {src}", file=sys.stderr)
        return 1

    img_out: Path | None = None
    with bundle_workspace(src) as bp:
        inter = bp.intermediate
        story_path = inter / "template_story.json"
        themed = inter / "themed.pptx"
        bp_path = bp.scratch / "blueprint.json"
        mapping_path = inter / "content_mapping.json"
        with_text = inter / "with_text.pptx"

        print("[1/7] analyze_template …")
        _run(
            [
                sys.executable,
                SCRIPTS / "analyze_template.py",
                str(src),
                "--out",
                str(inter / "analyze.json"),
            ]
        )

        print("[2/7] parse_template_story …")
        thresholds = dict(pts.DEFAULT_THRESHOLDS)
        story = pts.parse(src, thresholds=thresholds, logo_action="keep_original")
        story_path.write_text(
            json.dumps(story, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        amb = story.get("vision_ambiguous") or []
        if amb:
            print(f"    vision_ambiguous={len(amb)} → 默认按正文槽位回灌")
            vision_payload = {
                "results": [
                    {
                        "shape_id": a["shape_id"],
                        "slide_index": a["slide_index"],
                        "role": "content_slot",
                        "preserve_action": None,
                        "confidence": 0.88,
                        "reason": "clone_ops_midyear_driver: 按可替换正文处理",
                    }
                    for a in amb
                ]
            }
            (bp.scratch / "vision_result.json").write_text(
                json.dumps(vision_payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            story = resume_with_vision(story, vision_payload)
            story_path.write_text(
                json.dumps(story, ensure_ascii=False, indent=2), encoding="utf-8"
            )

        print("[3/7] mode_a 复制主题（不改色）…")
        shutil.copyfile(src, themed)

        print("[4/7] generate_blueprint（运营年中主题）…")
        bp_doc = generate_blueprint(story, OPS_MIDYEAR_THEME)
        bp_path.write_text(
            json.dumps(bp_doc, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        val = bcb.validate_blueprint(story, bp_doc)
        err_n, warn_n = len(val["errors"]), len(val["warnings"])
        print(f"    validate errors={err_n} warnings={warn_n}")
        if val["errors"]:
            for e in val["errors"][:8]:
                print(
                    f"      ERR {e.get('kind')} "
                    f"slide={e.get('slide_index')} {e.get('shape_id')}: "
                    f"{e.get('msg', '')[:80]}"
                )
            raise RuntimeError("blueprint 校验失败，中止")

        print("[5/7] map_blueprint_to_template …")
        _run(
            [
                sys.executable,
                SCRIPTS / "map_blueprint_to_template.py",
                "--story",
                str(story_path),
                "--blueprint",
                str(bp_path),
                "--out",
                str(mapping_path),
            ]
        )

        print("[6/7] apply_content …")
        _run(
            [
                sys.executable,
                SCRIPTS / "apply_content.py",
                str(themed),
                "--mapping",
                str(mapping_path),
                "--story",
                str(story_path),
                "--out",
                str(with_text),
            ]
        )

        print("[7/7] 落盘最终 pptx …")
        shutil.copyfile(with_text, bp.final_pptx)
        print(f"完成: {bp.final_pptx}")

        if args.regenerate_images:
            img_out = bp.final_pptx.parent / f"{bp.final_pptx.stem}_img.pptx"
            print(f"[8] regenerate_slide_images → {img_out.name} …")
            _run(
                [
                    sys.executable,
                    SCRIPTS / "regenerate_slide_images.py",
                    "--pptx",
                    str(bp.final_pptx),
                    "--story",
                    str(story_path),
                    "--blueprint",
                    str(bp_path),
                    "--provider",
                    args.image_provider,
                    "--out",
                    str(img_out),
                ]
            )

        lint_out = bp.report / "lint.json"
        _run(
            [
                sys.executable,
                SCRIPTS / "lint_pptx.py",
                str(bp.final_pptx),
                "--story",
                str(story_path),
                "--out",
                str(lint_out),
            ],
            ok_codes=(0, 2),
        )
        try:
            lint = json.loads(lint_out.read_text(encoding="utf-8"))
            print(f"lint findings={len(lint.get('findings', []))}")
        except OSError:
            pass

        _run(
            [
                sys.executable,
                SCRIPTS / "verify_effect.py",
                "--before",
                str(src),
                "--after",
                str(bp.final_pptx),
                "--engine",
                "powerpoint_com",
                "--out-dir",
                str(bp.verify),
            ],
            ok_codes=(0, 1, 2),
        )

    clone_path = src.parent / f"{src.stem}_clone.pptx"
    msg = f"工作区已清理。请打开: {clone_path}"
    if img_out is not None:
        msg += f"\n配图版: {img_out}"
    print(msg)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"FAIL: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
