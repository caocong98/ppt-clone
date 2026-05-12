# -*- coding: utf-8 -*-
"""将「南京旅游」类模板克隆为「武汉旅游」正文（mode_a 保持原主题色）。

用法：
  python scripts/clone_nanjing_to_wuhan_driver.py <源.pptx>
  python scripts/clone_nanjing_to_wuhan_driver.py <源.pptx> --regenerate-images
  python scripts/clone_nanjing_to_wuhan_driver.py <源.pptx> --regenerate-images --image-provider openai

交付：<源同目录>/<源stem>_clone.pptx
可选配图：<源同目录>/<源stem>_clone_img.pptx（需 --regenerate-images）
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

WUHAN_THEME: dict = {
    "title": "武汉江城五一旅游攻略计划",
    "thesis": (
        "以黄鹤楼、东湖绿道与汉口江滩为主线，串联过早文化与地铁出行，"
        "安排轻松高效的武汉五一行程与实用贴士。"
    ),
    "seasonal_hint": "劳动节五一",
    "unit": "江城旅记",
    "tone": "warm_informative",
    "vocab": {
        "nouns": [
            "黄鹤楼", "东湖", "江滩", "光谷", "户部巷", "武大", "长江",
            "轮渡", "地铁", "热干面", "豆皮", "糊汤粉", "吉庆街",
            "楚河汉街", "省博", "昙华林", "黎黄陂路", "汉阳", "武昌",
            "汉口", "绿道", "夜景", "江风", "知音号", "大桥", "粮道街",
        ],
        "verbs_short": [
            "探访", "漫步", "品尝", "乘坐", "打卡", "安排", "预留",
            "错峰", "收藏", "体验", "骑行", "夜游",
        ],
        "bullets": [
            "粮道街过早热干面配蛋酒，暖胃顶饱开启一天",
            "黄鹤楼远眺长江大桥，讲解与排队请预留两小时",
            "东湖绿道骑行或游船，城市绿心放松半日刚好",
            "汉口江滩散步看灯光，夜景观景别错过轮渡班次",
            "省博编钟与越王剑展，提前预约可少排队",
            "光谷步行街傍晚逛街小吃，地铁直达适合补给",
            "楚河汉街购物观景一站式，雨天可作室内备选",
            "户部巷吉庆街尝市井小吃，注意错峰与人潮",
        ],
    },
    "titles": [
        "武汉江城五一旅游攻略计划",
        "行程总览",
        "经典名片",
        "东湖游玩",
        "江滩夜色",
        "美食过早",
        "交通住宿",
        "天气穿着",
        "预算提示",
        "安全贴士",
        "祝旅途愉快",
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
    ap = argparse.ArgumentParser(description="南京旅游模板 → 武汉旅游克隆")
    ap.add_argument("src", type=Path, help="源 .pptx")
    ap.add_argument(
        "--regenerate-images",
        action="store_true",
        help="克隆后按 blueprint 为 decorative_picture 重新生成配图",
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

    themed = None
    with_text = None
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
                        "reason": "clone_wuhan_driver: 按可替换正文处理",
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

        print("[4/7] generate_blueprint（武汉主题）…")
        bp_doc = generate_blueprint(story, WUHAN_THEME)
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
