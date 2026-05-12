"""闭环验证：渲染 before/after 两份 PPT，并拼出对比图供 Agent 用多模态判读。

CLI:
    python verify_effect.py --before <pptx> --after <pptx> --out <dir> [--width 800] [--slides 1,3,5]

输出：
    <out>/before/slide_001.png ...
    <out>/after/slide_001.png ...
    <out>/diff/slide_001.png   左右拼图（左=before, 右=after）
    <out>/verify_meta.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
from render_slides import render  # noqa: E402


def _pixel_diff_stats(a_png: Path, b_png: Path, idx: int) -> dict:
    """计算两张 PNG 的 mean-abs 像素差与高变化区域占比（resize 到 256 宽以加速）。"""
    import numpy as np
    target_w = 256
    a = Image.open(a_png).convert("RGB")
    b = Image.open(b_png).convert("RGB")
    a = a.resize((target_w, int(a.height * target_w / a.width)))
    b = b.resize((target_w, int(b.height * target_w / b.width)))
    if a.size != b.size:
        b = b.resize(a.size)
    arr_a = np.asarray(a, dtype=np.int16)
    arr_b = np.asarray(b, dtype=np.int16)
    diff = np.abs(arr_a - arr_b)
    mean_abs = float(diff.mean())
    big_pixel_ratio = float((diff.mean(axis=2) > 40).mean())
    return {"index": idx, "mean_abs": round(mean_abs, 2),
            "big_change_ratio": round(big_pixel_ratio, 3)}


# ============================================================
# L5.3 装饰 shape bbox 级 pixel diff
# ============================================================

_DECORATION_ROLES = {"decoration_number", "style_tag", "logo_text",
                     "decoration_micro", "logo_image", "time_marker"}


def _decoration_bbox_diff(story: dict, before_pngs: list[Path],
                          after_pngs: list[Path],
                          *, mean_abs_threshold: float = 18.0,
                          big_change_threshold: float = 0.25) -> list[dict]:
    """对 story 中 role 为装饰类 + preserve_action=keep_original 的 shape，
    截取 bbox 区域做 pixel diff；mean_abs 超阈值 -> 视觉回归。

    仅做粗略切片对齐（按 bbox 百分比映射到图像像素）；渲染引擎一致时足够稳定。
    """
    import numpy as np

    reports: list[dict] = []
    slide_map = {s["slide_index"]: s for s in story.get("slides", [])}
    for idx, (b, a) in enumerate(zip(before_pngs, after_pngs), start=1):
        slide = slide_map.get(idx)
        if not slide:
            continue
        try:
            img_before = Image.open(b).convert("RGB")
            img_after = Image.open(a).convert("RGB")
            if img_before.size != img_after.size:
                img_after = img_after.resize(img_before.size)
        except Exception:
            continue
        W, H = img_before.size
        arr_b = np.asarray(img_before, dtype=np.int16)
        arr_a = np.asarray(img_after, dtype=np.int16)

        for nc in slide.get("non_content_shapes", []):
            role = nc.get("role") or ""
            if role not in _DECORATION_ROLES:
                continue
            if nc.get("preserve_action") != "keep_original":
                continue
            bbox = nc.get("bbox") or {}
            if not bbox:
                continue
            x = int(max(0, bbox.get("left", 0) / 100 * W))
            y = int(max(0, bbox.get("top", 0) / 100 * H))
            w = int(min(W - x, bbox.get("w", 0) / 100 * W))
            h = int(min(H - y, bbox.get("h", 0) / 100 * H))
            if w <= 2 or h <= 2:
                continue
            crop_b = arr_b[y:y + h, x:x + w]
            crop_a = arr_a[y:y + h, x:x + w]
            if crop_b.size == 0 or crop_a.size == 0:
                continue
            diff = np.abs(crop_b - crop_a)
            mean_abs = float(diff.mean())
            big_ratio = float((diff.mean(axis=2) > 40).mean())
            flag = "stable"
            if mean_abs >= mean_abs_threshold or big_ratio >= big_change_threshold:
                flag = "visual_regression"
            reports.append({
                "slide": idx,
                "shape_id": nc.get("shape_id"),
                "role": role,
                "mean_abs": round(mean_abs, 2),
                "big_change_ratio": round(big_ratio, 3),
                "flag": flag,
            })
    return reports


def _side_by_side(left_png: Path, right_png: Path, out_png: Path) -> None:
    a = Image.open(left_png).convert("RGB")
    b = Image.open(right_png).convert("RGB")
    h = max(a.height, b.height)
    if a.height != h:
        a = a.resize((int(a.width * h / a.height), h))
    if b.height != h:
        b = b.resize((int(b.width * h / b.height), h))
    gap = 16
    canvas = Image.new("RGB", (a.width + gap + b.width, h + 32), (255, 255, 255))
    canvas.paste(a, (0, 32))
    canvas.paste(b, (a.width + gap, 32))
    draw = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except Exception:  # noqa: BLE001
        font = ImageFont.load_default()
    draw.text((10, 6), "BEFORE", fill=(80, 80, 80), font=font)
    draw.text((a.width + gap + 10, 6), "AFTER", fill=(80, 80, 80), font=font)
    out_png.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out_png, "PNG")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--before", required=True, type=Path)
    ap.add_argument("--after", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--width", type=int, default=800)
    ap.add_argument("--slides", type=str, default="",
                    help="只对指定页生成 diff，逗号分隔；空表示全部")
    ap.add_argument("--engine", choices=["auto", "libreoffice", "powerpoint_com"],
                    default="auto", help="渲染引擎，必须 before/after 一致")
    ap.add_argument("--pixel-diff", action="store_true",
                    help="额外计算每页 PNG 的 mean-abs 差，写入 verify_meta.json")
    ap.add_argument("--story", type=Path, default=None,
                    help="可选 template_story.json；配合 --decoration-bbox-diff 使用")
    ap.add_argument("--decoration-bbox-diff", action="store_true",
                    help="按 story 中装饰 shape 的 bbox 做像素级 diff（需要 --story）")
    args = ap.parse_args()

    if not args.before.exists() or not args.after.exists():
        print("ERROR: before/after pptx not found", file=sys.stderr)
        return 1

    before_dir = args.out / "before"
    after_dir = args.out / "after"
    diff_dir = args.out / "diff"
    before_dir.mkdir(parents=True, exist_ok=True)
    after_dir.mkdir(parents=True, exist_ok=True)
    diff_dir.mkdir(parents=True, exist_ok=True)

    before_pngs, used_b = render(args.before, before_dir, args.width, engine=args.engine)
    after_pngs, used_a = render(args.after, after_dir, args.width, engine=args.engine)
    if used_b != used_a:
        print(f"WARN: before/after 引擎不一致 ({used_b} vs {used_a})", file=sys.stderr)

    n = min(len(before_pngs), len(after_pngs))
    selected = list(range(1, n + 1))
    if args.slides.strip():
        try:
            selected = [int(x) for x in args.slides.split(",") if x.strip()]
        except ValueError:
            pass
    selected = [s for s in selected if 1 <= s <= n]

    diff_paths: list[Path] = []
    pixel_stats: list[dict] = []
    for idx in selected:
        out_png = diff_dir / f"slide_{idx:03d}.png"
        _side_by_side(before_pngs[idx - 1], after_pngs[idx - 1], out_png)
        diff_paths.append(out_png)
        if args.pixel_diff:
            pixel_stats.append(_pixel_diff_stats(before_pngs[idx - 1], after_pngs[idx - 1], idx))

    if args.pixel_diff and pixel_stats:
        # 标注「变化过大/过小」便于 Agent 重点看
        means = [s["mean_abs"] for s in pixel_stats]
        mn, mx = min(means), max(means)
        for s in pixel_stats:
            if s["mean_abs"] >= max(mx * 0.85, 25):
                s["flag"] = "high_change"
            elif s["mean_abs"] <= max(mn * 1.2, 2):
                s["flag"] = "near_zero_change"
            else:
                s["flag"] = "normal"

    decoration_bbox_reports: list[dict] | None = None
    if args.decoration_bbox_diff:
        if args.story is None or not args.story.exists():
            print("WARN: --decoration-bbox-diff 需要 --story 指向合法 template_story.json，已跳过",
                  file=sys.stderr)
        else:
            try:
                story = json.loads(args.story.read_text(encoding="utf-8"))
                decoration_bbox_reports = _decoration_bbox_diff(
                    story, before_pngs, after_pngs)
            except Exception as exc:  # noqa: BLE001
                print(f"WARN: decoration_bbox_diff 失败: {exc}", file=sys.stderr)

    meta = {
        "before_pptx": str(args.before),
        "after_pptx": str(args.after),
        "engine_used_before": used_b,
        "engine_used_after": used_a,
        "engine_consistent": used_b == used_a,
        "slide_count_before": len(before_pngs),
        "slide_count_after": len(after_pngs),
        "diffs": [{"index": s, "path": str(diff_dir / f"slide_{s:03d}.png")} for s in selected],
        "pixel_diff": pixel_stats if args.pixel_diff else None,
        "decoration_bbox_diff": decoration_bbox_reports,
    }
    (args.out / "verify_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({
        "ok": True, "out": str(args.out),
        "diff_count": len(diff_paths),
        "first_diff": str(diff_paths[0]) if diff_paths else None,
    }))
    return 0


if __name__ == "__main__":
    sys.exit(main())
