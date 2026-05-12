# -*- coding: utf-8 -*-
"""按幻灯片文案为「装饰性图片」槽位重新生成位图并写回 pptx（不改 chart / SmartArt / logo）。

依赖：
  - python-pptx, Pillow
  - 可选：环境变量 OPENAI_API_KEY + --provider openai，调用 DALL·E 3 生成配图

典型用法（克隆得到 story + blueprint 之后）：

  python scripts/regenerate_slide_images.py \\
      --pptx \"..._clone.pptx\" \\
      --story template_story.json \\
      --blueprint blueprint.json \\
      --provider pil \\
      --out \"..._clone_img.pptx\"

仅 dry-run 查看将处理的图片槽位：

  python scripts/regenerate_slide_images.py --pptx x.pptx --story s.json --dry-run
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import shutil
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.image import ImagePart

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

from apply_content import _build_shape_index  # noqa: E402


def _emu_to_px(emu: int, dpi: int = 96) -> int:
    return max(1, int(round(emu / 914400 * dpi)))


def _slide_context_from_blueprint(blueprint: dict | None, slide_index: int) -> str:
    if not blueprint:
        return ""
    for s in blueprint.get("slides", []):
        if int(s.get("slide_index", -1)) != slide_index:
            continue
        parts: list[str] = []
        pt = (s.get("page_theme") or "").strip()
        if pt:
            parts.append(pt)
        st = s.get("shape_texts") or {}
        if isinstance(st, dict):
            vals = list(st.values())[:6]
            for v in vals:
                if isinstance(v, str) and v.strip() and v != "__preserve__":
                    parts.append(v.strip()[:40])
                elif isinstance(v, list):
                    for t in v[:2]:
                        if isinstance(t, str) and t.strip():
                            parts.append(t.strip()[:40])
        return "；".join(parts)[:500]
    return ""


def _build_image_prompt(
    *,
    slide_index: int,
    story: dict,
    blueprint: dict | None,
    extra_hint: str,
) -> str:
    ctx = _slide_context_from_blueprint(blueprint, slide_index)
    if not ctx:
        # 无 blueprint 时用该页 story_role + 装饰图旁白兜底
        for s in story.get("slides", []):
            if int(s.get("slide_index", -1)) == slide_index:
                ctx = (s.get("story_role") or "content") + " 页配图"
                break
    base = (
        f"PPT slide illustration, travel / city scenery, mood matches: {ctx}. "
        f"{extra_hint} "
        "Photorealistic or high-quality editorial photo style, wide composition, "
        "suitable as presentation background art. No text, no letters, no watermark, no logo."
    )
    return base[:3500]


def _encode_for_part(pil_img: Any, ext: str) -> bytes:
    from PIL import Image as PILImage

    buf = io.BytesIO()
    e = ext.lower().strip(".")
    if e in ("jpg", "jpeg"):
        pil_img.convert("RGB").save(buf, format="JPEG", quality=90)
    else:
        pil_img.save(buf, format="PNG")
    return buf.getvalue()


def _pil_generate(
    *,
    width_px: int,
    height_px: int,
    label: str,
    slide_index: int,
    seed: int,
) -> bytes:
    """无需外网：按文案生成简单渐变 + 标题条（便于离线验收管线）。"""
    from PIL import Image as PILImage, ImageDraw, ImageFont

    w = max(64, min(width_px, 2048))
    h = max(64, min(height_px, 2048))
    img = PILImage.new("RGB", (w, h))
    draw = ImageDraw.Draw(img)
    # 渐变（按 seed 变色相）
    r0, g0, b0 = 30 + (seed * 17) % 80, 60 + (seed * 31) % 120, 90 + (seed * 7) % 100
    r1, g1, b1 = 180 + seed % 60, 120 + (seed * 3) % 80, 80
    for y in range(h):
        t = y / max(h - 1, 1)
        r = int(r0 + (r1 - r0) * t)
        g = int(g0 + (g1 - g0) * t)
        b = int(b0 + (b1 - b0) * t)
        draw.line([(0, y), (w, y)], fill=(r, g, b))
    text = (label or f"slide {slide_index}")[:60]
    font = None
    for fp in (
        os.environ.get("PPT_CLONE_FONT", ""),
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ):
        if not fp:
            continue
        if Path(fp).is_file():
            try:
                font = ImageFont.truetype(fp, max(14, min(w, h) // 18))
                break
            except OSError:
                continue
    if font is None:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = max(8, (w - tw) // 2)
    ty = max(8, (h - th) // 2)
    pad = 6
    draw.rectangle(
        [tx - pad, ty - pad, tx + tw + pad, ty + th + pad],
        fill=(24, 24, 28),
    )
    draw.text((tx, ty), text, fill=(255, 255, 255), font=font)
    return _encode_for_part(img, "png")


def _openai_generate(
    *,
    prompt: str,
    model: str,
    api_key: str,
    timeout_s: int = 120,
) -> bytes:
    """返回 PNG/JPEG bytes（以 API 返回格式为准，后续统一转 PIL）。"""
    from PIL import Image as PILImage

    body = json.dumps(
        {
            "model": model,
            "prompt": prompt[:4000],
            "n": 1,
            "size": "1024x1024",
            "response_format": "b64_json",
        },
        ensure_ascii=False,
    ).encode("utf-8")
    req = urllib.request.Request(
        "https://api.openai.com/v1/images/generations",
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            payload = json.loads(resp.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI HTTP {e.code}: {err[:800]}") from e
    data = payload.get("data") or []
    if not data or "b64_json" not in data[0]:
        raise RuntimeError(f"OpenAI 响应异常: {str(payload)[:500]}")
    raw = base64.b64decode(data[0]["b64_json"])
    pil_img = PILImage.open(io.BytesIO(raw)).convert("RGB")
    return _encode_for_part(pil_img, "png")


def _resize_cover(pil_img: Any, target_w: int, target_h: int) -> Any:
    from PIL import Image as PILImage

    img = pil_img.convert("RGB")
    tw, th = max(1, target_w), max(1, target_h)
    src_w, src_h = img.size
    scale = max(tw / src_w, th / src_h)
    nw, nh = int(src_w * scale), int(src_h * scale)
    img = img.resize((nw, nh), PILImage.Resampling.LANCZOS)
    left = max(0, (nw - tw) // 2)
    top = max(0, (nh - th) // 2)
    return img.crop((left, top, left + tw, top + th))


def _replace_image_part(shape: Any, new_bytes: bytes) -> None:
    rId = shape._pic.blip_rId
    if not rId:
        raise ValueError("picture 无嵌入 blip（可能为链接图），跳过")
    part = shape.part.related_part(rId)
    if not isinstance(part, ImagePart):
        raise TypeError(f"期望 ImagePart，得到 {type(part)}")
    ext = part.partname.ext or ".png"
    from PIL import Image as PILImage

    pil = PILImage.open(io.BytesIO(new_bytes))
    # 按占位框比例 center-crop 缩放
    tw = _emu_to_px(int(shape.width))
    th = _emu_to_px(int(shape.height))
    fitted = _resize_cover(pil, tw, th)
    part.blob = _encode_for_part(fitted, ext)


def _collect_targets(story: dict) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for s in story.get("slides", []):
        idx = int(s["slide_index"])
        for nc in s.get("non_content_shapes", []):
            if nc.get("role") != "decorative_picture":
                continue
            rows.append(
                {
                    "shape_id": nc["shape_id"],
                    "slide_index": idx,
                    "hint": (nc.get("current_text") or "")[:80],
                }
            )
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--pptx", type=Path, required=True)
    ap.add_argument("--story", type=Path, required=True, help="template_story.json（含 non_content_shapes）")
    ap.add_argument("--blueprint", type=Path, default=None, help="content_blueprint.json，用于拼配图 prompt")
    ap.add_argument("--out", type=Path, default=None, help="默认 <pptx 同目录>/<stem>_img.pptx")
    ap.add_argument(
        "--provider",
        choices=("pil", "openai"),
        default="pil",
        help="pil=离线渐变示意；openai=需 OPENAI_API_KEY",
    )
    ap.add_argument("--openai-model", default="dall-e-3")
    ap.add_argument("--extra-prompt", default="", help="附加英文/中文风格提示")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    pptx = args.pptx.resolve()
    story_path = args.story.resolve()
    if not pptx.is_file():
        print(f"ERROR: pptx 不存在: {pptx}", file=sys.stderr)
        return 1
    story = json.loads(story_path.read_text(encoding="utf-8"))
    blueprint: dict | None = None
    if args.blueprint and args.blueprint.is_file():
        blueprint = json.loads(args.blueprint.read_text(encoding="utf-8"))

    targets = _collect_targets(story)
    out = args.out
    if out is None:
        out = pptx.parent / f"{pptx.stem}_img.pptx"
    else:
        out = out.resolve()

    print(
        json.dumps(
            {
                "targets": len(targets),
                "out": str(out),
                "provider": args.provider,
            },
            ensure_ascii=False,
        )
    )
    if args.dry_run:
        for t in targets[:50]:
            print(json.dumps(t, ensure_ascii=False))
        if len(targets) > 50:
            print(f"... 其余 {len(targets) - 50} 条省略")
        return 0

    shutil.copyfile(pptx, out)
    prs = Presentation(str(out))
    idx = _build_shape_index(prs)

    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if args.provider == "openai" and not api_key:
        print("ERROR: openai 模式需要环境变量 OPENAI_API_KEY", file=sys.stderr)
        return 1

    done_r: set[str] = set()
    ok = 0
    err = 0
    for i, t in enumerate(targets):
        sid = t["shape_id"]
        slide_i = t["slide_index"]
        shape = idx.get(sid)
        if shape is None:
            print(f"[skip] 找不到 shape: {sid}")
            err += 1
            continue
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            print(f"[skip] 非 PICTURE: {sid} type={shape.shape_type}")
            err += 1
            continue
        try:
            r_id = shape._pic.blip_rId
        except Exception:  # noqa: BLE001
            r_id = None
        if not r_id:
            print(f"[skip] 无嵌入图: {sid}")
            err += 1
            continue
        if r_id in done_r:
            print(f"[share] 同资源已更新 rId={r_id}，跳过重复生成: {sid}")
            continue

        prompt = _build_image_prompt(
            slide_index=slide_i,
            story=story,
            blueprint=blueprint,
            extra_hint=args.extra_prompt,
        )
        tw = _emu_to_px(int(shape.width))
        th = _emu_to_px(int(shape.height))

        if args.provider == "openai":
            raw = _openai_generate(
                prompt=prompt, model=args.openai_model, api_key=api_key
            )
            time.sleep(1.5)
        else:
            label = _slide_context_from_blueprint(blueprint, slide_i) or t.get("hint", "")
            raw = _pil_generate(
                width_px=tw,
                height_px=th,
                label=label,
                slide_index=slide_i,
                seed=slide_i * 10007 + i,
            )

        try:
            _replace_image_part(shape, raw)
            done_r.add(r_id)
            ok += 1
            print(f"[ok] {sid} slide={slide_i} rId={r_id}")
        except Exception as ex:  # noqa: BLE001
            err += 1
            print(f"[fail] {sid}: {ex}")

    prs.save(str(out))
    print(json.dumps({"written": str(out), "ok": ok, "errors": err}, ensure_ascii=False))
    return 0 if err == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
