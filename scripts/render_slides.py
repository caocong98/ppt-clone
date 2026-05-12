"""把 PPT 每一页渲染为 PNG，可选叠加图片区域 mask。

后端：
- 优先 LibreOffice：先转 PDF，再用 pypdfium2 渲染每页（最稳定，跨平台）
- 备选 PowerPoint COM（仅 Windows，Slide.Export 直接出 PNG）

CLI:
    python render_slides.py <pptx> --out <dir> [--width 1280] [--mask-json palette.json]

mask-json 是 collect_ooxml_colors.py 的输出，里面的 image_regions_per_slide 用于在
渲染好的 PNG 上叠加半透明灰色 mask（让多模态模型识别图片区域，不参与主题色提取）。

输出：
    <out>/slide_001.png
    <out>/slide_002.png
    ...
    <out>/render_meta.json   # 每张图的尺寸 / 是否带 mask 等元数据
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

from pptx import Presentation
from pptx.util import Emu


def _emu_to_px(emu: int, slide_width_emu: int, png_width_px: int) -> int:
    return int(round(emu / slide_width_emu * png_width_px))


def _parse_only_indexes(s: str | None) -> set[int] | None:
    """解析 --only "1,3,5-7" → {1,3,5,6,7}。None/空返回 None（=全渲染）。"""
    if not s:
        return None
    out: set[int] = set()
    for part in s.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return out


def render_with_libreoffice(pptx_path: Path, out_dir: Path, width_px: int,
                           only: set[int] | None = None) -> list[Path]:
    """LibreOffice -> PDF -> PNG (via pypdfium2)。only=None 渲染所有页。"""
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        raise RuntimeError("soffice 未在 PATH 中")

    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise RuntimeError(
            "需要 pypdfium2 才能将 PDF 渲染为 PNG: pip install pypdfium2"
        ) from e

    with tempfile.TemporaryDirectory(prefix="ppt-clone-render-") as tmp:
        tmp = Path(tmp)
        # soffice 转 PDF（输出到指定目录）
        proc = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(tmp), str(pptx_path)],
            capture_output=True, text=True, timeout=180,
        )
        if proc.returncode != 0:
            raise RuntimeError(f"soffice 转 PDF 失败: {proc.stderr or proc.stdout}")

        pdf_path = tmp / (pptx_path.stem + ".pdf")
        if not pdf_path.exists():
            raise RuntimeError(f"未生成预期 PDF: {pdf_path}")

        pdf = pdfium.PdfDocument(str(pdf_path))
        out_paths: list[Path] = []
        for i in range(len(pdf)):
            idx_1based = i + 1
            if only is not None and idx_1based not in only:
                continue
            page = pdf[i]
            scale = width_px / page.get_width()
            pil_img = page.render(scale=scale).to_pil()
            out = out_dir / f"slide_{idx_1based:03d}.png"
            pil_img.save(out, "PNG")
            out_paths.append(out)
        pdf.close()
        return out_paths


def render_with_powerpoint_com(pptx_path: Path, out_dir: Path, width_px: int,
                              only: set[int] | None = None) -> list[Path]:
    """Windows PowerPoint COM 后端。only=None 渲染所有页。"""
    if platform.system() != "Windows":
        raise RuntimeError("PowerPoint COM 仅在 Windows 可用")
    try:
        import win32com.client
        import pythoncom
    except ImportError as e:
        raise RuntimeError("需要 pywin32: pip install pywin32") from e

    pythoncom.CoInitialize()
    ppt = None
    pres = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # 部分版本要求窗口可见才能 Open
        try:
            ppt.Visible = 0  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass
        pres = ppt.Presentations.Open(
            str(pptx_path.resolve()),
            ReadOnly=True, Untitled=False, WithWindow=False,
        )

        # 通过 slide_width 推算高度比例
        slide_w = pres.PageSetup.SlideWidth   # points (1/72 inch)
        slide_h = pres.PageSetup.SlideHeight
        height_px = int(round(width_px * slide_h / slide_w))

        out_paths: list[Path] = []
        for i, slide in enumerate(pres.Slides, start=1):
            if only is not None and i not in only:
                continue
            out = out_dir / f"slide_{i:03d}.png"
            slide.Export(str(out.resolve()), "PNG", width_px, height_px)
            out_paths.append(out)
        return out_paths
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:  # noqa: BLE001
            pass
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:  # noqa: BLE001
            pass
        pythoncom.CoUninitialize()


def render(pptx_path: Path, out_dir: Path, width_px: int = 1280,
           engine: str = "auto", only: set[int] | None = None
           ) -> tuple[list[Path], str]:
    """渲染 PPT 每页为 PNG。only 不为 None 时只渲染集合内的页（1-based）。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []

    if engine == "libreoffice":
        return render_with_libreoffice(pptx_path, out_dir, width_px, only), "libreoffice"
    if engine == "powerpoint_com":
        return render_with_powerpoint_com(pptx_path, out_dir, width_px, only), "powerpoint_com"
    if engine != "auto":
        raise ValueError(f"未知 engine: {engine}")

    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        try:
            return render_with_libreoffice(pptx_path, out_dir, width_px, only), "libreoffice"
        except Exception as e:  # noqa: BLE001
            errors.append(f"libreoffice: {e}")

    if platform.system() == "Windows":
        try:
            return render_with_powerpoint_com(pptx_path, out_dir, width_px, only), "powerpoint_com"
        except Exception as e:  # noqa: BLE001
            errors.append(f"powerpoint_com: {e}")

    raise RuntimeError(
        "未找到可用渲染引擎。已尝试: " + " | ".join(errors or ["none"])
    )


def overlay_image_mask(
    png_path: Path,
    image_boxes_emu: list[dict],
    slide_width_emu: int,
    slide_height_emu: int,
) -> None:
    """在原图上叠加半透明灰色 mask 标出图片区域。

    image_boxes_emu: [{"x": emu, "y": emu, "w": emu, "h": emu, "image": "name"}]
    会原地覆盖 png_path。
    """
    img = Image.open(png_path).convert("RGBA")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    for box in image_boxes_emu:
        x = _emu_to_px(box["x"], slide_width_emu, img.width)
        y = _emu_to_px(box["y"], slide_height_emu, img.height)
        w = _emu_to_px(box["w"], slide_width_emu, img.width)
        h = _emu_to_px(box["h"], slide_height_emu, img.height)
        # 半透明灰 + 红色细边框便于识别
        draw.rectangle([x, y, x + w, y + h], fill=(128, 128, 128, 110), outline=(255, 0, 0, 255), width=3)
        try:
            draw.text((x + 6, y + 6), "[IMAGE]", fill=(255, 0, 0, 255))
        except Exception:  # noqa: BLE001
            pass

    out = Image.alpha_composite(img, overlay).convert("RGB")
    out.save(png_path, "PNG")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--out", required=True, type=Path,
                    help="输出目录（不存在会创建）")
    ap.add_argument("--width", type=int, default=1280,
                    help="渲染宽度（像素），默认 1280")
    ap.add_argument("--mask-json", type=Path, default=None,
                    help="可选：collect_ooxml_colors.py 的输出 json，用于叠加图片 mask")
    ap.add_argument("--engine", choices=["auto", "libreoffice", "powerpoint_com"],
                    default="auto",
                    help="渲染引擎，默认 auto；同一任务建议显式锁定以避免 before/after 不一致")
    ap.add_argument("--only", default=None,
                   help="只渲染指定页（1-based），支持 '1,3,5-7' 格式；默认全渲染")
    args = ap.parse_args()

    if not args.pptx.exists():
        print(f"ERROR: pptx not found: {args.pptx}", file=sys.stderr)
        return 1

    only = _parse_only_indexes(args.only)
    pngs, used_engine = render(args.pptx, args.out, args.width,
                              engine=args.engine, only=only)

    # mask 叠加
    if args.mask_json and args.mask_json.exists():
        prs = Presentation(str(args.pptx))
        sw, sh = prs.slide_width, prs.slide_height
        mask_data = json.loads(args.mask_json.read_text(encoding="utf-8"))
        image_regions = mask_data.get("image_regions_per_slide", [])
        # 按 slide_index 索引 pngs
        png_by_idx = {_idx_from_name(p): p for p in pngs}
        for region in image_regions:
            slide_idx = region["slide"]
            png = png_by_idx.get(slide_idx)
            if png is not None:
                overlay_image_mask(png, region.get("boxes", []), sw, sh)

    def _idx_from_name(p: Path) -> int:
        import re as _re
        m = _re.search(r"slide_(\d+)", p.name)
        return int(m.group(1)) if m else 0

    meta = {
        "pptx": str(args.pptx),
        "width_px": args.width,
        "slides": [{"index": _idx_from_name(p), "path": str(p)} for p in pngs],
        "engine_requested": args.engine,
        "engine_used": used_engine,
        "masked": bool(args.mask_json),
        "only": sorted(only) if only else None,
    }
    (args.out / "render_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({"ok": True, "count": len(pngs), "out": str(args.out),
                      "engine": used_engine}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
