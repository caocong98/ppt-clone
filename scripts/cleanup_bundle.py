"""清理孤儿工作目录。

正常情况下 `bundle_workspace` 退出时会自动硬清整个工作目录——用户不需要手工调用本脚本。

仅在以下场景使用：
- 进程被 kill -9 / 机器断电 / 异常无法触发 finally，留下了孤儿工作目录（含 manifest.json）
- 用户想强制删掉某个工作目录

工作目录语义：
- 目录命名 `<src.stem>_clone/`，里面含 intermediate/ / snapshots/ / verify/ / report/ /
  scratch/ / manifest.json 等中间产物
- 最终交付 pptx **不在** 该目录内（在 src 同级），因此清理时整目录递归删除即可

CLI:
  python cleanup_bundle.py <工作目录> --dry-run
  python cleanup_bundle.py <工作目录> --yes
  python cleanup_bundle.py <工作目录> --yes --force   # 跳过 manifest 校验
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

MANIFEST_FILENAME = "manifest.json"
EXPECTED_SCHEMAS = {"ppt-clone-bundle/v1", "ppt-clone-bundle/v2"}


def _load_manifest(bundle: Path) -> dict | None:
    mp = bundle / MANIFEST_FILENAME
    if not mp.is_file():
        return None
    try:
        return json.loads(mp.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None


def cleanup(bundle: Path, *, dry_run: bool, force: bool) -> dict:
    if not bundle.exists():
        return {"ok": False, "reason": "bundle_not_found", "bundle": str(bundle)}
    if not bundle.is_dir():
        return {"ok": False, "reason": "bundle_not_directory", "bundle": str(bundle)}

    manifest = _load_manifest(bundle)
    if not force:
        if manifest is None:
            return {"ok": False, "reason": "missing_or_invalid_manifest",
                    "hint": "目录可能不是工作目录；用 --force 才能继续",
                    "bundle": str(bundle)}
        if manifest.get("schema") not in EXPECTED_SCHEMAS:
            return {"ok": False, "reason": "schema_mismatch",
                    "expected": sorted(EXPECTED_SCHEMAS),
                    "actual": manifest.get("schema"),
                    "bundle": str(bundle)}

    # 工作目录本身没有保留价值，整目录删除
    if dry_run:
        return {
            "ok": True, "dry_run": True,
            "bundle": str(bundle),
            "would_delete": [str(bundle)],
            "manifest_present": manifest is not None,
            "final_pptx_hint": manifest.get("final_pptx") if manifest else None,
        }

    try:
        shutil.rmtree(bundle)
    except OSError as e:
        return {"ok": False, "reason": "rmtree_failed",
                "bundle": str(bundle), "err": str(e)}

    return {
        "ok": True, "dry_run": False,
        "bundle": str(bundle),
        "deleted": [str(bundle)],
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("bundle", type=Path, help="孤儿工作目录路径")
    ap.add_argument("--dry-run", action="store_true",
                    help="只列出，不删除")
    ap.add_argument("--yes", action="store_true",
                    help="确认删除（非 dry-run 时必填）")
    ap.add_argument("--force", action="store_true",
                    help="跳过 manifest 校验（危险）")
    args = ap.parse_args()

    if not args.dry_run and not args.yes:
        print(json.dumps({
            "ok": False,
            "reason": "missing_confirmation",
            "hint": "非 dry-run 必须显式传 --yes 才会真正删除",
        }, ensure_ascii=False, indent=2))
        return 1

    result = cleanup(args.bundle.resolve(),
                     dry_run=args.dry_run,
                     force=args.force)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    sys.exit(main())
