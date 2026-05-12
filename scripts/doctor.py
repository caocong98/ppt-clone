"""环境自检 + 孤儿临时目录清理 + bundle 目录扫描。

启动 Skill 前第一步运行。检查项：
- Python >= 3.9
- python-pptx, lxml, Pillow, numpy 安装与版本
- LibreOffice 在 PATH（soffice），Windows 备选 PowerPoint COM
- 清理 %TEMP% / /tmp 下名为 ppt-clone-* 的孤儿目录（旧 temp_workspace 残留）
- 可选 --scan-bundles <root>：列出包含合法 manifest.json 的 bundle 目录（仅查看，不删除）

输出 JSON 到 stdout（便于 Agent 解析）。Exit code 0=OK, 1=有问题。
"""

from __future__ import annotations

import argparse
import json
import platform
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _workspace import cleanup_orphans, MANIFEST_FILENAME  # noqa: E402

EXPECTED_SCHEMAS = {"ppt-clone-bundle/v1", "ppt-clone-bundle/v2"}

REQUIRED_PY = (3, 9)
REQUIRED_PKGS = {
    "pptx": "python-pptx",
    "lxml": "lxml",
    "PIL": "Pillow",
    "numpy": "numpy",
}


def _check_python() -> dict:
    v = sys.version_info
    ok = (v.major, v.minor) >= REQUIRED_PY
    return {
        "ok": ok,
        "version": f"{v.major}.{v.minor}.{v.micro}",
        "required": f">= {REQUIRED_PY[0]}.{REQUIRED_PY[1]}",
    }


def _check_packages() -> dict:
    results = {}
    missing = []
    for mod, pip_name in REQUIRED_PKGS.items():
        try:
            m = __import__(mod)
            ver = getattr(m, "__version__", "unknown")
            results[pip_name] = {"ok": True, "version": ver}
        except ImportError:
            results[pip_name] = {"ok": False, "version": None}
            missing.append(pip_name)
    return {"ok": not missing, "details": results, "missing": missing}


def _probe_libreoffice() -> dict | None:
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        return None
    try:
        out = subprocess.run(
            [soffice, "--version"],
            capture_output=True, text=True, timeout=10,
            errors="replace",
        )
        return {"ok": True, "engine": "libreoffice", "path": soffice,
                "version": (out.stdout or out.stderr).strip()}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "engine": "libreoffice", "error": str(e)}


def _probe_powerpoint_com() -> dict | None:
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client  # noqa: F401
        return {"ok": True, "engine": "powerpoint_com",
                "note": "Windows PowerPoint COM 可用"}
    except ImportError:
        return {"ok": False, "engine": "powerpoint_com",
                "fix": "pip install pywin32"}


def _check_renderer() -> dict:
    """探测全部可用引擎，给出 active_engine（约定优先级：LibreOffice > PowerPoint COM）。"""
    candidates = []
    lo = _probe_libreoffice()
    if lo is not None:
        candidates.append(lo)
    pc = _probe_powerpoint_com()
    if pc is not None:
        candidates.append(pc)

    available = [c for c in candidates if c.get("ok")]
    active = available[0]["engine"] if available else "none"
    return {
        "ok": bool(available),
        "active_engine": active,
        "candidates": candidates,
        "fix": None if available else "请安装 LibreOffice 并把 soffice 加入 PATH，或在 Windows 安装 pywin32",
    }


def _scan_bundles(root: Path) -> list[dict]:
    """扫描 root 下的孤儿工作目录（含合法 manifest.json）。

    正常工作流下 bundle_workspace 退出时会自动清理；只有异常中断（kill -9 等）
    才会留下孤儿工作目录。本函数用于发现这些孤儿，提示用户用 cleanup_bundle.py 清。
    """
    if not root.exists() or not root.is_dir():
        return []
    found: list[dict] = []
    candidates: list[Path] = [root]
    for child in root.iterdir():
        if child.is_dir():
            candidates.append(child)
            for grand in child.iterdir():
                if grand.is_dir():
                    candidates.append(grand)
    seen: set[Path] = set()
    for c in candidates:
        if c in seen:
            continue
        seen.add(c)
        mp = c / MANIFEST_FILENAME
        if not mp.is_file():
            continue
        try:
            data = json.loads(mp.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if data.get("schema") not in EXPECTED_SCHEMAS:
            continue
        try:
            total = sum(p.stat().st_size for p in c.rglob("*") if p.is_file())
        except OSError:
            total = 0
        final_pptx = data.get("final_pptx")
        final_pptx_name = data.get("final_pptx_name")
        final_exists = False
        if final_pptx:
            try:
                final_exists = Path(final_pptx).is_file()
            except OSError:
                final_exists = False
        found.append({
            "bundle": str(c),
            "kind": "orphan_working_dir",
            "started_at": data.get("started_at"),
            "finished_at": data.get("finished_at"),
            "status": data.get("status"),
            "src_pptx": data.get("src_pptx"),
            "final_pptx": final_pptx,
            "final_pptx_name": final_pptx_name,
            "final_pptx_exists": final_exists,
            "artifact_count": data.get("artifact_count"),
            "total_size_bytes": total,
            "cleanup_hint": f"python cleanup_bundle.py \"{c}\" --yes",
        })
    found.sort(key=lambda x: x.get("finished_at") or "", reverse=True)
    return found


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--scan-bundles", type=Path, default=None,
                    help="扫描该目录下的孤儿工作目录（异常中断残留，仅查看不删）")
    args = ap.parse_args()

    report = {
        "platform": platform.platform(),
        "scanned_at": datetime.now().isoformat(timespec="seconds"),
        "python": _check_python(),
        "packages": _check_packages(),
        "renderer": _check_renderer(),
        "orphans_cleaned": cleanup_orphans(),
    }
    if args.scan_bundles is not None:
        report["bundles"] = _scan_bundles(args.scan_bundles.resolve())
        report["bundles_root"] = str(args.scan_bundles.resolve())

    all_ok = (
        report["python"]["ok"]
        and report["packages"]["ok"]
        and report["renderer"]["ok"]
    )
    report["ok"] = all_ok

    report["active_engine"] = report["renderer"].get("active_engine", "none")

    if not all_ok:
        fixes = []
        if not report["python"]["ok"]:
            fixes.append(f"升级 Python 到 {report['python']['required']}")
        if report["packages"]["missing"]:
            fixes.append(
                "pip install " + " ".join(report["packages"]["missing"])
            )
        if not report["renderer"]["ok"]:
            fixes.append(report["renderer"].get("fix") or "安装渲染引擎")
        report["fix_commands"] = fixes

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
