"""统一工作空间管理。

本模块定义 PPT Clone Skill 的"单一交付 + 临时工作目录"模型：

- 最终交付 pptx 固定落在**输入 pptx 同级目录**，命名 `<src.stem>_clone.pptx`。
- 运行期间的所有中间产物（intermediate / snapshots / verify / report / scratch / manifest）
  都放在工作目录 `<src.parent>/<src.stem>_clone/` 内；该目录**只在运行期存在**。
- `bundle_workspace` contextmanager 退出时自动硬清：
  - 成功：工作目录整删，只保留同级的 `<src.stem>_clone.pptx`
  - 失败（异常 / final_pptx 未落盘）：工作目录整删 + 清除半写入的 `<src.stem>_clone.pptx`

Agent 在 skill 运行期创建的 python driver、state json、日志文件等**必须**放在
`<工作目录>/scratch/` 内；出口会一并被清除。
"""

from __future__ import annotations

import atexit
import json
import os
import shutil
import signal
import sys
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterator

WORKSPACE_PREFIX = "ppt-clone-"
MANIFEST_FILENAME = "manifest.json"
MANIFEST_SCHEMA = "ppt-clone-bundle/v2"


# =====================================================================
# Bundle 模式（临时工作目录 + src 同级单文件交付）
# =====================================================================

@dataclass
class BundlePaths:
    """一次 clone 任务的工作区。

    `root`     工作目录，位于 src 同级，命名 `<src.stem>_clone/`
    `final_pptx` 最终交付 pptx 路径，位于 src 同级 `<src.stem>_clone.pptx`
                 （不在 root 内！）
    """
    root: Path
    src_pptx: Path | None = None
    started_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    extra: dict = field(default_factory=dict)

    @property
    def final_pptx(self) -> Path:
        """最终交付路径：src 同级目录 + <src.stem>_clone.pptx。"""
        if self.src_pptx is None:
            # 兜底：没有 src_pptx 时退回工作目录内（仅用于测试）
            return self.root / f"{self.root.name}.pptx"
        return self.src_pptx.parent / f"{self.src_pptx.stem}_clone.pptx"

    @property
    def manifest_path(self) -> Path:
        return self.root / MANIFEST_FILENAME

    @property
    def report(self) -> Path:
        return self.root / "report"

    @property
    def intermediate(self) -> Path:
        return self.root / "intermediate"

    @property
    def snapshots(self) -> Path:
        return self.root / "snapshots"

    @property
    def verify(self) -> Path:
        return self.root / "verify"

    @property
    def scratch(self) -> Path:
        """Agent 运行期临时文件（driver / state / log）放这里，出口会被清除。"""
        return self.root / "scratch"

    def ensure_subdirs(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        for sub in (self.report, self.intermediate, self.snapshots,
                    self.verify, self.scratch):
            sub.mkdir(parents=True, exist_ok=True)

    def _scan_artifacts(self) -> list[dict]:
        """扫描工作目录下所有文件（跳过 scratch/ 与 manifest 自身）。"""
        items: list[dict] = []
        if not self.root.exists():
            return items
        for p in sorted(self.root.rglob("*")):
            if not p.is_file():
                continue
            if p.name == MANIFEST_FILENAME and p.parent == self.root:
                continue
            try:
                rel = p.relative_to(self.root)
            except ValueError:
                continue
            if rel.parts and rel.parts[0] == "scratch":
                continue
            try:
                st = p.stat()
                items.append({
                    "path": rel.as_posix(),
                    "size": st.st_size,
                    "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                })
            except OSError:
                continue
        return items

    def write_manifest(self, *, status: str = "ok", error: str | None = None) -> Path:
        """在工作目录内写 manifest（运行期排查用；出口会被清除）。"""
        if not self.root.exists():
            return self.manifest_path
        artifacts = self._scan_artifacts()
        payload = {
            "schema": MANIFEST_SCHEMA,
            "schema_version": 2,
            "src_pptx": str(self.src_pptx) if self.src_pptx else None,
            "started_at": self.started_at,
            "finished_at": datetime.now().isoformat(timespec="seconds"),
            "status": status,
            "error": error,
            "final_pptx": str(self.final_pptx),
            "final_pptx_name": self.final_pptx.name,
            "final_pptx_present": self.final_pptx.is_file(),
            "artifact_count": len(artifacts),
            "total_size_bytes": sum(a["size"] for a in artifacts),
            "artifacts": artifacts,
            "extra": self.extra,
        }
        self.manifest_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return self.manifest_path

    def write_partial_manifest(self, *, error: str) -> Path:
        return self.write_manifest(status="failed", error=error)

    def _hard_cleanup(self, *, success: bool) -> None:
        """出口硬清。

        - 成功（success=True 且 final_pptx 存在）：只删工作目录 root；保留 final_pptx
        - 失败（异常 或 final_pptx 未落盘）：删工作目录 root + 清除半写入的 final_pptx
        """
        try:
            if self.root.exists():
                shutil.rmtree(self.root, ignore_errors=True)
        except OSError:
            pass
        if not success:
            try:
                if self.final_pptx.exists():
                    self.final_pptx.unlink()
            except OSError:
                pass


def _resolve_bundle_root(src_pptx: Path, bundle_root: Path | None,
                         name: str | None) -> Path:
    """工作目录默认位于 src_pptx 同级，命名 `<src.stem>_clone/`；冲突加 _run_<ts>。"""
    base = name or f"{src_pptx.stem}_clone"
    parent = bundle_root if bundle_root is not None else src_pptx.parent
    candidate = parent / base
    if not candidate.exists():
        return candidate
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_candidate = parent / f"{base}_run_{ts}"
    counter = 2
    while new_candidate.exists():
        new_candidate = parent / f"{base}_run_{ts}_{counter}"
        counter += 1
    return new_candidate


@contextmanager
def bundle_workspace(
    src_pptx: Path | str,
    *,
    bundle_root: Path | str | None = None,
    name: str | None = None,
) -> Iterator[BundlePaths]:
    """临时工作目录 + src 同级单文件交付。

    退出时：
    - 成功（yield 正常结束 且 final_pptx 已落盘）→ 删除工作目录；保留 final_pptx
    - 失败（异常 或 final_pptx 未落盘）→ 删除工作目录 + 清除半写入的 final_pptx
    """
    src = Path(src_pptx).resolve()
    root = _resolve_bundle_root(
        src,
        Path(bundle_root).resolve() if bundle_root else None,
        name,
    )
    bp = BundlePaths(root=root, src_pptx=src)
    bp.ensure_subdirs()
    ok = False
    try:
        yield bp
        ok = bp.final_pptx.is_file()
    except BaseException:
        ok = False
        raise
    finally:
        bp._hard_cleanup(success=ok)


# =====================================================================
# 兼容：temp_workspace（短任务/测试用）
# =====================================================================

def cleanup_orphans() -> int:
    """清理系统 tempdir 下名为 ppt-clone-* 的孤儿目录（temp_workspace 残留）。"""
    tmp_root = Path(tempfile.gettempdir())
    count = 0
    for child in tmp_root.glob(f"{WORKSPACE_PREFIX}*"):
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
            if not child.exists():
                count += 1
    return count


@contextmanager
def temp_workspace(prefix: str = WORKSPACE_PREFIX) -> Iterator[Path]:
    tmp = Path(tempfile.mkdtemp(prefix=prefix))

    cleaned = {"done": False}

    def _cleanup() -> None:
        if cleaned["done"]:
            return
        cleaned["done"] = True
        if tmp.exists():
            shutil.rmtree(tmp, ignore_errors=True)

    atexit.register(_cleanup)

    def _signal_handler(signum: int, _frame) -> None:
        _cleanup()
        sys.exit(128 + signum)

    prev_int = signal.signal(signal.SIGINT, _signal_handler)
    try:
        prev_term = signal.signal(signal.SIGTERM, _signal_handler)
    except (AttributeError, ValueError):
        prev_term = None

    try:
        yield tmp
    except BaseException:
        _cleanup()
        raise
    else:
        _cleanup()
    finally:
        signal.signal(signal.SIGINT, prev_int)
        if prev_term is not None:
            try:
                signal.signal(signal.SIGTERM, prev_term)
            except (AttributeError, ValueError):
                pass


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "cleanup-orphans":
        n = cleanup_orphans()
        print(json.dumps({"ok": True, "cleaned": n}))
    else:
        print(json.dumps({"usage": "python _workspace.py cleanup-orphans"}))
