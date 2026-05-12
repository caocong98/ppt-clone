# -*- coding: utf-8 -*-
"""
run_driver.py  —  编码安全的 driver 脚本执行入口
=================================================

用途
----
在 Windows / PowerShell 环境下，含中文字符的路径无法直接传给 `python <path>`。
本脚本作为**唯一的执行入口**，接受经由 JSON 安全编码的路径参数，从而避免：

  1. PowerShell 对 Unicode 路径的截断/乱码
  2. 在工作目录之外产生"中转脚本"游离文件

用法
----
    python scripts/run_driver.py --script <path_to_driver.py>

  · <path_to_driver.py> 为 $WORK/scratch/ 下的 driver 文件绝对路径
  · 本脚本在执行前先用 importlib 载入目标脚本，而非 subprocess，
    彻底规避路径传参时的 shell 编码问题

也支持直接运行子进程（当 driver 需要独立 sys.argv 时）：
    python scripts/run_driver.py --exec <path_to_driver.py> [-- <args...>]

  · 使用 subprocess 但以 Python list 方式传 argv，不经 shell 解析，
    因此路径不受 PowerShell 编码影响
"""

import argparse
import importlib.util
import os
import pathlib
import subprocess
import sys


def _import_and_run(script_path: pathlib.Path) -> int:
    """importlib 载入并执行 driver 脚本，返回 exit code。"""
    spec = importlib.util.spec_from_file_location("_driver_module", script_path)
    if spec is None or spec.loader is None:
        print(f"[run_driver] 无法加载: {script_path}", file=sys.stderr)
        return 1
    mod = importlib.util.module_from_spec(spec)
    # 让 driver 脚本的 __file__ 指向自身，使相对路径解析正确
    mod.__file__ = str(script_path)
    sys.modules["_driver_module"] = mod
    try:
        spec.loader.exec_module(mod)
        # 若 driver 定义了 main()，调用它
        if hasattr(mod, "main"):
            rc = mod.main()
            return int(rc) if rc is not None else 0
        return 0
    except SystemExit as e:
        return int(e.code) if e.code is not None else 0
    except Exception as exc:
        print(f"[run_driver] driver 异常: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


def _exec_subprocess(script_path: pathlib.Path, extra_args: list) -> int:
    """subprocess 方式运行 driver（argv 独立），不经 shell 解析路径。"""
    cmd = [sys.executable, str(script_path)] + extra_args
    result = subprocess.run(cmd, encoding="utf-8", errors="replace")
    return result.returncode


def main() -> int:
    # 强制 stdout/stderr 使用 utf-8，避免 Windows 控制台 cp936 截断
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="编码安全的 driver 脚本执行入口（importlib 或 subprocess）"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--script",
        type=pathlib.Path,
        help="用 importlib 载入并运行 driver.py（适合 driver 内部用 pathlib.Path(__file__) 定位自身）",
    )
    group.add_argument(
        "--exec",
        type=pathlib.Path,
        dest="exec_script",
        help="用 subprocess 运行 driver.py（argv 独立）",
    )
    parser.add_argument(
        "extra",
        nargs=argparse.REMAINDER,
        help="传给 --exec 模式下 driver 脚本的额外参数（用 -- 分隔）",
    )
    args = parser.parse_args()

    if args.script:
        target = args.script.resolve()
        if not target.exists():
            print(f"[run_driver] 脚本不存在: {target}", file=sys.stderr)
            return 1
        # 把 driver 所在目录加入 sys.path，使 driver 内的相对 import 可用
        driver_dir = str(target.parent)
        if driver_dir not in sys.path:
            sys.path.insert(0, driver_dir)
        return _import_and_run(target)

    if args.exec_script:
        target = args.exec_script.resolve()
        if not target.exists():
            print(f"[run_driver] 脚本不存在: {target}", file=sys.stderr)
            return 1
        extra = args.extra
        if extra and extra[0] == "--":
            extra = extra[1:]
        return _exec_subprocess(target, extra)

    return 0


if __name__ == "__main__":
    sys.exit(main())
