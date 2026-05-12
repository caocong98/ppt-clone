# -*- coding: utf-8 -*-
# =============================================================================
# PPT Clone Skill — Driver 脚本标准模板
# =============================================================================
# 复制此文件到 $WORK/scratch/ 并按指示修改标记了 <<FILL>> 的部分。
#
# 关键编码约定（必须遵守，勿删注释）
# ─────────────────────────────────
# 1. 文件头必须保留 `# -*- coding: utf-8 -*-`
# 2. **Blueprint 绝不写成 Python 字典字面量**。
#    原因：中文弯引号 "…" / 书名号等特殊标点一旦出现在 Python str 字面量的
#    双引号边界附近，就会产生 SyntaxError。
#    正确做法：把 Blueprint 内容写成 JSON 格式的 Python 字符串（三引号），
#    再通过 json.loads() 解析，或直接 json.dump 到文件后用 json.load 读回。
# 3. subprocess.run() 一律用 encoding="utf-8", errors="replace"，
#    不依赖系统默认编码（Windows 默认 cp936）。
# 4. 路径全部使用 pathlib.Path，不拼接字符串，不依赖 os.getcwd()。
# 5. 本脚本通过 run_driver.py --script <this_file> 执行，
#    不应直接被 PowerShell 以含中文字符的路径调用。
# =============================================================================

import json
import pathlib
import shutil
import subprocess
import sys

# ── 路径定义（用 pathlib.Path(__file__) 定位，不依赖 cwd）──────────────────
SCRATCH  = pathlib.Path(__file__).parent                         # scratch/
WORK     = SCRATCH.parent                                        # _clone/  (临时工作目录)
INTER    = WORK / "intermediate"
SCRIPTS  = pathlib.Path(__file__).parent.parent.parent / "ppt-clone-skill" / "scripts"
# <<FILL>> 修改为实际源文件路径
SRC      = pathlib.Path(r"<<FILL: /path/to/src.pptx>>")
# 最终交付路径（与 src 同级，追加 _clone）
FINAL_PPTX = SRC.parent / (SRC.stem + "_clone.pptx")


# ── 工具函数 ──────────────────────────────────────────────────────────────────
def run(cmd: list, ok_codes: tuple = (0,)) -> subprocess.CompletedProcess:
    """执行子进程；输出始终以 utf-8 解码，失败时 sys.exit。"""
    str_cmd = [str(c) for c in cmd]
    print(f"\n>>> {' '.join(str_cmd)}")
    result = subprocess.run(
        str_cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.stdout.strip():
        print(result.stdout[-6000:])
    if result.stderr.strip():
        print("[STDERR]", result.stderr[-3000:])
    if result.returncode not in ok_codes:
        print(f"FAIL exit={result.returncode}")
        sys.exit(result.returncode)
    return result


def write_blueprint(bp_data: dict) -> pathlib.Path:
    """把 blueprint dict 序列化为 JSON 并写入 intermediate/content_blueprint.json。
    使用 json.dumps 而非 Python 字面量，彻底规避引号冲突问题。"""
    dest = INTER / "content_blueprint.json"
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "w", encoding="utf-8") as f:
        json.dump(bp_data, f, ensure_ascii=False, indent=2)
    print(f"content_blueprint.json → {dest}")
    return dest


def hard_cleanup(success: bool) -> None:
    """成功时删工作目录；失败时同时删半写入的 final pptx。"""
    try:
        if WORK.exists():
            shutil.rmtree(WORK, ignore_errors=True)
            print(f"已删除工作目录: {WORK}")
    except OSError as e:
        print(f"WARNING 删除工作目录失败: {e}")
    if not success:
        try:
            if FINAL_PPTX.exists():
                FINAL_PPTX.unlink()
                print(f"已删除未完成的 final pptx: {FINAL_PPTX}")
        except OSError as e:
            print(f"WARNING 删除 final pptx 失败: {e}")


# ── Blueprint 数据 ─────────────────────────────────────────────────────────
# <<FILL>>
# 规则：
#   · 使用普通 ASCII 双引号包裹每个字符串值，如 "text": "清明习俗"
#   · 若文本本身含引号（如引用语），改用书名号 「」 或【】，
#     或用 \u201c \u201d Unicode 转义，**绝不**在值里嵌套同类引号
#   · 数字/布尔直接写，不加引号
# 建议格式：先写成 JSON 字符串，再 json.loads 解析（可在编辑器里验证 JSON 合法性）

BLUEPRINT_JSON = r"""
{
  "schema_version": 2,
  "artifact_type": "content_blueprint",
  "title": "<<FILL: PPT 标题>>",
  "global_style_guide": {
    "thesis_prompt": "<<FILL: 一句话核心论点>>",
    "terminology": [],
    "must_use_words": [],
    "banned_words": [],
    "tone": "professional",
    "voice": "third_person",
    "language": "zh-CN",
    "unit": "",
    "seasonal_hint": ""
  },
  "slides": [
    {
      "slide_index": 1,
      "story_role": "cover",
      "page_theme": "<<FILL: 一句话页面主题，<= 30 字>>",
      "shape_texts": {
        "slide_1::sp_5": "<<FILL: 标题文本（建议落 [char_limit.min, max*0.65]，留白）>>",
        "slide_1::sp_8": [
          "<<FILL: 第一段（≤ per_paragraph[0].char_limit.max）>>",
          "<<FILL: 第二段（≤ per_paragraph[1].char_limit.max）>>",
          "<<FILL: 结论段>>"
        ],
        "slide_1::sp_99": "<<FILL: 副标题或署名文本（默认仍需替换）>>"
      }
    }
  ]
}
"""
# shape_texts 语义说明（对照 blueprint_scaffold.json）：
#
#   · key 固定用 scaffold 给出的 shape_id（形如 "slide_N::sp_K"），不得自造；
#   · value 三种合法形态：
#       "一段文本"            — 单段 shape
#       ["段1", "段2", ...]    — 多段 shape（段数尽量对齐 paragraph_count）
#       "__preserve__"         — 【极少数例外】保留模板原文。仅限固定品牌名、
#                                版权声明、公司全称、固定日期格式等不适合换主题
#                                的内容；单页 __preserve__ 占比建议 < 20%，否则
#                                validate 会报 unnecessary_preserve_on_content /
#                                too_many_preserves_on_slide warning。
#   · **默认策略：对 scaffold 里所有 content_shape 都给出新主题对应的真实文本**。
#     non_content_shapes（LOGO / 序号 / 装饰）已被自动过滤，不在 scaffold 里，
#     content_shapes 里剩下的就是"应当被替换"的文本块。
#   · 每个 content_shape 都必须有条目；缺项会触发 missing_shape_text 硬 error。
#   · 字数约束（从严到宽）：
#       1) per_paragraph[i].char_limit.max → 每段硬上限（multi-paragraph shape）
#       2) char_limit.max → 整 shape 上限
#       3) hard_ceiling_chars → apply 层截断/拒写阈值（> 2× 会 forced_skipped）
#   · per_paragraph[i].is_emphasis=true 的段字号显著偏大，必须写短
#     （建议 [char_limit.min, min+2]）；否则会被 apply 层保护性跳过
#   · 留白优先：单段 shape 推荐落 [min, max*0.65]，fill_ratio_target=0.65
#   · 同 shape_group 的成员句式一致、字数接近、语义并列。
#
# 多段 shape 示例（emphasis 结尾强调段）：
#   "slide_8::sp_42": [
#     "要点要点要点约35字",
#     "另一要点约35字",
#     "简短结论6字"   // is_emphasis=true, per_paragraph[2].char_limit.max=10
#   ]
#
BLUEPRINT: dict = json.loads(BLUEPRINT_JSON)


# ── 主流程 ────────────────────────────────────────────────────────────────────
def main() -> int:
    ok = False
    try:
        assert SRC.exists(), f"源文件不存在: {SRC}"
        INTER.mkdir(parents=True, exist_ok=True)

        # Step 2: analyze_template
        run([sys.executable, SCRIPTS / "analyze_template.py",
             str(SRC),
             "--out", str(INTER / "template_spec.json")])

        # Step 3: parse_template_story
        run([sys.executable, SCRIPTS / "parse_template_story.py",
             str(SRC),
             "--out", str(INTER / "template_story.json")])

        # themed.pptx (mode_a: copy; mode_b: rebuild_theme)
        themed = INTER / "themed.pptx"
        shutil.copyfile(SRC, themed)

        # Step 10: scaffold
        run([sys.executable, SCRIPTS / "build_content_blueprint.py", "scaffold",
             "--story", str(INTER / "template_story.json"),
             "--out",   str(INTER / "blueprint_scaffold.json")])

        # Step 11: write blueprint (via json.dump, no Python literal issues)
        bp_file = write_blueprint(BLUEPRINT)

        # Step 12: validate blueprint
        run([sys.executable, SCRIPTS / "build_content_blueprint.py", "validate",
             "--story",     str(INTER / "template_story.json"),
             "--blueprint", str(bp_file)],
            ok_codes=(0, 2))

        # Step 13: map
        run([sys.executable, SCRIPTS / "map_blueprint_to_template.py",
             "--story",     str(INTER / "template_story.json"),
             "--blueprint", str(bp_file),
             "--out",       str(INTER / "content_mapping.json")])

        # Step 14: validate mapping
        run([sys.executable, SCRIPTS / "validate.py", "mapping",
             "--story",   str(INTER / "template_story.json"),
             "--mapping", str(INTER / "content_mapping.json")],
            ok_codes=(0, 2))

        # Step 15: apply content
        run([sys.executable, SCRIPTS / "apply_content.py",
             str(INTER / "themed.pptx"),
             "--mapping", str(INTER / "content_mapping.json"),
             "--story",   str(INTER / "template_story.json"),
             "--out",     str(INTER / "with_text.pptx")])

        # Step 16: copy final pptx to src 同级
        shutil.copyfile(INTER / "with_text.pptx", FINAL_PPTX)
        print(f"\n最终 PPTX: {FINAL_PPTX} ({FINAL_PPTX.stat().st_size:,} B)")

        ok = True
        return 0

    finally:
        hard_cleanup(ok)


if __name__ == "__main__":
    sys.exit(main())
