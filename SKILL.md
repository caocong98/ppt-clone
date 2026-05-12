---

## name: ppt-clone

description: 把任意 PPT 模板复刻为新主题。采用「三层混合架构」：L1 语义层（页角色 + 装饰 / LOGO / 序号过滤）+ L2 结构提示层（shape_group 弱并列提示）+ L3 硬约束层（每个 shape 独立的 char_limit / 段数 / 字号驱动容量）。LLM 只产 shape 级文本，apply 层带截断/拒写闸门 + 多段落 pPr 保留。最终 pptx 与输入同目录、文件名追加 `_clone`；运行期所有中间产物集中在 `<src.stem>_clone/` 临时工作目录，无论成败出口都被整目录硬清。当用户提供 PPT 模板（.pptx）+ 新大纲/主题、希望换文换色但保留排版时使用。
required_capabilities:

- vision: 多模态读图（主导色识别 / 装饰歧义文本判定 / 前后对比验证）
- language: 按 shape 级 char_limit + shape_group 并列约束写 blueprint，按 shape 单点压缩

# PPT Clone Skill

## 0. 你（Agent）必须完成的能力

- **视觉（Prompt A / B / verify）**：读 PNG 定主导色、给装饰歧义文本打标、验证前后一致
- **语言（Prompt C / C-retry / D）**：按 shape 级 char_limit + shape_group 并列约束写 content_blueprint，对超字 shape 做单点压缩
- **工具执行**：按工作流脚本顺序调用，按脚本返回 JSON 回灌结果

五个 Prompt：A 主题色 / B 装饰消歧 / C 蓝图作者（两步：page_theme → shape_texts）/ C-retry 蓝图修正 / D 单 shape 压缩。脚本侧用 `--resume <result.json>` 回灌。

## 1. 启动时必问 4 个问题


| #   | 问题                 | 选项 / 取值                       |
| --- | ------------------ | ----------------------------- |
| 1   | **新大纲 / 主题**       | 自由文本，至少含"主题方向 + 核心内容点"        |
| 2   | **是否改色**           | `mode_a` 保持原色 / `mode_b` 重新设计 |
| 3   | **配色策略**（仅 mode_b） | `preserve_style` / `redesign` |
| 4   | **输出语言**           | `zh-CN` / `en` / `bilingual`  |


> 最终 pptx 文件名固定为 `<src.stem>_clone.pptx`，与输入 pptx 同目录。
> 同名已存在时，启动前明确提示用户"将被覆盖"。

## 1.1 Agent 文件落位强约束（必读）

- 所有 driver / state json / 临时 png 等**必须**写到 `<工作目录>/scratch/` 内（`<工作目录>` = `<src.parent>/<src.stem>_clone/`）。
- **绝不**在 src 同级、skill 目录、项目根目录产生游离临时文件（包括用于规避 PowerShell 路径的"中转脚本"）。
- 工作目录**出口硬清**：`bundle_workspace` 退出时整目录递归删除。

## 1.2 Driver 脚本编码强约束（必读）

### A. 执行方式 — 统一通过 `run_driver.py`

绝不在 PowerShell 直接调用含中文路径的脚本：

```
python scripts/run_driver.py --script "$WORK/scratch/driver.py"
```

### B. Blueprint 写法 — JSON 字符串

```python
import json
BLUEPRINT_JSON = r"""
{
  "schema_version": 2,
  "slides": [
    {
      "slide_index": 1,
      "page_theme": "...",
      "shape_texts": {
        "slide_1::sp_5": "新标题",
        "slide_1::sp_8": ["第一段", "第二段"],
        "slide_1::sp_99": "__preserve__"
      }
    }
  ]
}
"""
BLUEPRINT: dict = json.loads(BLUEPRINT_JSON)
```

文本值中引用语用书名号 `「」` 或 `\u201c \u201d`；绝不在双引号字符串内嵌套双引号。

### C. 其他规范

- driver 首行 `# -*- coding: utf-8 -*-`
- `subprocess.run()` 一律 `encoding="utf-8", errors="replace"`
- 路径全部 `pathlib.Path`
- 复制 `scripts/_driver_template.py` 作为起点

## 2. 三层混合架构

```
┌──────────────────────────────────────────────────────────────┐
│ L1 语义层（页面理解 + 装饰过滤）                             │
│   story_role：cover / toc / chapter / content_list / ...     │
│   non_content_shapes：logo_image / decoration_number /        │
│     style_tag / footer_placeholder / template_sample_text    │
│   is_placeholder_text：识别"请输入内容"等占位提示语           │
│   global_style_guide：跨页一致性                              │
├──────────────────────────────────────────────────────────────┤
│ L2 结构提示层（弱提示 — 仅 LLM 可见，不影响 apply）          │
│   shape_group：同字数桶 + 桶内 1D 行/列对齐 + 数量 ≥ 3       │
│   group_hint："N 个 ~K 字 shape 横向并列，必须句式一致"      │
├──────────────────────────────────────────────────────────────┤
│ L3 硬约束层（apply 层不可逾越）                              │
│   char_limit.{min,max}（bbox 扣 inset + 1.4 行高 + 中文系数）│
│   hard_ceiling_chars（截断/拒写阈值，高于 char_limit.max）   │
│   per_paragraph[i]：逐段 font_size / char_limit / is_emphasis│
│   fill_ratio_target=0.75（占位 shape 强制留白 25%）           │
│   apply_content：                                            │
│     list + per_paragraph → 逐段按 p_hard 截断                │
│     emphasis 段超 p_max*1.5 → 单段保留原文（不波及整 shape） │
│     总字数 > hard_ceiling*2 才整 shape forced_skipped         │
│   多段落：保留 <a:pPr>（bullet / 缩进 / level）                │
└──────────────────────────────────────────────────────────────┘
```

旧的 `list_slot` / `enumeration_slot` / `bullet_group_slot` 概念已**全部废除**。

## 3. 工作流（共 17 步）

> `$WORK = <src.parent>/<src.stem>_clone/`，由 `_workspace.bundle_workspace(src_pptx)` 自动创建。
> 子目录：`intermediate/` / `snapshots/` / `verify/` / `report/` / `scratch/`，再加 `manifest.json`。
> **最终交付 pptx = `<src.parent>/<src.stem>_clone.pptx`**（不在 $WORK 内）。
> 出口自动整目录硬清。

### Step 0 · 环境自检

```
python scripts/doctor.py
```

### Step 1 · 打开工作目录 + 创建 driver

```python
from _workspace import bundle_workspace
with bundle_workspace(src_pptx) as bp:
    # bp.root, bp.scratch, bp.intermediate, bp.snapshots, bp.verify, bp.final_pptx
    ...
```

### Step 2 · analyze_template（几何/容量/颜色）

```
python scripts/analyze_template.py <src.pptx> --out $WORK/intermediate/analyze.json
```

### Step 3 · parse_template_story（三层混合解析，schema v2）

```
python scripts/parse_template_story.py <src.pptx> \
    --out $WORK/intermediate/template_story.json \
    [--logo-action keep_original|clear_to_empty] \
    [--shape-group-min 3]
```

产出 `template_story.json` (v2)：

- `slides[i].story_role`
- `slides[i].content_shapes[]`：每 shape 完整字段（`shape_id` = `slide_{i}::sp_{cNvPr.id}`、`bbox` / `bbox_region` / `original_text` / `char_count` / `paragraph_count` / `per_paragraph_char_count` / `font_size_pt` / `estimated_capacity` / `char_limit{min,max}` / `is_placeholder_text` / `has_bullet` / `parent_table_id`）
- `slides[i].shape_groups[]`：弱并列提示（`group_id` / `member_shape_ids` / `char_bucket` / `alignment_axis` / `group_hint`）
- `slides[i].non_content_shapes[]`：装饰/LOGO/页脚/序号/template_sample_text 等
- `vision_ambiguous[]`：待 Prompt B 判定

副产物 `template_story.json.debug.md`：人类可读清单（每 shape 一行 + group 列表 + non_content 列表），用于排查。

### Step 4 · render_slides（仅渲染 ambiguous 页）

```
python scripts/render_slides.py <src.pptx> --out $WORK/snapshots --engine <e> \
    --only <vision_pages>
```

### Step 5 · Prompt B → 装饰消歧（仅当 `vision_ambiguous` 非空）

输入：批次（≤ 8 shape / ≤ 3 slide）+ 对应 PNG。
输出：

```json
[{"shape_id": "slide_3::sp_42", "slide_index": 3,
  "role": "style_tag", "preserve_action": "keep_original",
  "confidence": 0.9, "reason": "封面风格标签"}]
```

写入 `$WORK/scratch/vision_result.json`，回灌：

```
python scripts/parse_template_story.py <src.pptx> \
    --out $WORK/intermediate/template_story.json \
    --story $WORK/intermediate/template_story.json \
    --resume $WORK/scratch/vision_result.json
```

### Step 6 · collect_ooxml_colors

```
python scripts/collect_ooxml_colors.py <src.pptx> --out $WORK/intermediate/ooxml_colors.json
```

### Step 7 · Prompt A → 主题色决策（mode_b）

输出 `decision.json` 写入 `$WORK/scratch/`。

### Step 8 · validate decision

```
python scripts/validate.py decision --decision $WORK/scratch/decision.json
```

### Step 9 · rebuild_theme

```
python scripts/rebuild_theme.py <src.pptx> --decision $WORK/scratch/decision.json \
    --out $WORK/intermediate/themed.pptx
```

mode_a：直接 `shutil.copyfile(src, themed)`。

### Step 10 · build_content_blueprint scaffold

```
python scripts/build_content_blueprint.py scaffold \
    --story $WORK/intermediate/template_story.json \
    [--outline $WORK/scratch/outline.txt] \
    [--topic-json $WORK/scratch/topic.json] \
    --out $WORK/intermediate/blueprint_scaffold.json
```

scaffold 对 LLM 暴露：

- 顶层 `global_style_guide`（thesis_prompt / terminology / must_use_words / banned_words / tone / voice / language / unit / seasonal_hint）
- `slides[i]`：`story_role` + `content_shapes[]`（含 `shape_id` / `bbox_region` / `char_limit` / `paragraph_count` / `per_paragraph_char_count` / `original_text_preview` / `is_placeholder_text` / `has_bullet` / `ph_type` / 可选 `parent_table_id`）+ `shape_groups[]`（弱并列提示）+ `non_content_count`

### Step 11 · Prompt C → content_blueprint（两步）

**第一步 — page_theme**：

> 阅读用户大纲 + global_style_guide + 每页 `story_role` + `original_text_preview`。
> 给每页一句话主题（≤ 30 字），形成全局叙事弧。

**第二步 — shape_texts**：按 shape_id 直接产文本：

```json
{
  "schema_version": 2,
  "title": "...",
  "global_style_guide": {...},
  "slides": [
    {
      "slide_index": 5,
      "page_theme": "团队 Q3 业绩与下季展望",
      "shape_texts": {
        "slide_5::sp_6":  "Q1",
        "slide_5::sp_10": "Q2",
        "slide_5::sp_13": "Q3",
        "slide_5::sp_16": "Q4",
        "slide_5::sp_23": "市场拓展",
        "slide_5::sp_42": "产品迭代",
        "slide_5::sp_46": "客户运营",
        "slide_5::sp_50": "团队建设",
        "slide_5::sp_22": ["第一段", "第二段"],
        "slide_5::sp_99": "__preserve__"
      }
    }
  ]
}
```

强约束：

1. **每个 content_shape 都必须有条目**（漏 shape_id → validate 报 `missing_shape_text`）
2. **所有 content_shape 默认都要替换为新主题的真实文本**（不论 `is_placeholder_text` 真假）。`non_content_shapes` 已经把 LOGO / 序号 / 装饰过滤掉了，`content_shapes` 里剩下的就是"本页需要填入新主题内容"的 shape。
3. **字数硬上限**：`len(text) ≤ char_limit.max`（多段为各段总和）
4. **多段**：`paragraph_count > 1` 的 shape 用 `string[]`，段数尽量等于原 paragraph_count（不等会有 warning，apply 仍能写入）
5. **per_paragraph 逐段硬约束**：scaffold 暴露 `per_paragraph[i].{font_size_pt, char_limit, is_emphasis}`；输出值必须是 `list[str]`，**每段字数不得超过对应段的 `char_limit.max`**；`is_emphasis=true` 段字号大、必须短（`[min, min+2]`）
6. **同 shape_group 句式一致**：`group_hint` 明确给出"必须句式一致、字数接近"
7. **留白优先**：单段 shape 推荐写到 `char_limit.max * 0.65`，不要写满（`fill_ratio_target = 0.65`）
8. `**is_placeholder_text=true` 的 shape 必须给真实新文本**，不可 `__preserve__`（validate 会强制报错）
9. `**__preserve__` 仅限极少数例外**：固定品牌名 / 版权声明 / 公司全称 / 固定日期格式。**每个 slide 使用 `__preserve__` 的 shape 数应 < 20% 的 content_shape 总数**，否则视为误用（validate 会给 warning）。
10. **不必管 `non_content_shapes`**（scaffold 不暴露具体 shape，只给数量）
11. **不得新增 scaffold 没有的 shape_id**（unknown_shape_id 会触发 warning）

写入 `$WORK/scratch/blueprint.json`。

### Step 12 · validate blueprint（+ Prompt C-retry）

```
python scripts/build_content_blueprint.py validate \
    --story $WORK/intermediate/template_story.json \
    --blueprint $WORK/scratch/blueprint.json
```

硬 error：

- `missing_shape_text`：缺 shape_id 条目
- `char_count_exceeds_limit`：超 `char_limit.max`
- `paragraph_char_count_exceeds_limit`：某段字数超 `per_paragraph[i].char_limit.max`
- `preserve_on_placeholder`：占位文本 shape 不可 `__preserve`__
- `invalid_shape_value_type` / `invalid_paragraph_value_type`

软 warning：

- `paragraph_count_mismatch` / `single_value_for_multi_paragraph`
- `emphasis_paragraph_too_long`：emphasis 段接近上限（建议 `[min, min+2]`）
- `single_paragraph_near_max`：单段 shape 接近 max（建议落 `[min, max*0.65]`）
- `unnecessary_preserve_on_content`：对 `is_placeholder_text=false` 的 shape 写了 `__preserve_`_，换主题场景下大概率是误用
- `too_many_preserves_on_slide`：单页 `__preserve__` 占比 > 30%，换主题语义变薄
- `char_count_below_min`
- `unknown_shape_id` / `missing_page_theme`
- `schema_version_mismatch`

若有 error → **Prompt C-retry**（最多 2 次）：

- 输入：上一版 bp + errors + scaffold
- 仅改报错 shape，其余原样返回
- 脚本用 JSON diff 强制：未报错 shape 回退上一版（保幂等）

```
python scripts/build_content_blueprint.py rollback \
    --prev $WORK/scratch/blueprint.json \
    --new $WORK/scratch/blueprint_retry.json \
    --errors $WORK/scratch/validate_errors.json \
    --out $WORK/scratch/blueprint.json
```

2 次仍失败 → fallback：把仍有 error 的 shape 直接写成 `"__preserve__"`（非占位 shape）或裁剪到 `char_limit.max-1` + …（占位 shape），并记入 `report/fallback.log`。

### Step 13 · map_blueprint_to_template

```
python scripts/map_blueprint_to_template.py \
    --story $WORK/intermediate/template_story.json \
    --blueprint $WORK/scratch/blueprint.json \
    --out $WORK/intermediate/content_mapping.json
```

产出 `content_mapping.json` (v4)：

- `slides[i].assignments[]`：`{key, value, shape_id, source, char_limit, ...}`
  - `value` 为 `string` / `string[]` / `"__skip__"` / `"__clear__"`
- `pending_shrink_requests[]`：超 `char_limit.max` 的 shape

若 `pending_shrink_requests` 非空 → **Prompt D**（按 shape 单点压缩）：

```json
{
  "text": "原文本",
  "char_limit": {"min": 8, "max": 14},
  "peer_texts": ["同页其它已定稿文本..."],
  "previous_attempts": []
}
```

输出单字符串。脚本对压缩结果做同页去重，冲突 → 重试最多 2 次。

```
python scripts/map_blueprint_to_template.py \
    --story ... --blueprint ... \
    --resume $WORK/scratch/shrink_result.json \
    --previous-mapping $WORK/intermediate/content_mapping.json \
    --out $WORK/intermediate/content_mapping.json
```

### Step 14 · validate mapping

```
python scripts/validate.py mapping \
    --story $WORK/intermediate/template_story.json \
    --mapping $WORK/intermediate/content_mapping.json
```

### Step 15 · apply_content（带 char_limit + 装饰兜底闸门）

```
python scripts/apply_content.py \
    $WORK/intermediate/themed.pptx \
    --mapping $WORK/intermediate/content_mapping.json \
    --story $WORK/intermediate/template_story.json \
    --out $WORK/intermediate/with_text.pptx
```

闸门顺序：

1. **L5.0 capacity gate**：`pending_shrink_requests` 非空 → 拒绝落盘（`--force-apply` 仅绕过此项）
2. **L5.1 装饰兜底**：原文是装饰字符（壹/Ⅰ/A/①）+ 新值远大于原文 → 强制跳过
3. **L5.2 硬字数闸门**（per_paragraph 优先，hard_ceiling 作阈值）：
  - list + per_paragraph：**逐段**用 `per_paragraph[i].hard_ceiling_chars` 截断；`is_emphasis` 段超 `p_max*1.5` → 本段保留原文（apply 层段内 `__skip_paragraph`__），不影响其它段
  - 总字数 > `sum(hard_ceiling_i) * 2` → 整 shape forced_skipped
  - str：chars > `hard_ceiling * 2` → forced_skipped；chars > `hard_ceiling` → 截断 + `…`
4. **多段落写入**：`string[]` value → `_apply_paragraphs`，保留 `<a:pPr>` bullet/缩进/level；段数不足时克隆最后一段的 pPr；`__skip_paragraph`__ 段保持原文字

### Step 16 · 落盘最终交付

```python
shutil.copyfile(bp.intermediate / "with_text.pptx", bp.final_pptx)
```

### Step 16.1 · 出口硬清（自动）

- 成功（final_pptx 存在）→ 整目录递归删除 $WORK
- 失败 → 整目录删 $WORK + 清半写入的 `_clone.pptx`

### Step 17 · lint + verify_effect（在 with 块内执行）

```
python scripts/lint_pptx.py \
    --pptx <src.parent>/<src.stem>_clone.pptx \
    --story $WORK/intermediate/template_story.json \
    --out $WORK/report/lint.json

python scripts/verify_effect.py \
    --before <src.pptx> --after <src.parent>/<src.stem>_clone.pptx \
    --engine <e> --out-dir $WORK/verify \
    [--content-bbox-overlap]   # 像素密度检测跑版兜底
```

## 4. Prompt 模板

### Prompt A — Theme Colors（mode_b）

输入 `ooxml_colors.json` + snapshots（遮罩后）+ `mode_b_strategy` + 用户主题。
输出：

```json
{
  "theme_name": "...",
  "mode": "preserve_style|redesign",
  "theme_colors": {"dk1":"...","lt1":"...","accent1":"...",...},
  "vision_corrected": [...],
  "wcag_checks": [...]
}
```

### Prompt B — Non-Content Disambiguation

输入：`vision_ambiguous[]` 批次 + 红框 PNG。

> 红框文字是模板预埋的"示例/装饰文字"还是用户应替换的内容？
>
> - 装饰标签 / 序号 / 已定稿品牌名 → `role=style_tag/decoration_number/logo_text`，`preserve_action=keep_original`
> - 真实占位提示语 → `role=template_sample_text`，`preserve_action=clear_to_empty`
> - 真实内容 → `role=content_slot`，`preserve_action=null`

每项输出：`{shape_id, slide_index, role, preserve_action, confidence, reason}`。

### Prompt C — Blueprint Author（两步）

**Step 1 — page_theme**：给每页写一句话主题（≤ 30 字）。

**Step 2 — shape_texts**：按 shape_id 产文本，遵守 §3 Step 11 强约束。

特别注意：

- `shape_groups` 中的成员**必须**句式一致、字数接近、属于同一个语义维度（如 4 个产品名 / 4 个季度 / 4 张并列卡片）
- **不要**把短标签和长正文写反 — `char_limit.max ≤ 10` 的 shape 只能写短标签，`char_limit.max ≥ 50` 才适合写长段落
- `bbox_region` 提示位置（如 `top-center` 通常是页标题，`bottom-right` 通常是页脚或装饰）
- `is_placeholder_text=true` 的 shape 在模板中是"Please enter your content / 内容概述"等占位提示语，**必须**用真实新内容替换
- **per_paragraph（逐段硬约束，multi-paragraph shape 必读）**：
  - scaffold 会给出 `per_paragraph[i] = {idx, font_size_pt, char_limit, is_emphasis, hint}`；输出值必须是 `list[str]`，段数与 `per_paragraph.length` 对齐，**每段的字数不得超过对应 `per_paragraph[i].char_limit.max`**。
  - `is_emphasis=true` 段字号明显偏大，**必须**短，优先落在 `[min, min+2]`；超过 `max*1.5` 会被 apply 层自动放弃该段写入并保留原模板文字占位。
  - 不同段字号差距越大，越要**按字号反比**分配字数：字号大段写 3~6 字结论；字号小段写常规要点。
- **留白优先**：单段 shape 推荐落在 `[char_limit.min, char_limit.max*0.75]`，保留 25% 视觉留白；只有真需要长段落时再填到 `max*0.9` 以内。

### Prompt C-retry — 修正

输入：上一版 bp + `validate_errors.json` + scaffold。只改报错 shape，其余原样。脚本仍 JSON diff 回退保幂等。

### Prompt D — Shrink-on-Demand（按 shape）

```json
{
  "text": "原文本",
  "char_limit": {"min": 8, "max": 14},
  "peer_texts": ["同页其它已定稿..."],
  "previous_attempts": []
}
```

输出单字符串：`len ≤ char_limit.max` 且优先落 `[min, max*0.95]`，不与 peer_texts 语义重合。

## 5. 主题色与内容编排顺序

```
parse(Step 3) ─┬─ render_slides(4) ─ Prompt B(5) ─ resume(5)
               │
               └─ collect_ooxml_colors(6) ─ Prompt A(7) ─ rebuild_theme(9) ─┐
                                                                              │
build_scaffold(10) ─ Prompt C(11) ─ validate(12) ─ map(13) ─ Prompt D(13) ─ apply(15)
                                                                              │
                                                                              ↓
                                                            <src.stem>_clone.pptx
```

主题色与内容生成**互不依赖**，可并行；apply_content 的输入是 `themed.pptx`，最终交付前可单独跑 lint/verify。

## 6. 不变项 & 绝不做的事

### 6.1 绝不做

- ❌ 不在 `$WORK` 外写临时文件
- ❌ 不在项目根/src 同级/skill 目录产生游离 `.py` / `.json`
- ❌ 不把 Blueprint 写成 Python dict 字面量
- ❌ 不在 `subprocess.run()` 省略 `encoding="utf-8"`
- ❌ 不把最终 pptx 输出到 `$WORK` 内
- ❌ 不修改源 pptx
- ❌ 不对 `non_content_shapes` (preserve_action=keep_original) 做任何 XML 改动
- ❌ 不对 chart / SmartArt / decorative_picture 做**文本**替换（装饰图位图可选见 §6.3）
- ❌ 不超 `char_limit.max`（apply 层会拒写或截断）
- ❌ 不把模板装饰字符（壹/Ⅰ/①/YOUR LOGO 等）替换为新内容

### 6.2 不变项

- 主题色链路：`collect_ooxml_colors` / `rebuild_theme` / `verify_effect`
- 渲染流程：`render_slides`（仅渲染 vision_pages）
- 字体策略、环境自检 (`doctor.py`)

### 6.3 可选：按文案重生成装饰图（位图）

在 `apply_content` 落盘 `<stem>_clone.pptx` 之后，可对 `template_story.json` 中 `role=decorative_picture` 的嵌入图（`p:pic` + `r:embed`）按**该页 blueprint** 拼 prompt 并重写 `ImagePart.blob`。**不处理**：`logo_image`、chart、SmartArt、链接图（无 `blip_rId`）、幻灯片背景 `blipFill`。

- 脚本：`python scripts/regenerate_slide_images.py --pptx <clone.pptx> --story $WORK/intermediate/template_story.json --blueprint $WORK/scratch/blueprint.json --provider pil|openai --out <路径>`
- `pil`：离线渐变 + 页主题字（无需密钥，用于管线验收）
- `openai`：需环境变量 `OPENAI_API_KEY`，默认 `dall-e-3`，生成后再按占位框比例 center-crop 写入
- 模板内多 shape **复用同一** `rId` 时只生成一次，所有引用同图（共享素材常见）

## 7. 命令速查


| 步   | 命令                                                                                                                                         |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| 0   | `python scripts/doctor.py`                                                                                                                 |
| 2   | `python scripts/analyze_template.py <src> --out $WORK/intermediate/analyze.json`                                                           |
| 3   | `python scripts/parse_template_story.py <src> --out $WORK/intermediate/template_story.json`                                                |
| 4   | `python scripts/render_slides.py <src> --out $WORK/snapshots --engine <e> --only <pages>`                                                  |
| 5   | `python scripts/parse_template_story.py <src> --out ... --story ... --resume $WORK/scratch/vision_result.json`                             |
| 6   | `python scripts/collect_ooxml_colors.py <src> --out $WORK/intermediate/ooxml_colors.json`                                                  |
| 8   | `python scripts/validate.py decision --decision $WORK/scratch/decision.json`                                                               |
| 9   | `python scripts/rebuild_theme.py <src> --decision ... --out $WORK/intermediate/themed.pptx`                                                |
| 10  | `python scripts/build_content_blueprint.py scaffold --story ... --out $WORK/intermediate/blueprint_scaffold.json`                          |
| 12  | `python scripts/build_content_blueprint.py validate --story ... --blueprint ...`                                                           |
|     | `python scripts/build_content_blueprint.py rollback --prev ... --new ... --errors ... --out ...`                                           |
| 13  | `python scripts/map_blueprint_to_template.py --story ... --blueprint ... --out $WORK/intermediate/content_mapping.json [--resume ...]`     |
| 14  | `python scripts/validate.py mapping --story ... --mapping ...`                                                                             |
| 15  | `python scripts/apply_content.py <themed.pptx> --mapping ... --story ... --out $WORK/intermediate/with_text.pptx`                          |
| 16  | `shutil.copyfile(with_text.pptx, <src.parent>/<src.stem>_clone.pptx)`                                                                      |
| 16b | `python scripts/regenerate_slide_images.py --pptx <clone> --story ... --blueprint ... --provider pil --out <stem>_clone_img.pptx`（可选 §6.3） |
| 17  | `python scripts/lint_pptx.py --pptx <clone> --story ... --out $WORK/report/lint.json`                                                      |
|     | `python scripts/verify_effect.py --before <src> --after <clone> --engine <e> --out-dir $WORK/verify [--content-bbox-overlap]`              |


孤儿目录处理：

```
python scripts/doctor.py --scan-bundles <root>
python scripts/cleanup_bundle.py <orphan_dir> --yes
```

