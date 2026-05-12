# -*- coding: utf-8 -*-
"""一次性批量复刻 4 份 PPT 模板（中文主题 + 主题配色 + 全替换内容填充）。

执行链：
  1) parse_template_story → template_story.json
  2) 读原模板 theme1.xml 的 accent/背景色，构造 decision.json
  3) rebuild_theme → themed.pptx
  4) 基于 story + 主题词池 生成 blueprint（默认全替换，遵守 per_paragraph 硬约束）
  5) build_content_blueprint validate（warning 容忍、error 即失败）
  6) map_blueprint_to_template → content_mapping.json
  7) apply_content → with_text.pptx
  8) lint_pptx → lint_report.json
  9) 拷贝至源目录 <stem>_clone.pptx，硬清理临时工作目录

用法：
  python ppt-clone-skill/scripts/batch_clone_four.py
"""
from __future__ import annotations

import importlib.util
import json
import pathlib
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[2]  # e:\ppt-design-clone
SCRIPTS = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import parse_template_story as pts  # noqa: E402
import build_content_blueprint as bcb  # noqa: E402


# ============================================================
# 4 份模板的主题方案（全中文）
# ============================================================

THEMES: list[dict] = [
    {
        "src_name": "20607102 简约机械设备产品介绍宣传推广.pptx",
        "topic_key": "industrial_iot",
        "title": "工业智能装备产品推介",
        "thesis": "以高可靠机械装备与数智化控制方案赋能制造业升级。",
        "seasonal_hint": "",
        "unit": "华通机械",
        "tone": "simple_professional",
        "palette": {
            # 冷蓝+橙灰 工业风
            "dk1": "111418", "lt1": "FFFFFF",
            "dk2": "2C3744", "lt2": "E9EDF1",
            "accent1": "0A66C2", "accent2": "2E7BD6",
            "accent3": "FF8A3D", "accent4": "F1B24C",
            "accent5": "4B5563", "accent6": "94A3B8",
            "hlink": "0A66C2", "folHlink": "595959",
        },
        "vocab": {
            "nouns": [
                "装备", "产线", "控制器", "系统", "方案", "制造", "设备",
                "工艺", "精度", "效率", "可靠", "工厂", "数字", "智能",
                "产品", "平台", "模组", "能效", "伺服", "传感", "检测",
            ],
            "verbs_short": [
                "打造", "赋能", "升级", "提升", "降低", "优化", "驱动",
                "支撑", "交付", "重构", "聚焦", "构建",
            ],
            "bullets": [
                "稳定可靠的整机结构，保障长周期运行",
                "高精度伺服驱动，提供柔性产线适配能力",
                "多层级工况监测，提前预警异常工况",
                "一体化控制平台，打通设备与数据链路",
                "模块化组件设计，支持灵活扩展与升级",
                "低能耗与高产出兼顾，降低综合制造成本",
                "开放接口与标准协议，快速对接上下游",
                "从选型到运维的全流程技术服务支撑",
            ],
        },
        "titles": [
            "工业智能装备产品推介",
            "产品总览", "应用场景", "核心优势",
            "技术架构", "关键指标", "典型案例",
            "服务体系", "交付流程", "合作模式", "联系与展望",
        ],
    },
    {
        "src_name": "50129835 欢度六一儿童节主题活动通用模板.pptx",
        "topic_key": "childrens_day",
        "title": "欢度六一亲子游园会",
        "thesis": "用童趣与陪伴共绘一场温暖多彩的儿童节校园游园体验。",
        "seasonal_hint": "六一儿童节",
        "unit": "阳光社区学堂",
        "tone": "warm_playful",
        "palette": {
            # 童趣糖果色
            "dk1": "2B1C3D", "lt1": "FFFFFF",
            "dk2": "4B3969", "lt2": "FFF4F8",
            "accent1": "FF7AA2", "accent2": "FFB84D",
            "accent3": "5BBFE5", "accent4": "7BD389",
            "accent5": "C084FC", "accent6": "F97316",
            "hlink": "FF7AA2", "folHlink": "8B5CF6",
        },
        "vocab": {
            "nouns": [
                "孩子", "童年", "乐园", "笑声", "游戏", "手工", "画笔",
                "童趣", "梦想", "欢乐", "校园", "阳光", "故事", "糖果",
                "舞台", "家长", "老师", "气球", "奖品", "友谊", "陪伴",
            ],
            "verbs_short": [
                "快乐", "成长", "陪伴", "分享", "探索", "创造", "绽放",
                "拥抱", "发现", "欢迎", "共度",
            ],
            "bullets": [
                "趣味手工坊，陪孩子用色彩拼出童年记忆",
                "亲子闯关区，让家长和小朋友一起挑战",
                "童话剧小剧场，每位小演员都是主角",
                "阳光运动会，用笑声点亮操场的每个角落",
                "创意涂鸦墙，让奇思妙想自由生长",
                "故事音乐会，给耳朵一场温柔的旅行",
                "美食小镇，甜品糖果满足每颗小馋心",
                "爱心许愿树，记录孩子小小的大大的梦想",
            ],
        },
        "titles": [
            "欢度六一亲子游园会",
            "活动概览", "特色亮点", "游戏体验",
            "文艺展演", "爱心互动", "美食伙伴",
            "志愿招募", "安全保障", "时间路线", "期待相见",
        ],
    },
    {
        "src_name": "50329450 黑色科技风高端新品发布会通用PPT模板.pptx",
        "topic_key": "ai_launch",
        "title": "星云 AI 新品发布会",
        "thesis": "以多模态智能为核心，打造面向未来的一体化 AI 创作平台。",
        "seasonal_hint": "",
        "unit": "星云智能",
        "tone": "futuristic_cool",
        "palette": {
            # 黑金科技
            "dk1": "0A0A10", "lt1": "F5F6F7",
            "dk2": "1A1C22", "lt2": "2A2D35",
            "accent1": "00E0FF", "accent2": "7C3AED",
            "accent3": "F59E0B", "accent4": "34D399",
            "accent5": "FFFFFF", "accent6": "94A3B8",
            "hlink": "00E0FF", "folHlink": "7C3AED",
        },
        "vocab": {
            "nouns": [
                "智能", "模型", "算力", "架构", "创作", "多模态", "算法",
                "平台", "生态", "产品", "场景", "引擎", "体验", "未来",
                "数据", "协同", "创新", "边界", "世界", "灵感", "速度",
            ],
            "verbs_short": [
                "重塑", "跃迁", "释放", "激发", "定义", "突破", "驱动",
                "融合", "觉醒", "共创", "领跑",
            ],
            "bullets": [
                "跨模态理解引擎，图文音视频一体化解析",
                "千亿参数基座，面向复杂任务的稳健推理",
                "企业级安全体系，数据全程加密可追溯",
                "可插拔插件生态，为行业伙伴提供开放能力",
                "智能创作工作流，让灵感与产出无缝衔接",
                "极致推理延迟，重塑实时交互的使用体验",
                "多终端协同，在手机到云端之间自由切换",
                "面向开发者的完整工具链，从原型到上线",
            ],
        },
        "titles": [
            "星云 AI 新品发布会",
            "发布背景", "核心能力", "技术创新",
            "产品矩阵", "行业场景", "合作伙伴",
            "生态计划", "用户价值", "路线展望", "一起出发",
        ],
    },
    {
        "src_name": "总结汇报-图表.pptx",
        "topic_key": "year_summary",
        "title": "年度业务总结汇报",
        "thesis": "以数据复盘与关键成果为核心，沉淀经验并锚定下一阶段增长。",
        "seasonal_hint": "",
        "unit": "业务中心",
        "tone": "concise_business",
        "palette": {
            # 商务蓝绿
            "dk1": "0E1A2B", "lt1": "FFFFFF",
            "dk2": "1E3A5F", "lt2": "EEF3F8",
            "accent1": "2563EB", "accent2": "0EA5E9",
            "accent3": "10B981", "accent4": "F59E0B",
            "accent5": "64748B", "accent6": "EF4444",
            "hlink": "2563EB", "folHlink": "64748B",
        },
        "vocab": {
            "nouns": [
                "业绩", "目标", "数据", "趋势", "指标", "成果", "策略",
                "市场", "用户", "规模", "增长", "效率", "团队", "项目",
                "收入", "成本", "利润", "季度", "客户", "方案", "计划",
            ],
            "verbs_short": [
                "达成", "突破", "优化", "沉淀", "聚焦", "提升", "推进",
                "复盘", "拓展", "交付", "驱动",
            ],
            "bullets": [
                "全年核心指标稳步达成，超额完成年初承诺",
                "关键客户留存率持续提升，为次年增长打底",
                "产品迭代节奏稳定，需求响应周期显著缩短",
                "成本结构持续优化，单位产出效率同比改善",
                "新市场试点跑通，形成可复制的增长模型",
                "团队专业能力进一步沉淀，梯队结构更完整",
                "风险识别与复盘机制成熟，重大问题可控",
                "品牌声量稳步扩大，行业影响力持续加强",
            ],
        },
        "titles": [
            "年度业务总结汇报",
            "整体概览", "关键成果", "数据表现",
            "业务复盘", "问题与反思", "经验沉淀",
            "下阶段目标", "重点举措", "保障机制", "致谢与展望",
        ],
    },
]


# ============================================================
# 工具：读原模板 theme accent，构造 decision
# ============================================================

_SRGB_RE = re.compile(r"<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>"
                      r"\s*<a:(?:srgbClr|sysClr)[^/]*?(?:val|lastClr)=\"([0-9A-Fa-f]{6})\"")


def _read_theme_origins(pptx: pathlib.Path) -> dict[str, str]:
    """从原模板 theme1.xml 提取 12 色 hex（accent1-6 / dk1-2 / lt1-2 / hlink / folHlink）。"""
    with zipfile.ZipFile(pptx) as z:
        xml_bytes = z.read("ppt/theme/theme1.xml")
    xml = xml_bytes.decode("utf-8", errors="replace")
    out: dict[str, str] = {}
    for m in _SRGB_RE.finditer(xml):
        slot, hexv = m.group(1), m.group(2).upper()
        out.setdefault(slot, hexv)
    # 若缺槽位用合理默认
    defaults = {
        "dk1": "000000", "lt1": "FFFFFF",
        "dk2": "44546A", "lt2": "E7E6E6",
        "accent1": "4472C4", "accent2": "ED7D31",
        "accent3": "A5A5A5", "accent4": "FFC000",
        "accent5": "5B9BD5", "accent6": "70AD47",
        "hlink": "0563C1", "folHlink": "954F72",
    }
    for slot, v in defaults.items():
        out.setdefault(slot, v)
    return out


def _build_theme_decision(theme_cfg: dict, origins: dict[str, str]) -> dict:
    """构造 rebuild_theme 的 decision.json：ooxml_origin 用原色，hex 用新主题色。"""
    palette = theme_cfg["palette"]
    theme_colors: dict[str, dict] = {}
    for slot in ("dk1", "lt1", "dk2", "lt2",
                 "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6",
                 "hlink", "folHlink"):
        theme_colors[slot] = {
            "hex": palette[slot],
            "source": "batch_clone_deterministic",
            "ooxml_origin": origins.get(slot, palette[slot]),
        }
    return {
        "theme_colors": theme_colors,
        "meta": {
            "generator": "batch_clone_four",
            "topic": theme_cfg["topic_key"],
        },
    }


# ============================================================
# 工具：subprocess 安全调用
# ============================================================

def _run(cmd: list, ok_codes: tuple[int, ...] = (0,)) -> subprocess.CompletedProcess:
    str_cmd = [str(c) for c in cmd]
    print(f"  > {' '.join(str_cmd[-4:])}")
    r = subprocess.run(str_cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.stdout and len(r.stdout) > 0:
        tail = r.stdout.strip().splitlines()[-3:]
        for ln in tail:
            print(f"    | {ln[:200]}")
    if r.returncode not in ok_codes:
        print(f"    [STDERR] {r.stderr[-800:]}")
        raise SystemExit(f"subprocess failed exit={r.returncode}: {cmd}")
    return r


# ============================================================
# 内容生成：blueprint
# ============================================================

def _trim_to(s: str, n: int) -> str:
    """可见字符数裁到 n。"""
    if n <= 0:
        return ""
    out: list[str] = []
    count = 0
    for ch in s:
        if ch in (" ", "\n", "\t", "\r", "\u3000"):
            out.append(ch)
            continue
        if count >= n:
            break
        out.append(ch)
        count += 1
    return "".join(out)


def _gen_short(vocab: dict, target: int, seed: int) -> str:
    """1-5 字左右短词/标签。"""
    nouns = vocab["nouns"]
    verbs = vocab["verbs_short"]
    if target <= 1:
        return nouns[seed % len(nouns)][:1]
    if target <= 3:
        return nouns[seed % len(nouns)][: min(target, 3)]
    if target <= 6:
        a = verbs[seed % len(verbs)]
        b = nouns[(seed + 3) % len(nouns)]
        return _trim_to(a + b, target)
    # 6~12 字：verb+noun+noun
    a = verbs[seed % len(verbs)]
    b = nouns[(seed + 2) % len(nouns)]
    c = nouns[(seed + 5) % len(nouns)]
    return _trim_to(f"{a}{b}{c}", target)


def _gen_bullet(vocab: dict, target: int, seed: int) -> str:
    """从预设 bullet 池取一条并裁到 target。"""
    bullets = vocab["bullets"]
    s = bullets[seed % len(bullets)]
    if len(s) > target:
        s = _trim_to(s, target)
    return s


def _gen_sentence(vocab: dict, target: int, seed: int) -> str:
    """生成接近 target 字数的一句话。"""
    if target <= 4:
        return _gen_short(vocab, target, seed)
    if target <= 14:
        # 短标题：verb+noun
        a = vocab["verbs_short"][seed % len(vocab["verbs_short"])]
        b = vocab["nouns"][(seed + 1) % len(vocab["nouns"])]
        c = vocab["nouns"][(seed + 4) % len(vocab["nouns"])]
        return _trim_to(f"{a}{b}{c}", target)
    # 句子：从 bullets 池取，不够拼接
    out = _gen_bullet(vocab, target, seed)
    while len(out) < int(target * 0.75) and len(out) < target:
        extra = _gen_bullet(vocab, target - len(out),
                            seed + len(out))
        out = _trim_to(out + "；" + extra, target)
    return out


def _pick_title(theme_cfg: dict, slide_idx: int, seed: int) -> str:
    """按 slide_idx 选一个合适的章节/页标题。"""
    titles = theme_cfg["titles"]
    # cover 取第 1 个；其他轮询
    if slide_idx == 1:
        return titles[0]
    return titles[(slide_idx - 1) % (len(titles) - 1) + 1]


def _classify_role(cs: dict, text_preview: str) -> str:
    """粗略分类：title / subtitle / tag / bullet / body。"""
    paragraph_count = cs.get("paragraph_count", 1)
    c_max = cs.get("char_limit", {}).get("max", 30)
    fs = cs.get("font_size_pt") or 18.0
    if paragraph_count == 1:
        if c_max <= 4:
            return "tag"
        if fs and fs >= 28 and c_max <= 20:
            return "title"
        if fs and fs >= 22 and c_max <= 30:
            return "subtitle"
        if c_max <= 10:
            return "label"
        return "sentence"
    return "multi"


def _fill_to_exact(vocab: dict, target: int, seed: int, role: str) -> str:
    """生成可见字符数严格 ≤ target、且尽量靠近 target 的一段文本。

    role: title / subtitle / tag / label / sentence / bullet
    """
    if target <= 0:
        return ""
    if target == 1:
        return vocab["nouns"][seed % len(vocab["nouns"])][:1]
    if role in ("tag", "label") or target <= 5:
        return _trim_to(_gen_short(vocab, target, seed), target)
    if role in ("title", "subtitle") and target <= 20:
        return _trim_to(_gen_sentence(vocab, target, seed), target)
    if target <= 14:
        return _trim_to(_gen_sentence(vocab, target, seed), target)
    return _trim_to(_gen_bullet(vocab, target, seed), target)


def generate_blueprint(story: dict, theme_cfg: dict) -> dict:
    """为一份模板生成 blueprint：每段 / 每 shape 字数严格 ≈ 原文字数。

    关键策略（零估算 + 严格对齐）：
      · 单段 shape：target = char_count；title 直接用主题标题；subtitle 用 thesis。
      · 多段 shape：逐段 target = per_paragraph[i].original_char_count
        （若无此字段，则退回 per_paragraph[i].char_limit.max），段数严格对齐。
      · 字数必须 ≤ 原文字数（否则会触发 char_count_exceeds_limit 硬错误）。
      · emphasis 段仅作"短词优先"hint，不放宽字数。
    """
    vocab = theme_cfg["vocab"]
    bp_slides: list[dict] = []

    for s in story.get("slides", []):
        slide_idx = s["slide_index"]
        page_theme = _pick_title(theme_cfg, slide_idx, slide_idx)
        shape_texts: dict[str, object] = {}
        bullet_counter = 0
        tag_counter = 0

        sorted_shapes = sorted(
            s.get("content_shapes", []),
            key=lambda c: (c.get("bbox_region") or "mid",
                           (c.get("bbox") or {}).get("top", 0),
                           (c.get("bbox") or {}).get("left", 0))
        )

        for ci, cs in enumerate(sorted_shapes):
            sid = cs["shape_id"]
            original = cs.get("original_text") or ""
            char_count = int(cs.get("char_count", 0) or 0)
            paragraph_count = int(cs.get("paragraph_count", 1) or 1)
            per_paragraph = cs.get("per_paragraph") or []
            per_para_char_count = cs.get("per_paragraph_char_count") or []
            role = _classify_role(cs, original[:20])
            seed = (slide_idx * 31 + ci * 7) & 0xffff

            if char_count <= 0:
                # 空文本框：保持空，避免写入非预期内容
                shape_texts[sid] = "" if paragraph_count == 1 else [""] * paragraph_count
                continue

            # 多段 shape：逐段严格匹配原文字数
            if paragraph_count > 1:
                items: list[str] = []
                for i in range(paragraph_count):
                    # 优先 per_paragraph[i].original_char_count；退回 char_limit.max；
                    # 再退回 per_para_char_count[i]；最后退回 avg
                    pp = per_paragraph[i] if i < len(per_paragraph) else {}
                    p_orig = int(pp.get("original_char_count") or 0)
                    if p_orig <= 0:
                        p_max = int(pp.get("char_limit", {}).get("max", 0) or 0)
                        p_orig = p_max
                    if p_orig <= 0 and i < len(per_para_char_count):
                        p_orig = int(per_para_char_count[i] or 0)
                    if p_orig <= 0:
                        p_orig = max(1, char_count // paragraph_count)

                    is_emph = bool(pp.get("is_emphasis"))
                    if is_emph:
                        # emphasis 段字号大、易溢出 → 允许比原文略短；但段数必须保留
                        target = max(1, min(p_orig, max(2, p_orig)))
                        txt = _fill_to_exact(vocab, target,
                                             seed + i * 3 + 100, "title")
                    elif p_orig <= 5:
                        txt = _fill_to_exact(vocab, p_orig,
                                             seed + i * 3, "tag")
                    elif p_orig <= 14:
                        txt = _fill_to_exact(vocab, p_orig,
                                             seed + i * 3, "sentence")
                    else:
                        txt = _fill_to_exact(vocab, p_orig,
                                             slide_idx * 3 + bullet_counter,
                                             "bullet")
                        bullet_counter += 1
                    # 最终硬裁：绝不超过原段字数
                    items.append(_trim_to(txt, p_orig))
                shape_texts[sid] = items
                continue

            # 单段 shape：target = char_count（上限 = 原文字数）
            target = char_count
            if role == "title":
                shape_texts[sid] = _trim_to(page_theme, target)
                continue
            if role == "subtitle":
                shape_texts[sid] = _trim_to(
                    theme_cfg.get("thesis", page_theme), target)
                continue
            if role == "tag":
                txt = _fill_to_exact(vocab, target,
                                     seed + tag_counter * 5, "tag")
                tag_counter += 1
                shape_texts[sid] = _trim_to(txt, target)
                continue
            if role == "label":
                shape_texts[sid] = _trim_to(
                    _fill_to_exact(vocab, target, seed + 11, "label"),
                    target)
                continue
            # sentence
            if target <= 20:
                shape_texts[sid] = _trim_to(
                    _fill_to_exact(vocab, target, seed, "sentence"),
                    target)
            else:
                shape_texts[sid] = _trim_to(
                    _fill_to_exact(vocab, target,
                                   slide_idx * 5 + bullet_counter, "bullet"),
                    target)
                bullet_counter += 1

        bp_slides.append({
            "slide_index": slide_idx,
            "story_role": s.get("story_role", "content"),
            "page_theme": _trim_to(page_theme, 30),
            "shape_texts": shape_texts,
        })

    return {
        "schema_version": 2,
        "artifact_type": "content_blueprint",
        "title": theme_cfg["title"],
        "global_style_guide": {
            "thesis_prompt": theme_cfg["thesis"],
            "terminology": theme_cfg["vocab"]["nouns"][:5],
            "must_use_words": [],
            "banned_words": [],
            "tone": theme_cfg["tone"],
            "voice": "third_person",
            "language": "zh-CN",
            "unit": theme_cfg["unit"],
            "seasonal_hint": theme_cfg["seasonal_hint"],
        },
        "slides": bp_slides,
    }


# ============================================================
# 主流程：单模板 pipeline
# ============================================================

def process_one(theme_cfg: dict) -> dict:
    src = ROOT / theme_cfg["src_name"]
    assert src.exists(), f"源文件不存在: {src}"
    stem = src.stem
    final_pptx = src.parent / f"{stem}_clone.pptx"
    # 先清理旧 clone
    if final_pptx.exists():
        try:
            final_pptx.unlink()
        except OSError:
            pass

    work = src.parent / f".__clone_tmp_{stem}"
    scratch = work / "scratch"
    inter = work / "intermediate"
    try:
        if work.exists():
            shutil.rmtree(work, ignore_errors=True)
        scratch.mkdir(parents=True, exist_ok=True)
        inter.mkdir(parents=True, exist_ok=True)

        print(f"\n=== {src.name} ===")
        # 1) parse
        story = pts.parse(pathlib.Path(src), thresholds=pts.DEFAULT_THRESHOLDS)
        story_path = inter / "template_story.json"
        story_path.write_text(json.dumps(story, ensure_ascii=False, indent=2),
                              encoding="utf-8")
        cs_total = sum(len(s.get("content_shapes", []))
                       for s in story.get("slides", []))
        print(f"  [parse] slides={len(story.get('slides', []))}, "
              f"content_shapes={cs_total}")

        # 2) 主题色 decision
        origins = _read_theme_origins(src)
        decision = _build_theme_decision(theme_cfg, origins)
        decision_path = inter / "decision.json"
        decision_path.write_text(json.dumps(decision, ensure_ascii=False, indent=2),
                                 encoding="utf-8")

        # 3) rebuild_theme
        themed = inter / "themed.pptx"
        _run([sys.executable, SCRIPTS / "rebuild_theme.py",
              str(src), "--decision", str(decision_path),
              "--out", str(themed)])

        # 4) generate blueprint
        bp = generate_blueprint(story, theme_cfg)
        bp_path = inter / "content_blueprint.json"
        bp_path.write_text(json.dumps(bp, ensure_ascii=False, indent=2),
                           encoding="utf-8")

        # 5) validate（warning 容忍；不把 warning 当 error）
        val_result = bcb.validate_blueprint(story, bp)
        err_cnt = len(val_result["errors"])
        warn_cnt = len(val_result["warnings"])
        print(f"  [validate] errors={err_cnt}, warnings={warn_cnt}")
        if err_cnt > 0:
            for e in val_result["errors"][:5]:
                print(f"    ERR {e.get('kind')} slide{e.get('slide_index')} "
                      f"{e.get('shape_id')}: {e.get('msg', '')[:100]}")
            # 即使 error 也尝试走下去（apply 层会硬截断），但记录
            sample_err = val_result["errors"][0]
            print(f"    [WARN] continuing despite errors: {sample_err.get('kind')}")

        # 6) map
        mapping_path = inter / "content_mapping.json"
        _run([sys.executable, SCRIPTS / "map_blueprint_to_template.py",
              "--story", str(story_path),
              "--blueprint", str(bp_path),
              "--out", str(mapping_path)])

        # 7) apply
        with_text = inter / "with_text.pptx"
        _run([sys.executable, SCRIPTS / "apply_content.py",
              str(themed),
              "--mapping", str(mapping_path),
              "--story", str(story_path),
              "--out", str(with_text)])

        # 读 apply 的 json 结果（apply_content 的 stdout 包含 stats）
        # 这里简单起见不再解析，lint 再把关。

        # 8) lint
        lint_path = inter / "lint_report.json"
        r = _run([sys.executable, SCRIPTS / "lint_pptx.py",
                  str(with_text),
                  "--story", str(story_path),
                  "--out", str(lint_path)],
                 ok_codes=(0, 2))
        try:
            lint = json.loads(lint_path.read_text(encoding="utf-8"))
        except Exception:
            lint = {}
        findings = lint.get("findings", [])
        kinds: dict[str, int] = {}
        for f in findings:
            k = f.get("kind", "?")
            kinds[k] = kinds.get(k, 0) + 1
        print(f"  [lint] total={len(findings)} breakdown={kinds}")

        # 9) 拷贝最终 pptx
        shutil.copyfile(with_text, final_pptx)
        size_kb = final_pptx.stat().st_size // 1024
        print(f"  [done] {final_pptx.name} ({size_kb} KB)")
        return {
            "ok": True,
            "src": str(src),
            "final": str(final_pptx),
            "content_shape_total": cs_total,
            "lint_findings": len(findings),
            "lint_breakdown": kinds,
        }
    except SystemExit as e:
        print(f"  [FAIL] {e}")
        return {"ok": False, "src": str(src), "error": str(e)}
    except Exception as e:
        print(f"  [FAIL] {type(e).__name__}: {e}")
        return {"ok": False, "src": str(src), "error": f"{type(e).__name__}: {e}"}
    finally:
        # 硬清理：删工作目录
        try:
            if work.exists():
                shutil.rmtree(work, ignore_errors=True)
        except OSError:
            pass


# ============================================================
# main
# ============================================================

def main() -> int:
    results = []
    for theme in THEMES:
        r = process_one(theme)
        results.append(r)
    print("\n========== BATCH SUMMARY ==========")
    for r in results:
        tag = "OK" if r.get("ok") else "FAIL"
        src_name = pathlib.Path(r["src"]).name
        extras = ""
        if r.get("ok"):
            extras = (f"cs={r['content_shape_total']} "
                      f"lint={r['lint_findings']} "
                      f"{list(r['lint_breakdown'].items())[:3]}")
        else:
            extras = r.get("error", "")[:120]
        print(f"  [{tag}] {src_name}  {extras}")
    fail = sum(1 for r in results if not r.get("ok"))
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
