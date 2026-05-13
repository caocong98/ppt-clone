# PPT Clone Skill

把任意 PPT 模板复刻成新主题的 PPT：

- 用户提供 **新大纲** 和 **目标主题**
- 原 PPT 的排版、对象、位置、字号、图片**完全不变**
- 文本被填入对应的占位区（按容量约束自动压缩）
- 主题色被替换为符合新主题语义的配色（OOXML × 视觉双向印证）

## 快速开始

```bash
pip install -r requirements.txt
python scripts/doctor.py
```

`doctor.py` 通过后，让你的 Cursor Agent 读取 [SKILL.md](SKILL.md) 即可启动。

## 环境要求

- Python ≥ 3.9
- LibreOffice（提供 `soffice` 命令）—— 用于把 slide 渲染成 PNG
- Windows 备选：Microsoft PowerPoint（通过 COM 接口）

## 设计哲学

- **Skill = 工具箱 + 操作手册**，不内嵌 AI 调用
- Agent 自身完成多模态/语言任务
- Python 脚本完成确定性 OOXML 操作
- 中间产物**零残留**（完成、失败、Ctrl+C 都清理）

## 目录结构

```
ppt-clone-skill/
  SKILL.md            Agent 入口与操作手册
  README.md           本文件（给人看）
  requirements.txt
  scripts/
    _workspace.py     统一临时空间 context manager
    doctor.py         环境自检 + 孤儿目录清理
    render_slides.py  PPT -> PNG（图片区域加 mask）
    collect_ooxml_colors.py  从 OOXML 采集可控色板
    analyze_template.py      占位区清单 + slide 角色识别
    rebuild_theme.py         theme 重写 + srgbClr 重映射 + 12 色槽替换
    apply_content.py         仅替换文本
    verify_effect.py         前后对比渲染
    color_utils.py           色差/对比度
```

## 使用模式

- **mode_a 不改色**：仅按占位区填入新文本
- **mode_b 改色 + 主题色化**：完整跑双向印证 + 新配色 + 闭环验证

详见 SKILL.md。