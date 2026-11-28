# research-ppt-maker

从论文 DOI、标题、PDF 或项目描述出发，自行检索资料，制作直观、美观、可编辑的科研 PowerPoint。

## 使用

**默认仅制作论文内容。** 不自动插入项目进展、个人实验数据或未来工作。用户要求汇报当前工作时，独立规划经材料支持的项目页；局部修订已有混合稿时，按要求保留项目区；没有真实证据不得编造实验结果。可单独做项目进展，也可以做“论文讲解＋项目进展”。详见 `references/project-progress-mode.md`。

**组会篇幅按讲解难度分配。** 重点方法论文与辅助综述的紧凑汇报可参考合计 9–11 页。两篇原始研究或指定精读时，可分开讲观测任务、前作路线、原图总览与关键计算。详见 `references/journal-club-density-and-figures.md`。多篇论文默认逐篇独立讲解，以最后一篇的必要证据或局限结束；不自动添加双论文对照、横向比较或跨论文总结。仅在用户明确要求比较/综合型汇报时增加此类内容。各篇内部的相关工作、实验基线与消融不受影响。

将本仓库放入个人 skills 目录，文件夹名保持 `research-ppt-maker`。Windows 本机已安装路径为 `D:/CodexData/.codex/skills/research-ppt-maker`。

在 Codex 新聊天中输入：

```text
使用 $research-ppt-maker，根据论文《论文名称或 DOI》制作中文组会 PPT。
总共 4–5 页（含论文首页截图），重点讲清背景、前作、问题、设计与创新。
实验少而精，逻辑连贯，机制解释充分、美观直观，只交付 PPT，不要讲稿。
```

支持项目构想、进展汇报以及已有 PPT 的补全和美化。构想阶段使用验证计划，不能编造实验结果。

## 内容

- `SKILL.md`：技能入口与工作流。
- `references/`：资料检索、论文讲解、**面向听众的中文表达规范**、科研叙事、视觉设计及制作验收指导。
- `scripts/`：PPT 内容提取、渲染与结构检查辅助工具。
- `agents/openai.yaml`：技能展示信息。
- `examples/flow-matching/`：可运行的规格示例和论文首页图片。
- `CHANGELOG.md`：更新记录。

## 运行环境

技能说明由 Codex 读取；辅助脚本按需使用 Python。脚本依赖随用途不同，包含 `python-pptx`、Pillow、PyMuPDF，以及可选 OCR 工具。Windows 渲染优先调用已安装的 PowerPoint（需要 `pywin32`）；其他可用后端及具体命令见 `references/build-and-verify.md` 与脚本帮助。

对中文 PPT 可在构建前运行 `python scripts/lint_deck_language.py OUT/work/deck.json`，检查机械审稿栏目和抽象术语；预警仅供改写参考，不代替读原文和人工审阅。特别建议用本仓库 `references/chinese-paper-explanation.md` 的对照实例验收核心思想、方法和实验页。

复杂科研图可选使用当前环境可用的科研绘图技能；它们不属于本仓库的运行依赖。没有额外技能时，可用论文原图或可编辑图解继续。

本仓库仅包含技能及其支持资源，包含可运行示例及其首页截图，不包含参考 PPT、论文 PDF、生成的示范成品或本地制作记录。参考文件中提到的历史本地路径只说明来源，不是运行依赖。
