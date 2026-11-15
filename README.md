# research-ppt-maker

从论文 DOI、标题、PDF 或项目描述出发，自行检索资料，制作直观、美观、可编辑的科研 PowerPoint。

## 使用

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

复杂科研图可选调用同级个人技能 `gpt-image-2-research-studio`；该技能未包含在本仓库。没有该技能时，可用论文原图或可编辑图解继续。

本仓库仅包含技能及其支持资源，包含可运行示例及其首页截图，不包含参考 PPT、论文 PDF、生成的示范成品或本地制作记录。参考文件中提到的历史本地路径只说明来源，不是运行依赖。
