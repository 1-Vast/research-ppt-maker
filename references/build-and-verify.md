# 执行与视觉验收

命令中的 `SKILL_DIR` 指当前技能文件夹，`OUT` 指任务输出文件夹；执行时替换为实际绝对路径，不依赖运行时 cwd。优先复用已安装依赖。

## 制作

**论文与项目页范围检查**：默认模式为 `paper-only`，只有用户明确要求汇报当前工作才可以设为 `paper-plus-project`，并在单独项目区记录实际实验来源。项目进展页不属于多篇论文 9–11 页的默认参考预算；具体按 [project-progress-mode.md](project-progress-mode.md)。

**多论文组会先做页数与模型图顺序审查**：按 [journal-club-density-and-figures.md](journal-club-density-and-figures.md) 写 `work/storyboard.json`（每篇 `papers[]` 记录角色、页数、总图的原始来源和核心实验），运行非阻断预警：

```text
python "SKILL_DIR/scripts/check_storyboard.py" "OUT/work/storyboard.json"
```

它只检查结构与选择候选问题，不读论文、不验证真正的图像内容，超出建议页数也不直接视作失败。**须人工确认可用原图实际用于总览；使用例外替代时核对来源与理由，并确认精选实验足以支撑主张。**

最小格式如下；完整计划继续补充背景、机制和证据，不把此例当作页数模板：

```json
{"mode":"journal-club","papers":[{"name":"论文题名","paper_type":"original-method","role":"primary","budget":7,"available_original_overview":true,"slides":[{"purpose":"cover"},{"purpose":"method-overview","visual":"original-paper-figure","source":"作者·年份·Fig.X及来源路径"},{"purpose":"evidence","experiment_id":"main"}]}]}
```

字段及深读/原图例外见 [journal-club-density-and-figures.md](journal-club-density-and-figures.md) 第 5 节。旧的自由格式计划仍可作为内部笔记；若运行检查器，须先整理为该格式。检查器不直接接受构建脚本的 `deck.json`。

制作规格 JSON 后、运行构建脚本前，建议执行语言预警工具：

```text
python "SKILL_DIR/scripts/lint_deck_language.py" "OUT/work/deck.json"
# 包含自定义绘制页面时仍需检查最终可编辑文字
python "SKILL_DIR/scripts/lint_deck_language.py" "OUT/主题-完成版.pptx"
```

它只检查部分易出现“审稿腔”、口语化提示和过长句，结果不阻断 PPT 制作；对 PPTX 读取可编辑文字及组合/表格文字，不读取图片内文字和 SmartArt，必须再结合逐页渲染审查；制作前还必须人工审核 [teaching-first-workflow.md](teaching-first-workflow.md) 的领域脉络、总模型图、图证对应和综述类型。出现警告时回到 [chinese-paper-explanation.md](chinese-paper-explanation.md) 人工改写。没有警告也不能说明幻灯片语言自然、机制讲得清楚。


**新建论文讲解**先按内容选布局。五种通用布局足够时可用构建脚本（依赖 `python-pptx`、Pillow；PDF 首页渲染需要 PyMuPDF）；复杂总览、机制放大或综述可按 deck-spec 的扩展示例制作可编辑页面，验收要求相同：

```text
python "SKILL_DIR/scripts/build_deck.py" "OUT/work/deck.json" -o "OUT/主题-完成版.pptx"
```

规格格式见 [deck-spec.md](deck-spec.md)。先写内容计划，再写规格，构建后立即渲染查看；输出里的 `WARNING` 表示首页缺失或文字在 13 pt 下仍放不下，删减对应文字后重建。脚本依赖本机字体文件测量文字，非 Windows 环境找不到字体时会按估算排版，更要逐页看图。

**修改已有 PPT 或使用用户模板**时，可用 `python-pptx`、PptxGenJS 或本地 PowerPoint。先查母版、版式、字体和现有页；对原生对象做修改，避免重建破坏动画、嵌入对象或备注；无法保真时说明并保留原件。

内容先规划，再实际保存 `.pptx`，不能只输出 Markdown 大纲。资料不足的页面不放虚构数据，必要时明确待补充。

## 读取和渲染

```text
python "SKILL_DIR/scripts/read_slides.py" "input.pptx" --mode compact --output "OUT/work/slide_extract.json"
python "SKILL_DIR/scripts/render_slides.py" "input.pptx" --output-dir "OUT/work/render-source"
```

复杂图示同时检查原图和文字。现有提取器主要处理顶层对象，并补取部分 OOXML 文本；组合对象、SmartArt、链接图表、继承格式或嵌入对象可能不完整。`compact` 不保留所有几何信息。页面视觉与 XML 不一致时复核原文件，不据不完整提取断言“没有图”。

新版渲染器：Windows 先用 PowerPoint 只读导出全部页；不可用时尝试 LibreOffice 导出 PDF 后逐页栅格化。输出按实际页序命名 `slide-001.png`，写 `render_manifest.json`，含源文件 SHA256、页数和后端。使用空输出目录，每轮修改用新子目录，避免旧图混入。

验证 PNG 数量等于 `.pptx` 实际页数；单个缩略预览不能通过。Quick Look 只生成预览，不用于整套验收。隐藏页也属检查范围。LibreOffice 需要可调用的 `soffice` 或安装在常规路径，PDF 栅格化需要 PyMuPDF (`fitz`)。

宏文件、旧 `.ppt` 或无法处理的格式使用本地应用导出副本，不覆盖原件。不要把工具失败描述成已验收。

## 必要时 OCR 与视觉证据

```text
python "SKILL_DIR/scripts/visual_inventory.py" --extract "OUT/work/slide_extract.json" --rendered-dir "OUT/work/render-source" --output "OUT/work/visual_inventory.json" --ocr auto --ocr-scope image-regions
python "SKILL_DIR/scripts/vision_review.py" --inventory "OUT/work/visual_inventory.json" --output "OUT/work/vision_review_packet.json"
```

两个辅助脚本用于图片/截图/复杂图密集的材料，不强制简单页生成冗余包。OCR 默认只处理图片区域，无 Tesseract 时直接判读截图。OCR 不是权威数据，不保证识别中文；组合图片或错误 bbox 时可用 `--ocr-scope full`，仍须核对页面。

审阅包仅是待填模板，不代表已经看过图。用图像查看工具实际打开 PNG，在 `work/qa.json` 或简短 QA 记录中注明文件、页序、发现及修复状态。

## 成品检查

**重做讲解前后对照**（只有用户提供旧稿时才加 `--baseline`）：

```text
python "SKILL_DIR/scripts/review_deck.py" "OUT/主题-完成版.pptx" --output "OUT/work/semantic-candidates.json"
python "SKILL_DIR/scripts/review_deck.py" "OUT/主题-完成版.pptx" --baseline "INPUT/旧稿.pptx" --output "OUT/work/revision-review.json"
```

检查器提示重复解释、失效标注和架构/指标并存候选；读取可编辑文字（含组合与表格），不读取图片文字、图表或 SmartArt。版本比较按页索引及图片顺序配对，统计图片字节、位置、裁剪与旋转；不自动对齐重排页面，组合对象需人工检查。图片变化统计不能判断讲解质量：重构任务逐项核对原始反馈、修复内容、图证与实际渲染，合适的旧图允许保留。详见 [revision-gate.md](revision-gate.md)。


```text
python "SKILL_DIR/scripts/check_deck.py" "OUT/主题-完成版.pptx" --output "OUT/work/structure-check.json"
python "SKILL_DIR/scripts/render_slides.py" "OUT/主题-完成版.pptx" --output-dir "OUT/work/render-final"
```

`check_deck.py` 检查可打开、ZIP 完整、页数、顶层越界对象、显式小字号（默认低于 12 pt，可用 `--min-pt` 调整）及可编辑文字/表/图表数量，并提示**图文相交候选**。后者只检查顶层图片与文本框，交叠面积达到文本框面积 35% 才提示，不判断叠放次序、透明度和图像内容；`overlap_groups_not_inspected` 报告未检查的组合对象数量。需对照实际渲染判断是否遮住原图的模块、文字、箭头、坐标或数据。越界背景可能合理，10 pt 的页脚出处和图注属于设计；这些都是**需要视觉判断的线索**，不自动报为设计失败。继承字体、文本溢出、组合坐标、原图小字、实际可读性和论证质量仍须看图。

视觉验收每页至少回答：核心信息在哪里，证据能否读懂，有无重叠/溢出，数据与条件是否对应，修改是否已复查。**模型总览第一次出现的是不是原论文原始 Figure（如可用），而非新绘制卡片？核心机制是否放在后续页面深入？实验是否精选且满足用户范围，超出默认 1–2 项时是否有必要？标题是否混入“据原文重绘”等制作过程语言？**先看联系表再看大图，密集图要局部放大。单页截图无法证明整套正常。

输出根目录只放完成版 `.pptx`，需要时附 PDF/预览；`work/` 放规格 JSON、来源、提取、计划、渲染和 QA。不要复制用户的参考 PPT 或私有图像到技能目录作为通用模板；`examples/` 只放公开论文的示范规格。
