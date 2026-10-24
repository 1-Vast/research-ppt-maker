# 执行与视觉验收

命令中的 `SKILL_DIR` 指当前技能文件夹，`OUT` 指任务输出文件夹；执行时替换为实际绝对路径，不依赖运行时 cwd。优先复用已安装依赖。

## 制作

使用可用的 presentation authoring 工具、PptxGenJS、`python-pptx` 或本地 PowerPoint；根据模板保真、图表编辑和工具可用性选择，不要求固定引擎。用模板时先查母版、版式、字体和现有页。对原生对象做修改，避免重建破坏动画、嵌入对象或备注；无法保真时说明并保留原件。

全新制作时先用少量代表页面确认实际效果，再复用字体和颜色；这不要求暂停等用户批准。内容先规划，再实际保存 `.pptx`，不能仅输出 Markdown 大纲。资料不足的页面不放虚构数据，必要时明确待补充。

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

```text
python "SKILL_DIR/scripts/check_deck.py" "OUT/主题-完成版.pptx" --output "OUT/work/structure-check.json"
python "SKILL_DIR/scripts/render_slides.py" "OUT/主题-完成版.pptx" --output-dir "OUT/work/render-final"
```

`check_deck.py` 检查可打开、ZIP 完整、页数、顶层越界对象、显式小字号及可编辑文字/表/图表数量。越界背景可能合理，小字可能是出处；这些都是**需要视觉判断的线索**，不自动报为设计失败。继承字体、文本溢出、组合坐标、原图小字、实际可读性和论证质量仍须看图。

视觉验收每页至少回答：核心信息在哪里，证据能否读懂，有无重叠/溢出，数据与条件是否对应，修改是否已复查。先看联系表再看大图，密集图要局部放大。单页截图无法证明整套正常。

输出根目录只放完成版 `.pptx`，需要时附 PDF/预览；`work/` 放提取、计划、渲染和 QA。不要复制本次维护的参考 PPT 或私有图像到技能 assets 作为通用模板。
