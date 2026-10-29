# 构建脚本规格：`scripts/build_deck.py`

规格是一个 UTF-8 JSON 文件，放在 `OUT/work/deck.json`。相对路径（首页、图片、输出）都相对规格文件所在目录解析。

```text
python "SKILL_DIR/scripts/build_deck.py" "OUT/work/deck.json"            # 输出到规格里的 output
python "SKILL_DIR/scripts/build_deck.py" "OUT/work/deck.json" -o "OUT/主题-完成版.pptx"
```

脚本打印实际使用的字体和 `WARNING`。警告只有两类：首页缺失，或某块文字在 13 pt 下仍放不下。后者按提示删减那一块再重建，不要改脚本去放宽字号下限。脚本不能代替渲染检查。

完整可运行的例子：[examples/flow-matching/spec.json](../examples/flow-matching/spec.json)。

## 顶层字段

| 字段 | 说明 |
|---|---|
| `lang` | `zh`（默认）或 `en`；决定页眉、栏目标签和换行规则；中文默认使用研究背景/方法框架/方法创新/实验结果等学术栏目 |
| `output` | 输出 `.pptx` 路径 |
| `cite` | 页脚出处，如 `Lipman et al., ICLR 2023 · arXiv:2210.02747`；`cover` 里的 `cite` 会覆盖本组后续页 |
| `slides` | 按顺序排列的页面；每个 `cover` 开启一组新论文，进度条按组重置 |

每页都可用 `eyebrow` 覆盖默认栏目页眉，用 `source` 覆盖本页页脚出处。

## 文字标记

- `**强调**`：加粗；在核心想法、待解决问题、方法导语、本文改变等位置显示为陶土色。
- `x_{t}` 下标、`x^{2}` 上标、`~{Ho 2020}` 小号灰色引用。
- `\n`：手动换行。标题含 `\n` 时，脚本会缩小字号直到每行不再折行，用来避免把词拆开（如“训/练”）。
- 公式直接写 Unicode（`ℒ θ σ ∇ 𝔼 ‖ ²`），配合 `_{}`、`^{}`；不要写 LaTeX 源码。

## 版式与字段

字数指中文字符的大致上限，按 13.33 × 7.5 in 画布和默认字号估计；英文约按 2 个字母折 1 个字。超出时脚本会缩字号或报警。

### `cover` 标题页

| 字段 | 上限 | 说明 |
|---|---|---|
| `eyebrow` | 20 | 会议/期刊与年份 · 领域，如 `ICLR 2023 · 生成建模` |
| `title` | 2 行，每行 ≤ 13 | 中文题名，概括核心改变，不是原题直译；用 `\n` 控制断行 |
| `original_title` | 1–2 行 | 原文题目 |
| `authors` / `affiliation` | 各 1 行 | 作者过多时写前 3–5 位加 `et al.` |
| `idea` | 80 | 核心想法：改变了什么，带来什么 |
| `first_page` | — | 论文 PDF（自动渲染第 1 页）或首页 PNG/JPG，**必填** |
| `first_page_caption` | 20 | 如 `原文首页 · arXiv:2210.02747v2`、`原文首页 · Nature Methods 2024` |

### `background` 背景与问题

| 字段 | 上限 | 说明 |
|---|---|---|
| `title` | 28 | 结论式标题，一行 |
| `context.text` | 100 | 研究场景：任务输入/输出、当前实验或应用情境 |
| `context.example` | 40 | 具体例子，自造的写“示意”；可选 |
| `prior[]` | 2–3 条 | 每条 `name` ≤ 10，`cite` 写作者+年份，`did` ≤ 30（进展），`limit` ≤ 35（不足及其适用条件） |
| `gap.text` | 40 | 研究需求：一句具体的能力缺口 |
| `gap.approach` | 50 | 本文思路：说明采用什么方法回应需求，与方法页衔接，避免重复设问 |
| `gap.question` | 40 | 兼容旧规格的研究问题；新规格优先使用 `approach`，两者同时出现时只展示 `approach` |

### `method` 方法与实现

| 字段 | 上限 | 说明 |
|---|---|---|
| `title` | 28 | 结论式标题 |
| `lede` | 80 | 一句话核心思路，回应背景页的问题 |
| `lanes[]` | 1–2 条 | 每条 `label`（如 训练/推断）+ `steps`；只有一条流程时可直接写顶层 `steps` |
| `steps[]` | 3–5 步 | `name` ≤ 8；`text` 每行 4 步时 ≤ 30、3 步时 ≤ 40；最关键的一步加 `"key": true` |
| `formula` | — | `text` 一行公式；`note` ≤ 60，解释符号或定理的作用；可选 |
| `figure` | — | `{"path": ..., "caption": ...}`：改为“左原图、右步骤”版式，用于原文方法图本身就清楚的情况 |

### `innovation` 关键创新

| 字段 | 上限 | 说明 |
|---|---|---|
| `title` | 28 | 结论式标题 |
| `items[]` | 2–3 条 | `name` ≤ 9；`before` ≤ 40（已有方法，回指背景页前作）；`after` ≤ 40（本文改进）；`why` ≤ 30（主要作用，必要时带一个关键数字，区分设计意图和实验证实） |
| `summary` | 50 | 一句话核心转变；可选 |

### `evidence` 证据（可选）

左侧放一种主证据，右侧放结论：

| 字段 | 上限 | 说明 |
|---|---|---|
| `title` | 28 | 结论式标题，写清比较条件 |
| `table` | ≤ 6 行 × 5 列 | `columns`、`rows`（字符串）、`highlight_row`（本文所在行，从 0 计）、可选 `widths` 列宽比例；表头写指标方向 `↓/↑` |
| `chart` | ≤ 7 类 | `categories`、`values`、`highlight`（本文类别名）、`number_format`；单一指标的横向条形图 |
| `figure` | — | 原图或关键面板的图片路径 |
| `caption` | 80 | 数据来源（原文表/图号）、数据集、设定和指标含义 |
| `points[]` | 2–3 条，各 ≤ 50 | 每条以 `**结论词**：` 开头，对应一条创新 |
| `boundary` | 70 | 结论成立的边界或未验证的解释 |

## 扩展

五种版式不够时，在 Python 里复用同一套页面框架和组件，不另起配色与字体：

```python
import sys
from pathlib import Path
sys.path.insert(0, r'SKILL_DIR/scripts')
from build_deck import Deck, W, MX, BOTTOM

spec = {'lang': 'zh', 'cite': '作者, 会议 年份'}
deck = Deck(spec, Path(r'OUT/work'))
deck.group = ['background', 'method', 'innovation']            # 进度条显示的栏目
s = deck.new_slide()
top = deck.chrome(s, {'title': '机制放大：……', 'eyebrow': '它是怎么做的'}, 'method')
deck.text(s, MX, top, 6.0, '正文，支持 **强调** 与 x_{t}', 15)  # 返回实际高度（英寸）
deck.rect(s, 7.2, top, W - MX - 7.2, BOTTOM - top, 'panel', radius=0.06)
deck.prs.save(r'OUT/主题-完成版.pptx')
```

可用组件：`text`（返回高度，`draw=False` 只测量）、`fit`（求能放下的最大字号）、`rect`、`line`（`head=True` 带箭头）、`badge`（编号圆点）、`picture`（按比例放入框内）以及各版式方法 `cover/background/method/innovation/evidence`。颜色名见 `C`，含义见 [visual-design.md](visual-design.md)。
