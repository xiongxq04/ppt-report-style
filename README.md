# 汇报 PPT 规范（ppt-report-style）

<<<<<<< HEAD
一套中文汇报 PPT 的排版规范 + 可执行的工作流。让 AI 生成的 PPT 在**字体、配色、版式骨架、配图、逐字稿**上稳定达标，而不是每次都返工。v0.4 起采用**三档严格度 × 三套配色预设**的可调节架构：草稿、组会、答辩各取所需，一套脚本全家通用。

<p align="center">
  <img src="skills/ppt-report-style/assets/example-slide.png" alt="示例：左文右图正文页" width="820">
</p>

<p align="center">
  <img src="skills/ppt-report-style/assets/example-table.png" alt="示例：三线表数据页" width="820">
=======
一套中文汇报 PPT 的排版规范 + 可执行的工作流。让 AI 生成的 PPT 在**字体、配色、版式骨架、配图与方块取舍、逐字稿**上稳定达标，而不是每次都返工。

<p align="center">
  <img src="skills/ppt-report-style/assets/example-slide.png" alt="示例：主图加侧栏的封面页" width="820">
</p>

<p align="center">
  <img src="skills/ppt-report-style/assets/example-table.png" alt="示例：强数据内容用表格承载" width="820">
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552
</p>

> 示例页为纯占位内容（虚构数据），用于展示版式，不含任何真实项目材料。

## 解决什么问题

<<<<<<< HEAD
- **档位化**：`fast`（内部草稿，一切放宽）/ `balanced`（默认，常规汇报）/ `strict`（答辩/投稿，个人学术规范全文生效）。字号下限、大数字、前缀、装饰预算、骨架重复、备注下限全部随档位分治，机审按档位执法。
- **配色成套**：三套色盲安全预设——`A` Okabe-Ito 学术蓝（默认）、`B` 深青蓝、`C` 砖红。每套是完整 token 角色表（墨/次级灰/发丝线/图表灰/点缀/点缀深/序列色），文字 token 全部 ≥4.5:1 且机审强制；换预设用 `--preset` 一键整套换。
- **字体分两槽**：中文只用宋体/楷体/黑体/微软雅黑（`a:ea`），西文只用 Arial / Times New Roman（`a:latin`）——英文不再借用中文字体的西文字形。
- **七档字阶**：`14(HF 页眉页脚档) / 16 / 18 / 20 / 30 / 36 / 44`，一页最多 5 档；正文硬下限 16pt，14pt 仅限页眉页脚极浅灰且仅在页眉/页脚几何区（机审按位置豁免，染灰正文照查）。
- **先判断内容关系，再选骨架**：关系模型库（并列/对比/流程/因果/层级/循环/辐射/矩阵/递进/演变/空间/时间线）+ 六种页面骨架；骨架重复检查按档位分治（strict 的统一骨架是美德，不是错误）。
- **点缀色纪律**：每页限额按形状计（strict ≤2 / balanced ≤3）；图表只点亮一处（`PAL.CHART_SERIES`）；关键数字句内黑体强调，strict 档不做大数字 callout。
- **截断轴诚实条款**：数值轴截断窗口必须在图注明写，结论句写真实量级——不许制造断层假象，也不许假装没截断。
- **配图分三档必要性**：必须实景 / 建议配图 / 可用图表替代；按放置尺寸判分辨率（<96dpi 报错，<150dpi 警告，`IMG.RES` 机审）。
- 每页备注写入约 1 分钟逐字稿（档位决定字数下限），数据来源与口径只放备注不上片。
- 交付前**统一检查入口**按错误/警告/建议/例外汇总；附含正例（三档 × 三预设全绿）与反例（必须命中 18 条预期规则）的**回归测试**。
=======
- **字体分两槽**：中文只用宋体/楷体/黑体（`a:ea`），西文只用 Arial / Times New Roman（`a:latin`）——英文不再借用中文字体的西文字形。
- **六档字阶**：`16 / 18 / 20 / 30 / 36 / 44`，内容对号入座（正文 18、副标与导语 20、页面标题 30、封面 44、表格与图注 16、大数字 36–44）；一页最多 5 档，黑体不叠加粗体。
- **配色成体系**：黑白灰是基础色、不占额度；主题色是一个**色阶系列**（默认 `#283E4E` 七档），另有一个对比色和一个辅助色。正文对比度 ≥ 4.5:1，单页色相组 ≤ 3 组，浅档只做底不做正文。
- **先判断内容关系，再选骨架**：关系模型至少覆盖 4 种，连续 3 页不得使用同一主骨架，方块矩阵类页面不超过约 35%。
- **每页至少一种视觉手段**，但允许纯文字页（不超过 15%～20%，且必须有字号层级与留白）；主视觉形状按语义用圆、平行四边形、圆柱、三角、六边形，不把每页都做成方块堆叠。
- **结论条按内容决定**，只有页面需要独立收束判断时才用，不作为统一骨架要求。
- **配图分三档必要性**：必须实景 / 建议配图 / 可用图表替代；只有"必须实景"才走完整搜索链路，批量在大纲阶段盘点，按放置尺寸判定分辨率，不用有水印、低画质的图。
- 每页备注写入**约 1 分钟逐字稿**，数据来源与口径只放备注不上片。
- 交付前**四道自检**：结构审查 → 样式校验 → HTML 预览图人眼检查 → 数字回溯原材料。
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552
- **不使用 SmartArt**；源文件已有 SmartArt 保持原状并用 `--allow-smartart` 声明。

## 安装

**方式一：skill-installer（Codex 内）**

```
$skill-installer install https://github.com/xiongxq04/ppt-report-style/tree/main/skills/ppt-report-style
```

**方式二：手动**

```bash
git clone https://github.com/xiongxq04/ppt-report-style.git
cp -r ppt-report-style/skills/ppt-report-style ~/.codex/skills/
```

安装后重启 Codex 即可被识别。

## 使用

<<<<<<< HEAD
安装后无需显式点名。说"按我的规范做 PPT""用答辩模式重排""把这个 PPT 重新排版"等需求时会自动启用；也可显式写：

```
$ppt-report-style 帮我做一套关于××的汇报 PPT，10 页，strict 档、配色用学术蓝
```

新建整套演示前，它会先定档位与配色预设、给一页纸大纲确认后再生成。

## 阈值真相源

所有阈值集中在 **`skills/ppt-report-style/references/profiles.json`**（三档 × 三预设）：字号下限、字阶、白名单、限额、备注下限、骨架规则。文档只引用不重定义，脚本经 `profile_lib.py` 读取——改一处，全仓生效。
=======
安装后无需显式点名。说"按我的规范做 PPT""把这个 PPT 重新排版""给这页配图表和逐字稿"等需求时会自动启用；也可显式写：

```
$ppt-report-style 帮我做一套关于××的汇报 PPT，10 页
```

新建整套演示前，它会先问清配色、素材、风格和特殊要求，并先给一页纸大纲确认后再生成。
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552

## 脚本

| 脚本 | 用途 |
| --- | --- |
<<<<<<< HEAD
| `scripts/check_ppt.py` | 统一检查入口：调度下面两个检查脚本，按**错误 / 警告 / 建议 / 例外**分级汇总，支持 `--tier`、`--preset`、`--json`、`--layout-report` |
| `scripts/verify_ppt_style.py` | 字阶档位、预设 token 配色与对比度、点缀每页限额、图表点亮数、双字体槽、黑体加粗、楷体入表格、前缀与口径（TYP./PAL./FONT. 规则） |
| `scripts/audit_ppt.py` | 结构审查：越界、文本溢出、西文不可断行、表格单元格撑高、元素重叠、字体槽、字形缺失、全角西文、备注、阴影与主题特效引用、配图分辨率、布局骨架（GEO./TXT./FONT./EFF./IMG./LAY./NOTE. 规则） |
| `scripts/palette_utils.py` | 预设对比度自检（`--check [--preset B|C]`）、对比度计算、自定义色阶派生（`--scale`） |
| `scripts/preview_ppt.py` | pptx → HTML 预览（`--slide N`，0 基；形状、图片、表格、柱/条/环/散点图近似渲染） |
| `scripts/build_example.py` | 生成 6 页学术示例（`--variant academic|negative`，`--preset A|B|C`） |
| `scripts/run_regression.py` | 回归基线：正例在 balanced/strict 零错误零警告，预设 B/C 端到端冒烟，反例命中全部预期规则 |

```bash
python3 -m pip install python-pptx
python scripts/palette_utils.py --check
python scripts/check_ppt.py 你的.pptx --tier balanced --preset A
python scripts/check_ppt.py 你的.pptx --tier strict --layout-report
python scripts/preview_ppt.py 你的.pptx out.html --slide 0
=======
| `scripts/check_ppt.py` | 统一检查入口：调度下面两个检查脚本，按**错误 / 警告 / 建议 / 例外**分级汇总，支持 `--json`、`--layout-report`、`--mode balanced\|strict` |
| `scripts/verify_ppt_style.py` | 配色方案与对比度、色相组数、面积配额、字号下限与**字阶档位**、中文字体（`a:ea`）与西文字体（`a:latin`）双白名单、黑体加粗、楷体入表格 |
| `scripts/audit_ppt.py` | 结构审查：越界、文本溢出、**西文不可断行**、**表格单元格撑高**、元素重叠、字体槽、字形缺失、全角西文、备注缺失、表宽，以及**布局骨架、配图质量与 SmartArt** |
| `scripts/palette_utils.py` | 配色定义、色阶派生、WCAG 对比度计算；`--check` 打印色阶对照表 |
| `scripts/preview_ppt.py` | 把 pptx 渲染成 HTML 预览（形状、图片、表格、柱/条/环/散点图近似渲染） |
| `scripts/build_example.py` | 重新生成 `assets/example-slide.pptx` 九页中性示例；`--variant negative` 生成回归反例 |
| `scripts/run_regression.py` | 回归测试：正例必须零错误零警告，反例必须命中全部预期规则 |

```bash
python3 -m pip install python-pptx
python scripts/check_ppt.py 你的.pptx
python scripts/check_ppt.py 你的.pptx --mode strict --layout-report
python scripts/check_ppt.py 你的.pptx --accent 1F4E79 --contrast-color 8C5A16 --min-size 18
python scripts/preview_ppt.py 你的.pptx out.html
python scripts/palette_utils.py --check
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552
python scripts/run_regression.py
```

预览图需要浏览器截图；没有 Playwright 时，把 HTML 直接打开或用 Chrome / Edge 的 headless 截图即可。

## 可配置项

| 项 | 默认 | 改法 |
| --- | --- | --- |
<<<<<<< HEAD
| 档位 | `balanced`（学术/答辩用 `strict`） | `--tier fast|balanced|strict`（旧参数 `--mode balanced\|strict` 等价） |
| 配色预设 | `A`（Okabe-Ito 学术蓝） | `--preset A|B|C`；临时换点缀色 `--accent` |
| 最小字号 / 字阶 / 档数 | 取自档位（profiles.json） | `--min-size` `--sizes` `--max-steps` |
| 字体白名单 | 取自档位 | `--fonts` `--latin-fonts`；不要往西文白名单加中文字体 |
| 额外允许色 | 无 | `--allow-colors` |
| 备注最少字数 | 取自档位（strict 150 / balanced 90 / fast 0） | `--notes-min` |
| 例外记录 | 无 | check_ppt `--exceptions exceptions.json`；页面级在备注写 `[例外:规则名]` + 理由 |
=======
| 主题色 | `#283E4E` | 生成脚本换常量；校验用 `--accent` |
| 对比色 | `#8C5A16` 系 | 校验用 `--contrast-color` |
| 辅助色 | `#2F4E48` 系 | 校验用 `--support` |
| 最小字号 | 16pt | 校验用 `--min-size` |
| 字阶档位 | 16,18,20,30,36,44 | `--sizes "16,18,20,28"`；每页档数 `--max-steps` |
| 中文字体白名单（`a:ea`） | 宋体,楷体,黑体 | 校验用 `--fonts "宋体,黑体"` |
| 西文字体白名单（`a:latin`） | Arial,Times New Roman | 校验用 `--latin-fonts "Arial"`；不要往这里加中文字体 |
| 额外允许色 | 无 | 校验用 `--allow-colors` |
| 备注最少字数 | 150 字/页 | 校验用 `--notes-min`；不需要时传 0 |
| 平衡档 / 严格档 | balanced | 校验用 `--mode strict` |
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552

系统缺少宋体/楷体时，AI 会先询问是否接受替代字体，不会静默替换。

## 目录结构

```
ppt-report-style/
├── .codex-plugin/plugin.json     # Codex skill-only plugin 清单
<<<<<<< HEAD
├── .github/workflows/validate.yml # CI：三档×预设机审 + 语义重建比对 + 正反例回归
├── CHANGELOG.md
├── LICENSE
└── skills/ppt-report-style/
    ├── SKILL.md                   # 硬性规范（加载即生效；档位与预设架构入口）
    ├── LICENSE.txt
    ├── agents/openai.yaml
    ├── assets/                    # 6 页学术示例与预览图
    ├── references/
    │   ├── profiles.json          # ★ 三档 × 三预设的阈值唯一真相源
    │   ├── workflow.md            # 分阶段工作流与返工触发条件
    │   ├── palette-system.md      # 配色预设与对比度
    │   ├── layout-patterns.md     # 关系模型与页面骨架（骨架重复按档位分治）
    │   ├── visual-means.md        # 视觉手段与纯文字页
    │   ├── effect-budget.md       # 章节母题与效果预算（fast/balanced 适用）
    │   ├── image-sourcing.md      # 配图必要性与素材来源
    │   ├── style-rules.md         # 档位表与形状、结论条规则
    │   ├── template-editing.md    # 模板跟随改版四步
    │   ├── checking.md            # 规则 ID、检查分级、例外协议与交付
    │   └── eval-queries.json      # 30 条触发测试用例
=======
├── LICENSE
├── README.md
└── skills/ppt-report-style/
    ├── SKILL.md                   # 硬性规范（加载即生效）
    ├── LICENSE.txt
    ├── agents/openai.yaml
    ├── assets/                    # 九页中性示例与预览图
    ├── references/
    │   ├── workflow.md            # 分阶段工作流与返工触发条件
    │   ├── palette-system.md      # 配色与对比度
    │   ├── layout-patterns.md     # 关系模型与页面骨架
    │   ├── visual-means.md        # 视觉手段与纯文字页
    │   ├── effect-budget.md       # 章节母题与效果预算
    │   ├── image-sourcing.md      # 配图必要性与素材来源
    │   ├── style-rules.md         # 平衡档 / 严格档与形状、结论条规则
    │   ├── template-editing.md    # 模板跟随改版四步
    │   ├── checking.md            # 检查分级、适配协议与交付
    │   └── eval-queries.json      # 51 条触发测试用例
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552
    └── scripts/                   # 检查 / 预览 / 示例生成 / 回归
```

## FAQ

<<<<<<< HEAD
- **档位怎么选**：内部草稿 `fast`；常规组会/项目汇报 `balanced`（默认）；答辩、投稿配套、正式评审 `strict`。
- **字体不全怎么办**：确认后可在 `--fonts` 里加入替代字体名。
- **英文为什么单独一槽**：python-pptx 的 `run.font.name` 只写 `a:latin`，中文字形取 `a:ea`。只设中文字体名会让英文用中文字体的西文字形渲染。跑历史文件可加 `--allow-cjk-latin` 暂时放行。
- **页眉页脚的浅灰为什么可以 14pt**：极浅灰档是刻意退后的辅助信息，且机审按几何位置豁免（页眉区 top+height≤0.7in、页脚区 top≥7.0in）——正文里染浅灰逃逸字号检查行不通。
- **纯文字页会不会被判失败**：封面/章节页（≥36pt 标题）不计入统计；内容页纯文字占比按档位阈值（balanced 30% / strict 60%）。
- **为什么不用 SmartArt**：SmartArt 跨版本打开容易变形且难以精确对齐；一律用普通形状、表格或图表表达。
- **图表预览与 PowerPoint 有差异**：预览是近似渲染，用于查溢出与对齐，最终以 PowerPoint 为准。
- **检查报错太多怎么办**：默认按"同页同规则最多 3 条"折叠（`--max-per-rule`）；合理的连续布局在备注写 `[例外:骨架重复]` 说明理由。

## 版本

v0.4.0（见 `CHANGELOG.md`）：三档 × 三预设架构落地；修复 v0.3→v0.4 过渡期 check_ppt 接口断裂、示例页数断言、D9D9D9 豁免位置盲洞等一致性问题；CI 升级为三档×预设矩阵 + 语义重建比对 + 正反例回归。
=======
- **字体不全怎么办**：确认后可在 `--fonts` 里加入替代字体名。
- **英文为什么单独一槽**：python-pptx 的 `run.font.name` 只写 `a:latin`，中文字形取 `a:ea`。只设中文字体名会让英文用中文字体的西文字形渲染（字距松、字形丑）；两槽分开设置后，英文走 Arial / Times New Roman。跑历史文件可加 `--allow-cjk-latin` 暂时放行。
- **灰色算不算限制颜色**：不算。黑白灰是基础色，不占配色额度；被限制的是主题色、对比色和辅助色这三类。
- **为什么不用 SmartArt**：SmartArt 在跨版本打开时容易变形，且难以精确对齐；本项目一律用普通形状、表格或图表表达同样的结构。
- **纯文字页会不会被判失败**：不会。允许纯文字页，但要求有字号层级、留白或对齐结构；只有"一大段文字居中放置"才报 `LAY.PURETEXT_FLAT`。
- **图表预览与 PowerPoint 有差异**：预览是近似渲染，用于查溢出与对齐，最终以 PowerPoint 为准。
- **检查报错太多怎么办**：默认按"同页同规则最多 3 条"折叠（`--max-per-rule`）；合理的连续布局可在备注写 `[例外:骨架重复]` 说明理由。
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552

## 许可证

MIT，见 [LICENSE](LICENSE)。

## 参考

- [Agent Skills 开放规范](https://agentskills.io)
- [Codex 插件与技能文档](https://developers.openai.com/codex/plugins/build)
<<<<<<< HEAD
- Okabe & Ito, Color Universal Design；Wong, B. (2011). Points of view: Color blindness. *Nature Methods*.

## English

A style guide + executable workflow for **Chinese-language reporting decks**, now with a **three-tier (fast / balanced / strict) × three-preset palette** architecture: draft, seminar and defense decks each get the right strictness, all driven from one machine-readable source of truth (`references/profiles.json`) and enforced by scripts.

- **Typography**: two font slots (Chinese in `a:ea`, Latin in `a:latin` — Arial / Times New Roman); seven-step type scale with a 16pt body floor (14pt reserved for the pale header/footer tier, position-checked).
- **Color**: three colorblind-safe presets (Okabe-Ito blue default, deep teal, brick red), each a full role table with WCAG-verified text tokens; per-slide accent budgets machine-enforced; charts light exactly one mark.
- **Honest data**: truncated axes must be declared in the caption; key numbers are emphasized in-sentence, not poster-sized (strict tier).
- **Layout from content relation**: a relation-model library and six skeletons; skeleton-repeat rules are tier-aware.
- **Checks**: unified entry (`check_ppt.py`) triages ERROR / WARNING / SUGGESTION / EXCEPTION with stable rule IDs; a regression suite requires the positive example to be clean on all tiers and presets and a synthetic negative deck to trip every expected rule.
- ~1 minute of speaker notes per slide; sources and caveats live in the notes. **No SmartArt**.

**Install**: `$skill-installer install https://github.com/xiongxq04/ppt-report-style/tree/main/skills/ppt-report-style` — **Scripts**: `check_ppt.py`, `verify_ppt_style.py`, `audit_ppt.py`, `palette_utils.py`, `preview_ppt.py`, `build_example.py`, `run_regression.py`. MIT licensed.
=======

## English

A style guide + executable workflow for **Chinese-language reporting decks**. It tells the agent how to lay out a slide so the result is consistent instead of something you have to re-lay out every time.

- Typography uses **two slots**: Chinese in SimSun / KaiTi / HeiTi (`a:ea`), Latin in **Arial or Times New Roman** (`a:latin`) — English never borrows a CJK font's Latin glyphs. A **six-step type scale** (`16 / 18 / 20 / 30 / 36 / 44` pt) maps sizes to roles; at most five steps per slide; HeiTi is never synthesised-bold.
- **Neutral black / white / grey are base colours and count against nothing.** The palette is a *theme scale* (seven steps of `#283E4E` by default) plus one contrast colour and one support colour. Body text must clear 4.5:1 contrast, a slide may use at most three hue groups, and light steps are for backgrounds — never for small body text.
- **Choose the layout from the content relation, not from a template**: at least four relation models per deck, no three consecutive slides on the same skeleton, and box/matrix pages capped at roughly 35%.
- **Every slide needs at least one visual means** — but plain-text slides are allowed (≤ 15–20% of the deck, and only with real type hierarchy, whitespace or alignment). Main shapes are chosen semantically (circle, parallelogram, cylinder, triangle, hexagon) instead of stacking boxes on every page.
- **The takeaway bar is content-driven**, not a mandatory component of every skeleton.
- **Images are triaged** as must-be-real / nice-to-have / chart-replaces-it; only the first tier justifies a full search chain, images are gathered in bulk at the outline stage, resolution is judged at the placed size, and watermarked or low-quality images are rejected.
- ~1 minute of speaker notes per slide; sources and caveats live in the notes, never on the slide. **No SmartArt** — existing SmartArt is left untouched and declared with `--allow-smartart`.
- Four checks before delivery: structural audit → style validation → rendered preview → trace every number back to the source material.

**Install**

```bash
$skill-installer install https://github.com/xiongxq04/ppt-report-style/tree/main/skills/ppt-report-style
```

**Scripts**: `check_ppt.py` (unified entry, error/warning/suggestion/exception triage), `verify_ppt_style.py` (palette, contrast, type scale, CJK + Latin font slots), `audit_ppt.py` (structure, layout, images, banned objects), `palette_utils.py` (colour scales and WCAG contrast), `preview_ppt.py` (pptx → HTML preview), `build_example.py` (regenerate the nine-slide example deck), `run_regression.py` (regression baseline).

MIT licensed.
>>>>>>> 5fc39cf9c0c9cf16c787045ee50c32d99c021552
