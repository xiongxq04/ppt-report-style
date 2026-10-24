# 汇报 PPT 规范（ppt-report-style）

一套中文汇报 PPT 的排版规范 + 可执行的工作流。让 AI 生成的 PPT 在**字体、配色、版式骨架、配图与方块取舍、逐字稿**上稳定达标，而不是每次都返工。

<p align="center">
  <img src="skills/ppt-report-style/assets/example-slide.png" alt="示例：主图加侧栏的封面页" width="820">
</p>

<p align="center">
  <img src="skills/ppt-report-style/assets/example-table.png" alt="示例：强数据内容用表格承载" width="820">
</p>

> 示例页为纯占位内容（虚构数据），用于展示版式，不含任何真实项目材料。

## 解决什么问题

- **字体分两槽**：中文只用宋体/楷体/黑体（`a:ea`），西文只用 Arial / Times New Roman（`a:latin`）——英文不再借用中文字体的西文字形。
- **六档字阶**：`16 / 18 / 20 / 30 / 36 / 44`，内容对号入座（正文 18、副标与导语 20、页面标题 30、封面 44、表格与图注 16、大数字 36–44）；一页最多 5 档，黑体不叠加粗体。
- **配色成体系**：黑白灰是基础色、不占额度；主题色是一个**色阶系列**（默认 `#283E4E` 七档），另有一个对比色和一个辅助色。正文对比度 ≥ 4.5:1，单页色相组 ≤ 3 组，浅档只做底不做正文。
- **先判断内容关系，再选骨架**：关系模型至少覆盖 4 种，连续 3 页不得使用同一主骨架，方块矩阵类页面不超过约 35%。
- **每页至少一种视觉手段**，但允许纯文字页（不超过 15%～20%，且必须有字号层级与留白）；主视觉形状按语义用圆、平行四边形、圆柱、三角、六边形，不把每页都做成方块堆叠。
- **结论条按内容决定**，只有页面需要独立收束判断时才用，不作为统一骨架要求。
- **配图分三档必要性**：必须实景 / 建议配图 / 可用图表替代；只有"必须实景"才走完整搜索链路，批量在大纲阶段盘点，按放置尺寸判定分辨率，不用有水印、低画质的图。
- 每页备注写入**约 1 分钟逐字稿**，数据来源与口径只放备注不上片。
- 交付前**四道自检**：结构审查 → 样式校验 → HTML 预览图人眼检查 → 数字回溯原材料。
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

安装后无需显式点名。说"按我的规范做 PPT""把这个 PPT 重新排版""给这页配图表和逐字稿"等需求时会自动启用；也可显式写：

```
$ppt-report-style 帮我做一套关于××的汇报 PPT，10 页
```

新建整套演示前，它会先问清配色、素材、风格和特殊要求，并先给一页纸大纲确认后再生成。

## 脚本

| 脚本 | 用途 |
| --- | --- |
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
python scripts/run_regression.py
```

预览图需要浏览器截图；没有 Playwright 时，把 HTML 直接打开或用 Chrome / Edge 的 headless 截图即可。

## 可配置项

| 项 | 默认 | 改法 |
| --- | --- | --- |
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

系统缺少宋体/楷体时，AI 会先询问是否接受替代字体，不会静默替换。

## 目录结构

```
ppt-report-style/
├── .codex-plugin/plugin.json     # Codex skill-only plugin 清单
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
    └── scripts/                   # 检查 / 预览 / 示例生成 / 回归
```

## FAQ

- **字体不全怎么办**：确认后可在 `--fonts` 里加入替代字体名。
- **英文为什么单独一槽**：python-pptx 的 `run.font.name` 只写 `a:latin`，中文字形取 `a:ea`。只设中文字体名会让英文用中文字体的西文字形渲染（字距松、字形丑）；两槽分开设置后，英文走 Arial / Times New Roman。跑历史文件可加 `--allow-cjk-latin` 暂时放行。
- **灰色算不算限制颜色**：不算。黑白灰是基础色，不占配色额度；被限制的是主题色、对比色和辅助色这三类。
- **为什么不用 SmartArt**：SmartArt 在跨版本打开时容易变形，且难以精确对齐；本项目一律用普通形状、表格或图表表达同样的结构。
- **纯文字页会不会被判失败**：不会。允许纯文字页，但要求有字号层级、留白或对齐结构；只有"一大段文字居中放置"才报 `LAY.PURETEXT_FLAT`。
- **图表预览与 PowerPoint 有差异**：预览是近似渲染，用于查溢出与对齐，最终以 PowerPoint 为准。
- **检查报错太多怎么办**：默认按"同页同规则最多 3 条"折叠（`--max-per-rule`）；合理的连续布局可在备注写 `[例外:骨架重复]` 说明理由。

## 许可证

MIT，见 [LICENSE](LICENSE)。

## 参考

- [Agent Skills 开放规范](https://agentskills.io)
- [Codex 插件与技能文档](https://developers.openai.com/codex/plugins/build)

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
