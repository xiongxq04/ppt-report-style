# PPT 检查与交付规则

规则阈值只在本 skill 的专项文件定义，检查脚本实现这些阈值；本文件规定分级、协议和交付方式。

## 检查结果分级

### 错误（ERROR）

交付前必须修复：

- 明确的文字溢出、越界或元素遮挡；
- 字号低于要求，或文字颜色不在配色方案内；
- 文字或图表对比度不达标；
- 图表负值、比例或数值表达错误；
- 图片缺失，或按放置尺寸计算的分辨率严重不足；
- 数字无法追溯到材料或实算；
- 用户要求保留的模板元素被意外删除；
- 严格档下新增的 SmartArt、继承字体颜色。

### 警告（WARNING）

需要人工确认，不自动判定失败：

- 连续三页使用同一主骨架、纯文字页比例偏高、方块矩阵页过半；
- 对比色或辅助色面积超过建议值；
- 图片分辨率偏低、长边像素不足、可能与文字重叠；
- 主题字体、继承颜色无法可靠解析；
- HTML 近似预览与 PowerPoint 最终渲染可能不同；
- 文本溢出估算接近阈值；
- 字符可能在特定字体中缺少字形（GBK 范围外）；
- 组合形状、嵌入对象未递归审查；
- 平衡档下检出 SmartArt。

### 建议（SUGGESTION）

用于提升质量：

- 关系模型覆盖不足，或全篇骨架种类偏少；
- 图表或图片没有明显信息作用；
- 讲稿过长或过短；
- 标题不够直接；
- 留白、对齐和密度可以进一步改善。

## 规则 ID

每条结果带稳定 ID，便于记录与例外声明：

| 前缀 | 覆盖范围 | 示例 |
| --- | --- | --- |
| `GEO.` | 几何与容器 | `GEO.OUTSIDE` 越界、`GEO.TABLE_WIDTH` 列宽不一致 |
| `TXT.` | 文字 | `TXT.OVERFLOW` 溢出、`TXT.OVERLAP` 重叠、`TXT.GLYPH` 字形缺失、`TXT.NOBREAK` 西文不可断行、`TXT.FULLWIDTH` 全角西文、`TXT.TABLE_OVERFLOW` 单元格撑高 |
| `TYP.` | 字号 | `TYP.MINSIZE` 字号过小、`TYP.LADDER` 不在字阶内、`TYP.STEPS` 一页档数过多、`TYP.INHERIT` 字号继承无法判定 |
| `PAL.` | 配色 | `PAL.CONTRAST` 对比度、`PAL.COLOR` 方案外颜色、`PAL.HUE` 色相组数、`PAL.AREA_CONTRAST` 对比色面积、`PAL.AREA_SUPPORT` 辅助色面积、`PAL.INHERIT` 继承色无法判定 |
| `FONT.` | 字体槽 | `FONT.EA` 中文字槽越界、`FONT.EA_MISSING` 缺中文槽、`FONT.LATIN` 西文槽越界、`FONT.LATIN_MISSING` 缺西文槽、`FONT.LATIN_CJK` 西文槽填了中文字体、`FONT.EA_BOLD` 黑体叠加粗体、`FONT.TABLE_KAI` 楷体入表格 |
| `LAY.` | 布局 | `LAY.REPEAT` 骨架重复、`LAY.PURETEXT` 纯文字页比例、`LAY.PURETEXT_FLAT` 纯文字页无排版层级、`LAY.BLOCKRATIO` 方块矩阵页占比、`LAY.MODEL` 关系模型覆盖、`LAY.MEANS` 缺少视觉手段 |
| `IMG.` | 配图 | `IMG.RES` 分辨率、`IMG.QUALITY` 画质、`IMG.MISSING` 缺失 |
| `MED.` | 媒体与对象 | `MED.SMARTART` 禁用对象 |
| `NOTE.` | 备注 | `NOTE.SHORT` 逐字稿过短 |

## 统一结果协议

每条结果至少包含：规则 ID、严重程度、页码或对象位置、实际值与预期值、违反的规则、修复建议、是否允许例外。用 `--json report.json` 导出机器可读结果，字段与上表一一对应。

## 例外机制

- 结构性例外用参数声明：沿用原模板保留 SmartArt 用 `--allow-smartart`。
- 页面级例外在备注写 `[例外:骨架重复]` 等标记并说明理由。
- 长期例外写进 `exceptions.json`，用 `--exceptions` 传入；命中例外的条目降级为"例外"并在报告中保留，不直接丢弃。
- 例外不会降级安全、版权与可读性规则：对比度、字号、缺失图片和禁用对象仍然报错。

## 检查顺序

1. 内容：数字、来源、口径、结论。
2. 结构：页面类型、关系模型、骨架、信息层级。
3. 可用性：溢出、越界、重叠、字体、颜色、对比度、图表。
4. 视觉：对齐、留白、密度、配图质量、模板一致性。
5. 兼容性：母版、动画、媒体、导出和 PowerPoint 实际显示。

## 工具

```text
uv run --python 3.14 --with 'python-pptx>=1.0,<2' python scripts/check_ppt.py deck.pptx --layout-report
uv run --python 3.14 --with 'python-pptx>=1.0,<2' python scripts/check_ppt.py deck.pptx --mode strict --json report.json
uv run --python 3.14 --with 'python-pptx>=1.0,<2' python scripts/preview_ppt.py deck.pptx preview.html
```

- `check_ppt.py` 是统一入口：调度 `audit_ppt.py`（几何、布局、配图、备注）与 `verify_ppt_style.py`（配色、对比度、字号、字体），去重后按错误、警告、建议和例外汇总。
- `palette_utils.py` 提供色阶与对比度自检，`--check` 可打印色阶对照表。
- `preview_ppt.py` 只负责渲染预览，不做风格判定。
- `run_regression.py` 是回归基线：正例必须零错误零警告，反例必须命中全部预期规则。
- 自动检查不能代替打开 PPT 后的视觉复核；布局分类是保守估算，最终以人工审阅为准。
- 用户反馈修改后，必须重新生成并重跑检查，而不是只重跑生成。

## 交付

- 修改已有 PPT 时输出新文件，不覆盖源文件。
- 先给预览或修改摘要，再交付文件。
- 不自动执行 Git add、commit、push 或分支切换。
- 交付说明列绝对路径、主要变化、验证结果、例外理由和仍需人工确认的风险。