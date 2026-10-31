# PPT 检查与交付规则

规则阈值只在 `references/profiles.json` 与 SKILL.md 定义，检查脚本实现这些阈值；本文件规定分级、协议和交付方式。严重度按档位分治（fast 多数降为建议；balanced/strict 从严），同一规则在不同档位可能是 ERROR 或 WARNING——以脚本输出为准。

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
| `TXT.` | 文字 | `TXT.OVERFLOW` 溢出、`TXT.OVERLAP` 重叠、`TXT.GLYPH` 字形缺失、`TXT.NOBREAK` 西文不可断行、`TXT.FULLWIDTH` 全角西文、`TXT.TABLE_OVERFLOW` 单元格撑高、`TXT.PREFIX` 结论行前缀、`TXT.CAVEAT` 口径声明超量 |
| `TYP.` | 字号 | `TYP.MINSIZE` 字号过小、`TYP.LADDER` 不在字阶内、`TYP.STEPS` 一页档数过多、`TYP.POSTER` 非封面海报大字、`TYP.CHARTSIZE` 图表字号过小 |
| `PAL.` | 配色 | `PAL.CONTRAST` 对比度、`PAL.COLOR` 方案外颜色、`PAL.CHARTCOLOR` 图表颜色、`PAL.ACCENT_PAGE` 点缀每页超额、`PAL.CHART_SERIES` 图表点亮超额、`PAL.HUE` 色相组数、`PAL.INHERIT` 继承色无法判定 |
| `FONT.` | 字体槽 | `FONT.EA` 中文字槽越界/缺失、`FONT.LATIN` 西文槽越界/缺失、`FONT.LATIN_CJK` 西文槽填了中文字体、`FONT.EA_BOLD` 黑体叠加粗体、`FONT.TABLE_KAI` 楷体入表格 |
| `LAY.` | 布局 | `LAY.REPEAT` 骨架重复（档位分治）、`LAY.PURETEXT` 纯文字页比例（档位分治） |
| `IMG.` | 配图 | `IMG.RES` 按放置尺寸分辨率不足（<96dpi ERROR，<150dpi WARNING） |
| `EFF.` | 效果 | `EFF.SHADOW` 阴影/特效、`EFF.STYLE_REF` 主题特效引用（effectRef） |
| `MED.` | 媒体与对象 | `MED.SMARTART` 禁用对象 |
| `NOTE.` | 备注 | `NOTE.SHORT` 逐字稿过短（档位分治） |

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
python scripts/check_ppt.py deck.pptx --tier balanced --preset A --layout-report
python scripts/check_ppt.py deck.pptx --tier strict --json report.json
python scripts/preview_ppt.py deck.pptx preview.html --slide 0
python run_regression.py          # 回归：正例零错误零警告；反例命中全部预期规则
```

- `check_ppt.py` 是统一入口：调度 `audit_ppt.py`（几何、布局、配图、备注）与 `verify_ppt_style.py`（配色、对比度、字号、字体），去重后按错误、警告、建议和例外汇总；档位经 `--tier`（旧参数 `--mode` 等价）传给两个子脚本。
- `palette_utils.py --check` 做预设对比度自检；`preview_ppt.py` 只负责渲染预览，不做风格判定。
- `run_regression.py` 是回归基线：正例（6 页学术示例）在 balanced 与 strict 下必须零错误零警告，三套配色预设各做一次端到端冒烟，反例必须命中全部预期规则（见脚本内 `EXPECTED_NEGATIVE_RULES`）。
- 自动检查不能代替打开 PPT 后的视觉复核；布局分类是保守估算，最终以人工审阅为准。
- 用户反馈修改后，必须重新生成并重跑检查，而不是只重跑生成。

## 交付

- 修改已有 PPT 时输出新文件，不覆盖源文件。
- 先给预览或修改摘要，再交付文件。
- 不自动执行 Git add、commit、push 或分支切换。
- 交付说明列绝对路径、主要变化、验证结果、例外理由和仍需人工确认的风险。