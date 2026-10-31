# 配色系统（预设 A/B/C/D）

> 阈值唯一权威是 `references/profiles.json` 与 SKILL.md；本文件只做展开说明，不重定义任何色值与限额。
> 可执行实现：`scripts/palette_utils.py`（对比度/色相/派生）与 `scripts/verify_ppt_style.py`（PAL.* 规则）。

## 四套预设，每套是完整的 token 角色表

| Token | 作用 | 文字可用？ |
| --- | --- | --- |
| `ink` 墨 | 正文、标题 | 是（21:1） |
| `muted` 次级灰 | 图注、口径、副标、落款 | 是（5.1:1） |
| `hair` 发丝线 | 页眉页脚线、分隔线；**页眉页脚文字专用档**（仅页眉区/页脚区位置享受 14pt+对比度豁免） | 仅 HF 区 |
| `chart_neutral` 图表灰 | 未点亮的柱/系列 | 否（只做图形） |
| `accent` 点缀 | 图表重点、表格关键行、当前章、正文关键数字 | 是（≥5.2:1） |
| `accent_deep` 点缀深 | 小字号强调（若需要） | 是 |
| `series` 序列色 | 多系列图表的第二、三色（色盲安全组合） | 否（图表专用） |

| 预设 | 定位 | accent | 出处 |
| --- | --- | --- | --- |
| `A`（默认） | 学术蓝 | `0072B2` | Okabe & Ito Color Universal Design；Wong, Nature Methods 2011 |
| `B` | 深青蓝 | `283E4E` | v0.3 主题色系保留（深档） |
| `C` | 砖红 | `A63603` | ColorBrewer 暖调（红紫/橄榄组合） |
| `D` | 青绿循环（深蓝锚点·暖色系列） | `26465E` | 科研五色循环色板（见 palette-library.md 组6）；脚手架档由 `--scale` 从 accent 派生 |

**换预设 = 整套换**：`--preset B` 一次替换全部 token，不接受只换点缀色的"半套配色"。
需要自定义品牌色时（模板跟随场景），用 `palette_utils.py --scale <基色>` 派生同色相色阶，
再用 `--contrast 前景 背景` 逐档核对比度；文字 token 必须 ≥4.5:1。

## 硬门槛（全部档位通用）

- 正文文字对比度 **≥4.5:1**（hair 的 HF 页眉页脚档豁免，且必须位于页眉/页脚几何区）。
- 非 token 色出现在文字上即报 `PAL.COLOR`；图表文字色同样受限（`PAL.CHARTCOLOR`）。
- 点缀色每页限额按档位（strict ≤2 / balanced ≤3，按形状计，`PAL.ACCENT_PAGE`）。
- 图表被点亮的标记 ≤1 处（`PAL.CHART_SERIES`）；"只点亮一根"必须建立在截断轴诚实标注、差异可见之上。
- 非中性色相 ≤3 组/页（`PAL.HUE`，WARNING）。
- 浅档只做填充、底色和分隔，不承载正文小字。

## 自检命令

```text
python scripts/palette_utils.py --check                # 自检 + 预设 A 对比度表
python scripts/palette_utils.py --check --preset C     # 预设 C 对比度表
python scripts/palette_utils.py --contrast 0072B2 FFFFFF
python scripts/palette_utils.py --scale 1F4E79         # 自定义基色派生色阶（模板跟随用）
```

## 常见错误

- 只换色相不看明度：色相差 180 度不代表对比度达标。
- 浅档当正文色，或在浅底上写浅字。
- 把 `chart_neutral`/`series` 当文字色用。
- 需要第二点缀色时自造颜色，而不是从预设 series 里取。
- 一页塞进四组以上色相，视觉上变成彩虹页。
