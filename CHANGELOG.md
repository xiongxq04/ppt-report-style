# Changelog

## v0.4.1（2026-10-31）—— 新增配色预设 D + 外部色板库；修复冲突标记入库

### 修复
- **合并冲突标记被原样提交**：e57063a（merge 45a4df3 × 5fc39cf）把 `<<<<<<<` 冲突现场提交进 20 个文档与脚本，导致文档不可读、脚本不可运行。统一保留 v0.4.0（HEAD）侧——profiles.json 真相源即 v0.4.0 架构，v0.3.0 旧文档与其矛盾；全部文件自 45a4df3 恢复。
- 补删 `_commit_msg.txt` 草稿（v0.4.0 声称已删但仍入库）。

### 新增
- **配色预设 D「青绿循环」**（`--preset D`）：深蓝锚点 `26465E`（9.9:1）+ 青绿/暖黄/橙/砖红四档暖色系列（`2A9D8F / E9C46A / F4A261 / E76F51`），来自科研五色循环色板；hair/chart_neutral/accent_deep 由 `generate_scale` 从锚点派生并逐档核对。`profiles.json` 升至 0.4.1。
- **外部色板库 `references/palette-library.md`**：收录 7 组外部配色 + 十四宫格的逐色实算数据（HSL 结构、对白底 WCAG 对比度、文字/图形/填充判定、每组最佳用途），全部数值与 `palette_utils.py` 同源；含"近亲警告"（26465E/26445E/283E4E 视为同一深蓝墨）与四步"晋升规则"（选锚点→派生脚手架→选序列色→入册回归）。
- **预设名单全链路动态化**：`check_ppt.py` / `build_example.py` 的 `--preset` choices、`run_regression.py` 的预设冒烟循环改为从 `profile_lib.preset_names()` 读取——今后新增预设只改 profiles.json 一个文件。
- SKILL.md / palette-system.md / README / checking.md / plugin.json 同步为四套预设表述。

## v0.4.0（2026-10-31）—— 档位 × 预设架构 + 一致性大修

### 新增
- **三档严格度** `fast / balanced / strict`（`--tier`；旧 `--mode` 等价）：字号下限、大数字、结论行前缀、装饰预算、骨架重复、纯文字页比例、备注下限全部随档位分治。
- **三套配色预设** `A/B/C`（`--preset`）：Okabe-Ito 学术蓝（默认）/ 深青蓝 / 砖红，每套为完整 token 角色表，文字 token 全部 ≥4.5:1。
- **阈值唯一真相源** `references/profiles.json` + 加载器 `scripts/profile_lib.py`；文档只引用不重定义。
- 新机审规则：`TYP.POSTER`（非封面海报大字）、`PAL.ACCENT_PAGE`（点缀每页限额，按形状计）、`PAL.CHART_SERIES`（图表点亮超额）、`TXT.PREFIX`、`TXT.CAVEAT`、`EFF.SHADOW`、`EFF.STYLE_REF`、`IMG.RES`（按放置尺寸 dpi）、`LAY.REPEAT`/`LAY.PURETEXT`（档位分治）。
- 规则 ID 与 ERROR/WARNING/SUGGESTION/例外 分级统一进 `checking.md`；`--json` 报告字段对齐。
- `build_example.py --variant negative` 回归反例生成器；`--preset A|B|C` 正例预设切换。
- `eval-queries.json` 扩充至 30 条（新增档位/预设触发与负例）。

### 修复（v0.3 → v0.4 过渡期断裂）
- **check_ppt.py 接口断裂**：旧入口导入的 `format_findings/format_layout` 与参数签名已不存在，CI 必然 ImportError——三脚本统一为 (severity, slide, rule, kind, detail, extra) 六元组协议。
- **示例页数断言矛盾**：run_regression 断言 9 页而资产/生成器为 6 页——统一为 6 页。
- **D9D9D9 豁免位置盲洞**：原按颜色豁免 14pt 下限与对比度，正文染灰即可逃逸——改为按页眉/页脚几何区判定。
- **图表配色不查**：原只查图表文字色（defRPr），系列/点填充色漏检——补 `c:dPt`/`c:ser` 系列色检查。
- **docstring 与代码不一致**：`--sizes` 默认值说明与 argparse 实际值统一由 profiles.json 提供。
- 删除仓库根目录的 `_commit_msg.txt` 草稿；README 中英双语与实际架构同步（九页→6 页、51 条→30 条、旧配色描述清除）。

### 文档换代
- SKILL.md 增加"档位与配色预设"入口章节，验收清单全部挂规则 ID。
- 全部 references 头部加"阈值唯一权威 = profiles.json + SKILL.md"仲裁声明。
- workflow.md / style-rules.md / palette-system.md / layout-patterns.md / visual-means.md / effect-budget.md 按档位化重写，清除 v0.3 残留（页眉 16pt、左侧竖条、每页几何记忆点、#283E4E 独裁、六档字阶）。

### 验收
- `run_regression.py`：正例 balanced/strict 零错误零警告；预设 B/C 端到端冒烟全绿；反例命中全部 18 条预期规则（实际检出 21+ 条）；预览生成通过。
- CI：三档 × 预设机审矩阵 + 示例语义重建比对（文本/图表/点缀计数）+ 正反例回归。

## v0.3.0 及更早
见 git 历史（配色系列化 + 布局关系模型 + 配图必要性分档 + 效果预算；v0.2.0：双槽字体 + 六档字阶）。
