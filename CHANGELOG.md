# Changelog

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
