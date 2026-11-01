# Changelog

## v0.4.5（2026-11-01）—— 机审误报/漏报与例外机制收口

> 背景：两份独立审查交叉验证（全库逐行 + 最小样本实测），修复 7 处硬缺陷，并对例外机制、文档口径与回归覆盖做结构性补强。

### 硬缺陷
- **SmartArt 检测恢复**：`.//p:graphicFrame` 只搜后代，顶层 SmartArt 永不命中（`MED.SMARTART` 空转）→ 改为 `_el.iter(a:graphicData)`，覆盖顶层与组内对象；反例补伪 SmartArt 页。
- **固定值行距不再误判**：python-pptx 的 `Length` 是 int 子类，`Pt(24)` 被当成 30 万倍行距 → 先判 `Length` 再换算倍数（audit 形状/表格两处）。
- **白底卡片底色修正**：显式白色填充不再被当作"无填充"，底层填充搜索不再跳过白色形状。
- **表格单元格底色参与判定**：cell fill 作为文字底色，并纳入点缀配额与色相统计。
- **点缀色文字不再豁免对比度**：accent/accent_deep 与普通文字同按实际底色判 4.5:1。
- **骨架分类全宽表判据修正**：宽度 `t[2]` 取代高度 `t[3]`。
- **字号统计口径**：HF 页眉页脚字号不计入 `TYP.STEPS` / `LAY.MEANS`；NOTE/MEANS 的页型豁免改取 `poster_pt`（fast 兜底 36），并支持占位符字号按版式/母版回落。
- **字体集合收敛到 profile_lib**：audit/verify 不再各持一份；`FONT.TABLE_KAI` 兼容 KaiTi；补 仿宋/FangSong/MS Mincho 判定。

### 例外机制
- 新增规则 ID 注册表（40 个，含 `NOTE.EXCEPTION_UNKNOWN`）；备注标记与 `--exceptions` JSON 均按注册表校验。
- 规范写法为 `[例外:规则ID]`；兼容 `[例外:骨架重复]` 与大小写；无法识别的标记报 `NOTE.EXCEPTION_UNKNOWN`（SUGGESTION），不再静默失效。
- **安全规则不可例外**：`EFF.SHADOW`、`EFF.STYLE_REF`、`GEO.OUTSIDE`、`IMG.MISSING`、`IMG.RES`、`MED.SMARTART`、`PAL.CONTRAST`、`TXT.NOBREAK`、`TXT.OVERFLOW`、`TXT.TABLE_OVERFLOW`、`TYP.MINSIZE`；备注与 JSON 两条路径均保留 ERROR 并注明「例外声明被拒绝」。
- 同页同规则折叠移到 `check_ppt.py` 合并后统一执行，audit 侧不再漏网；汇总行的「例外」计数改为包含备注例外。
- 未知 JSON 规则 ID 直接报错退出 2。

### 其他
- 预览：CSS `font` 简写修正、图片 MIME 按实际类型、页码越界友好报错、docstring 色值校准；`check_ppt --preview-out` 传递预设 accent。
- CLI：`--preset` 大小写统一；`audit_ppt.py` 退出码改为仅 ERROR 阻断。
- 文档口径：checking 例外机制/严重度、SKILL NOTE 豁免与 36pt/口径分档、layout/effect 例外写法与骨架 ≥6 页、workflow 备注豁免、README 四预设与阈值真相源表述。
- 回归：新增 guard 正例变体（固定行距/白卡/深底表头/点缀图表）与 safety 反例（安全例外被拒）；反例 EXPECTED 增补 `MED.SMARTART`；6 步回归全绿。

### 验证
- `run_regression.py` 全绿：正例 + 变体零误报、反例 19 条预期规则全命中、安全例外拒绝、预览生成通过。
- 最小样本复跑：行距/白卡/表格不再误报，accent 压浅灰卡报 3.67:1，SmartArt 检出。

## v0.4.4（2026-10-31）—— 规则体系冲突清零：9 处前后矛盾逐项裁定

> 背景：对规则体系自身做哲学级审计（宣称 ↔ 宣称、宣称 ↔ 实现），实证坐实 2 处硬冲突（预设 D 色觉宣称、`area_quotas` 死阈值），另清理 7 处适用域缺失与措辞漂移。修法原则：以 v0.4 机制为准，清除旧宣称；文档写清适用域而不是继续叠规则。

### 硬冲突（宣称与现实正面矛盾）
- **「四套色盲安全预设」为不实宣称**：README、plugin.json、palette-system 改为分级表述——A 有 Okabe-Ito 文献背书；B 靠明度阶梯区分；C/D 是通用序列色、未按色觉缺陷专项设计。实测 D 的 `E9C46A`（暖黄）与 `F4A261`（橙）在红绿色盲模拟下近乎同色（CIE76/CIEDE2000 实测 ΔE≈2–3，远低于可区分阈值）。palette-system 新增「色觉缺陷适用性」一节给出各预设用法与兜底手段；palette-library 的选材指引同步改为优先 A、D 需人工复核。
- **删除 `area_quotas` 死字段**：profiles.json 三档的 `area_quotas` 无任何脚本消费，checking.md 还把它当警告项宣传——v0.4 已改为「每页点缀限额」，阈值与文档残留一并清除。
- **3:1 大字档退出机审**：palette-system 明确机审统一 4.5:1；`RATIO_LARGE` 降级为仅人工判读大字与图形边界，`--contrast` 输出改口径，不再出现「仅大字达标」与机审两套答案。
- **NOTE.SHORT 按页型豁免**：封面/目录/章节页（≥36pt 标题）按实际讲解需要写备注，不适用字数下限；与 LAY.MEANS、纯文字页比例的页型豁免共用同一判定口径。checking.md 规则表同步。

### 软冲突（缺适用域导致的互相打架）
- style-rules 的对比度 P0 表述补上 profiles.json 声明的 HF 页眉页脚唯一例外。
- 主视觉形状表加适用域：仅 fast/balanced 且形状必须承载语义；strict 只允许功能性几何标记。layout-patterns 形状节同步。
- SKILL 骨架坐标标注为「默认参数、非强制」，与 style-rules「结论行不必固定底部」调和。
- checking.md 首句 fast 档概括改为准确描述：只有字体槽降警告、黑体加粗降建议，安全类规则不降级。
- SKILL「三道自检」统一为四道（结构/机器/视觉/内容），与 workflow 阶段 5 一致。

### 附带漂移
- README 与 plugin.json 的预设数量/描述校准（补 D、去「全部通过」式笼统保证）；README 回归矩阵表述校准。

### 验证
- `run_regression.py` 全绿（正例 balanced/strict 零错误零警告、A/B/C/D 端到端冒烟通过、反例命中全部预期规则）。
- CI 等价步骤全绿：frontmatter、四预设自检、示例 deck 三档 verify/audit/check、示例语义重建一致。
- 新增行为单独验证：清空备注后封面（44pt）与章节页（36pt）不再报 NOTE.SHORT，正文页（30pt）仍报。

## v0.4.3（2026-10-31）—— 全库审计修复：机审盲洞补齐 + 文档漂移清零

> 背景：对 SKILL.md、9 个 references、8 个脚本、CI 做全量交叉审计（文档 ↔ 代码 ↔ profiles.json 三方），按审计结论逐项修复。

### 修复（功能性 bug）
- **`--fonts`/`--latin-fonts` 逗号串从未被拆分**：白名单被逐字迭代成单字集合，传参与默认等价的白名单也会炸出 FONT.EA。`profile_lib.expand_font_whitelist()` 统一拆分 + 中英文别名展开（宋体↔SimSun、黑体↔SimHei 等），audit/verify/check 三入口全修。
- **备注级例外 `[例外:规则名]` 落地**：此前只存在于提示文字，从未被解析。现在 `profile_lib.note_exception_rules()` 在 audit/verify 两端解析备注标记，命中规则的发现降级为「例外」并保留可追溯（LAY.REPEAT 同样支持）；全片级规则（TXT.PREFIX/CAVEAT/LAY.PURETEXT）仍走 `--exceptions` JSON。
- **check_ppt 合并去重确定性化**：audit 与 verify 对同一发现的两种措辞按 (页码, 规则, 对象) 去重，替代原先的 set() 无序合并。
- **HF 页眉页脚几何区按页高等比缩放**：非标画布不再失准（`is_hf_zone` 增加 page_h 参数）。

### 机审盲洞补齐
- **PAL.CONTRAST 按实际底色判定**：此前只对纯白底算，淡卡片上的文字对比度无人查。现在取文字形状自身 solid 填充为底色，否则回溯 z 序下方最近填充形状；文字压图片上底色不可判定时跳过（由 GEO.OVERLAP_VISUAL 兜底）。白字对白底不再豁免（1:1 必报）。
- **PAL.HUE 纳入形状填充与图表系列色**：此前只统计文字色，彩虹图表页照样过「色相 ≤3 组」。
- **幽灵规则落地**：`IMG.MISSING`（图片数据无法读取，WARNING）、`IMG.QUALITY`（长边 <800px，WARNING）、`GEO.OVERLAP_VISUAL`（图片压文字 >18%，WARNING）、`LAY.MEANS`（纯文字页仅一档字号，SUGGESTION；封面/章节页豁免）——四条文档承诺、脚本未实现，现已实现并回写 checking.md 规则表。
- **字号继承降为警告**：run 未显式设字号从 TYP.MINSIZE ERROR 改为 `TYP.INHERIT` WARNING，与 PAL.INHERIT（颜色继承警告）口径一致，模板跟随场景不再被打穿。
- **黑体加粗检查覆盖 SimHei**；中文字体白名单接受中英文名。

### 文档漂移清零
- SKILL.md 硬性限制节声明"色值为预设 A 举例，换预设整体替换"，图表配色处改为预设相对表述；验收清单补 TYP.INHERIT、PAL.CONTRAST 实际底色口径。
- style-rules.md「默认门槛 150 字/页」改为按档位（balanced 90 / strict 150）；字阶"6 档/7 档"口径统一。
- workflow.md 与 checking.md 的 git 矛盾调和：提交前说明改动、push 需用户明确要求；workflow 预设表述补 D。
- workflow.md 阶段 0「换 B/C」补 D；阶段 4 token 硬编码改预设相对表述。
- checking.md 例外机制如实描述备注标记与折叠规则（同页同规则 ≤3 条，`--max-per-rule`）。
- visual-means.md LAY.MEANS 措辞与实现对齐。

### 其他
- **CI**：预设自检改为动态遍历全部预设；示例语义比对中的点缀色从 profiles.json 读取（不再硬编码 0072B2）。
- **palette_utils.py 瘦身**：删除 v0.3 遗留的 THEME/CONTRAST/SUPPORT 色阶、`family_of()` 模糊匹配与死代码 `report()`；`--scale` 派生输出增加文字可用性标注。
- **eval-queries.json** 扩至 32 条（补预设 D 与色板库触发用例）。
- 审计结论中「LAY.REPEAT 报错页码 off-by-one」经推演复核为**误报**（原实现报告的是重复段最后一页），未改动。

## v0.4.2（2026-10-31）—— layout-patterns.md 并入学术海报版式的可迁移经验

### 新增
- `references/layout-patterns.md` 新增"海报版式的可迁移经验"一节：方法论四条（字号下限 = 观看距离的函数、单页单视觉主角、阅读顺序显式引导、先内容减法再版式）+ 海报六骨架与现有骨架库的对照表 + 5 条人工验收清单。来源：埃米编辑知乎文章（2024-11-14，p/6846316845）及其 9 张版式实例与线框图（含规则三分法、Animate Your Science 侧栏轨、Mike Morrison #betterposter）。

### 边界（刻意不做）
- 只收可迁移项，海报专有内容（规格尺寸/横竖版/分辨率）不进本规范；SKILL.md description 维持"不用于海报"，海报需求仍需点名，`eval-queries.json` 负例不受影响。
- 不新增机审规则与阈值，不改动 profiles.json；新增验收项全部为人工项。

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
