# 汇报 PPT 规范（ppt-report-style）

一套中文汇报 PPT 的排版规范 + 可执行的工作流。让 AI 生成的 PPT 在**字体、配色、版式骨架、图表与方块取舍、逐字稿**上稳定达标，而不是每次都返工。

<p align="center">
  <img src="skills/ppt-report-style/assets/example-slide.png" alt="示例：左文右图与几何记忆点" width="820">
</p>

<p align="center">
  <img src="skills/ppt-report-style/assets/example-table.png" alt="示例：强数据内容用表格承载" width="820">
</p>

> 示例页为纯占位内容（虚构数据），用于展示版式，不含任何真实项目材料。

## 解决什么问题

- 字体只用**宋体/楷体/黑体**，正文**最小 16pt**；文字只有**黑 / 白 / 主题色**三色（默认 `#283E4E`）。
- **强数据、强主题 → 表格**；叙事、机制、并列要点 → **方块/方格**，块内直接写内容，不加"关键指标"这类小标题。
- 指标名与数值写在**同一句**里，不做两层结构。
- 统一页面骨架：页眉 → 25pt 命题式标题 → 副标 → 细线 → 左文右图双栏 → 底部结论条。
- 每页一个**几何记忆点**（圆 / 三角 / 环形图 / 圆点线路 / 六边形…），同批页面各不相同但骨架一致。
- 每页备注写入**约 1 分钟逐字稿**，数据来源与口径只放备注不上片。
- 交付前**三道自检**：字号/配色机器校验 → HTML 预览图人眼检查 → 数字回溯原材料。

## 安装

**方式一：skill-installer（Codex 内）**

```
$skill-installer install https://github.com/YOUR_GITHUB_USERNAME/ppt-report-style/tree/main/skills/ppt-report-style
```

**方式二：手动**

```bash
git clone https://github.com/YOUR_GITHUB_USERNAME/ppt-report-style.git
cp -r ppt-report-style/skills/ppt-report-style ~/.codex/skills/
```

安装后重启 Codex 即可被识别。

## 使用

安装后无需显式点名。说"按我的规范做 PPT""把这个 PPT 重新排版""给这页配图表和逐字稿"等需求时会自动启用；也可显式写：

```
$ppt-report-style 帮我做一套关于××的汇报 PPT，10 页
```

## 脚本

| 脚本 | 用途 |
| --- | --- |
| `scripts/verify_ppt_style.py` | 校验 pptx 的字号下限、配色、字体白名单（含图表内文字） |
| `scripts/preview_ppt.py` | 把 pptx 渲染成 HTML 预览（形状、表格、图表近似渲染） |
| `scripts/build_example.py` | 重新生成 `assets/example-slide.pptx` 中性示例页 |

```bash
python3 -m pip install python-pptx
python scripts/verify_ppt_style.py 你的.pptx
python scripts/verify_ppt_style.py 你的.pptx --accent 1F4E79 --min-size 18
python scripts/preview_ppt.py 你的.pptx out.html --slide 0
```

预览图需要 Playwright + Chrome 截图；没有浏览器时把 HTML 直接打开即可。

## 可配置项

| 项 | 默认 | 改法 |
| --- | --- | --- |
| 主题色 | `#283E4E` | 生成脚本换常量；校验用 `--accent` |
| 最小字号 | 16pt | 校验用 `--min-size` |
| 字体白名单 | 宋体,楷体,黑体 | 校验用 `--fonts "宋体,黑体"` |
| 额外允许色 | 无 | 校验用 `--allow-colors` |

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
    ├── assets/                    # 中性示例页与预览图
    ├── references/
    │   ├── workflow.md            # 六阶段工作流与返工触发条件
    │   └── eval-queries.json      # 20 条触发测试用例
    └── scripts/                   # 校验 / 预览 / 示例生成
```

## FAQ

- **字体不全怎么办**：确认后可在 `--fonts` 里加入替代字体名。
- **图表预览与 PowerPoint 有差异**：预览是近似渲染，用于查溢出与对齐，最终以 PowerPoint 为准。
- **想改整页骨架**：`references/workflow.md` 里写了每段的坐标与间距基准。

## 许可证

MIT，见 [LICENSE](LICENSE)。

## 参考

- [Agent Skills 开放规范](https://agentskills.io)
- [Codex 插件与技能文档](https://developers.openai.com/codex/plugins/build)
