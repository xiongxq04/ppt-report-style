# -*- coding: utf-8 -*-
"""档位（tier）与配色预设（preset）加载器：全部脚本的阈值唯一入口。

真相源：references/profiles.json。本模块只读取与解析，不定义任何阈值。
"""
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROFILES_PATH = os.path.join(_HERE, "..", "references", "profiles.json")

_CACHE = None

# 页眉页脚极浅灰档的几何判定区（基准画布 13.33x7.5 in，非标画布按页高等比缩放）
HF_ZONE_TOP_MAX = 0.70     # top + height ≤ 0.70 视为页眉区
HF_ZONE_TOP_MIN = 7.00     # top ≥ 7.00 视为页脚区

# 中文常用字体的中英文名等价对：PPT 文件里两者都可能出现，白名单必须同时放行
FONT_ALIASES = {"宋体": "SimSun", "黑体": "SimHei", "楷体": "KaiTi",
                "微软雅黑": "Microsoft YaHei", "仿宋": "FangSong"}

# 备注级例外标记：[例外:规则名]（checking.md 例外机制）
NOTE_EXCEPTION_RE = re.compile(r"\[例外[:：]\s*([A-Z]{2,8}\.[A-Z_]{2,})\s*\]")


def load_profiles():
    global _CACHE
    if _CACHE is None:
        with open(_PROFILES_PATH, "r", encoding="utf-8") as handle:
            _CACHE = json.load(handle)
    return _CACHE


def get_tier(name):
    tiers = load_profiles()["tiers"]
    if name not in tiers:
        raise SystemExit(f"未知档位 {name!r}，可选：{', '.join(tiers)}")
    return tiers[name]


def get_preset(name):
    presets = load_profiles()["presets"]
    key = str(name or "A").upper()
    if key not in presets:
        raise SystemExit(f"未知配色预设 {name!r}，可选：{', '.join(presets)}")
    return presets[key]


def preset_colors(preset):
    """预设的全部颜色 token（大写 hex，不含 series）。"""
    p = preset
    return {p["ink"], p["muted"], p["hair"], p["chart_neutral"],
            p["accent"], p["accent_deep"], p["white"]}


def preset_all_colors(preset):
    """预设的全部颜色 token 含 series（大写 hex）。"""
    return preset_colors(preset) | {c.upper() for c in preset.get("series", [])}


def accent_family(preset):
    """计入“点缀色限额”的颜色：accent + accent_deep。"""
    return {preset["accent"].upper(), preset["accent_deep"].upper()}


def is_hf_zone(top_in, height_in, page_h_in=7.5):
    scale = (page_h_in or 7.5) / 7.5
    if top_in is None:
        return False
    return ((top_in + (height_in or 0)) <= HF_ZONE_TOP_MAX * scale
            or top_in >= HF_ZONE_TOP_MIN * scale)


def expand_font_whitelist(names):
    """把字体白名单按中英文别名展开：写 '宋体' 也放行 'SimSun'，反之亦然。
    同时接受逗号分隔字符串（CLI 传参场景）或可迭代对象。"""
    if names is None:
        return None
    if isinstance(names, str):
        names = names.split(",")
    out = set()
    for name in names:
        name = str(name).strip()
        if not name:
            continue
        out.add(name)
        for zh, en in FONT_ALIASES.items():
            if name == zh:
                out.add(en)
            elif name == en:
                out.add(zh)
    return out


def note_exception_rules(notes_text):
    """解析备注里的 [例外:规则名] 标记，返回规则 ID 集合。"""
    return set(NOTE_EXCEPTION_RE.findall(notes_text or ""))


def tier_names():
    return list(load_profiles()["tiers"])


def preset_names():
    return list(load_profiles()["presets"])
