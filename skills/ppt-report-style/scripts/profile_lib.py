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

# 可判定为"中文字形"的字体集合（audit / verify 共用，避免两处副本分叉）
CJK_FACES = {"宋体", "楷体", "黑体", "微软雅黑", "仿宋",
             "SimSun", "KaiTi", "SimHei", "Microsoft YaHei", "FangSong", "MS Mincho"}

# 备注级例外标记：[例外:规则名]（checking.md 例外机制）
NOTE_EXCEPTION_RE = re.compile(r"\[例外[:：]\s*([^\]]+?)\s*\]")

# 规则 ID 注册表：备注例外、--exceptions JSON、文档对账共用一套词表
RULE_IDS = frozenset({
    "EFF.SHADOW", "EFF.STYLE_REF",
    "FONT.EA", "FONT.EA_BOLD", "FONT.LATIN", "FONT.LATIN_CJK", "FONT.TABLE_KAI",
    "GEO.OUTSIDE", "GEO.OVERLAP_VISUAL", "GEO.TABLE_WIDTH",
    "IMG.MISSING", "IMG.QUALITY", "IMG.RES",
    "LAY.MEANS", "LAY.PURETEXT", "LAY.REPEAT",
    "MED.SMARTART",
    "NOTE.SHORT", "NOTE.EXCEPTION_UNKNOWN",
    "PAL.ACCENT_PAGE", "PAL.CHARTCOLOR", "PAL.CHART_SERIES", "PAL.COLOR",
    "PAL.CONTRAST", "PAL.HUE", "PAL.INHERIT",
    "TXT.CAVEAT", "TXT.FULLWIDTH", "TXT.GLYPH", "TXT.NOBREAK", "TXT.OVERFLOW",
    "TXT.OVERLAP", "TXT.PREFIX", "TXT.TABLE_OVERFLOW",
    "TYP.CHARTSIZE", "TYP.INHERIT", "TYP.LADDER", "TYP.MINSIZE", "TYP.POSTER",
    "TYP.STEPS",
})

# 安全规则：例外机制不得降级（checking.md 例外机制）
SAFETY_RULES = frozenset({
    "EFF.SHADOW", "EFF.STYLE_REF",
    "GEO.OUTSIDE",
    "IMG.MISSING", "IMG.RES",
    "MED.SMARTART",
    "PAL.CONTRAST",
    "TXT.NOBREAK", "TXT.OVERFLOW", "TXT.TABLE_OVERFLOW",
    "TYP.MINSIZE",
})

# 备注例外的中文兼容别名（规范写法仍是规则 ID）
NOTE_ALIASES = {"骨架重复": "LAY.REPEAT"}


def normalize_rule_id(raw):
    """把例外标记里的名字规范化为规则 ID；无法识别返回 None。"""
    text = str(raw or "").strip()
    if not text:
        return None
    if text in RULE_IDS:
        return text
    if text in NOTE_ALIASES:
        return NOTE_ALIASES[text]
    upper = text.upper()
    if upper in RULE_IDS:
        return upper
    return None


def note_exception_rules(notes_text):
    """解析备注里的 [例外:规则名] 标记，返回已知规则 ID 集合。"""
    rules = set()
    for raw in NOTE_EXCEPTION_RE.findall(notes_text or ""):
        rule = normalize_rule_id(raw)
        if rule:
            rules.add(rule)
    return rules


def note_exception_unknown(notes_text):
    """返回无法识别的例外标记原文列表（用于给出 SUGGESTION，避免静默失效）。"""
    unknown = []
    for raw in NOTE_EXCEPTION_RE.findall(notes_text or ""):
        if normalize_rule_id(raw) is None:
            unknown.append(str(raw).strip())
    return unknown


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


def tier_names():
    return list(load_profiles()["tiers"])


def preset_names():
    return list(load_profiles()["presets"])
