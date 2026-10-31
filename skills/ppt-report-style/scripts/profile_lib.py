# -*- coding: utf-8 -*-
"""档位（tier）与配色预设（preset）加载器：全部脚本的阈值唯一入口。

真相源：references/profiles.json。本模块只读取与解析，不定义任何阈值。
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROFILES_PATH = os.path.join(_HERE, "..", "references", "profiles.json")

_CACHE = None

# 页眉页脚极浅灰档的几何判定区（单位 in，13.33x7.5 画布）
HF_ZONE_TOP_MAX = 0.70     # top + height ≤ 0.70 视为页眉区
HF_ZONE_TOP_MIN = 7.00     # top ≥ 7.00 视为页脚区


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


def is_hf_zone(top_in, height_in):
    if top_in is None:
        return False
    return (top_in + (height_in or 0)) <= HF_ZONE_TOP_MAX or top_in >= HF_ZONE_TOP_MIN


def tier_names():
    return list(load_profiles()["tiers"])


def preset_names():
    return list(load_profiles()["presets"])
