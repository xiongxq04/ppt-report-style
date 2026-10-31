# -*- coding: utf-8 -*-
"""配色系统工具：对比度、色相分组、色阶派生与预设自检。

预设配色（A/B/C/D）的权威定义在 references/profiles.json；本模块是它的可执行实现，
另保留 generate_scale() 作为自定义色阶派生工具（模板跟随与色板库晋升场景用）。

自检:
  python palette_utils.py --check            # 自检 + 预设 A 对比度表
  python palette_utils.py --check --preset B
其他用法:
  python palette_utils.py --contrast 0072B2 FFFFFF
  python palette_utils.py --scale 1F4E79
"""
import argparse
import colorsys
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from profile_lib import get_preset, preset_names, preset_all_colors

# —— 默认点缀色（= 预设 A 的 accent）——
ACCENT_DEFAULT = "0072B2"

NEUTRAL_KNOWN = ("000000", "1A1A1A", "333333", "595959", "808080", "A6A6A6", "D9D9D9",
                 "F1F5F8", "F2F2F2", "FFFFFF")

NEUTRAL_MAX_SPREAD = 8      # 同一颜色 rgb 最大最小差 <= 8 视为灰阶
MAX_HUE_GROUPS = 3          # 单页允许的色相组数（不含黑白灰）

RATIO_BODY = 4.5            # 正文对比度下限（机审唯一口径，大字同样适用）
RATIO_LARGE = 3.0           # 仅人工判读大字与图形边界用，不参与机审放行
LARGE_PT = 18.0             # WCAG 大字口径
LARGE_BOLD_PT = 14.0


def normalize_hex(value, default=None):
    raw = str(value or "").strip().lstrip("#").upper()
    if len(raw) == 6 and all(ch in "0123456789ABCDEF" for ch in raw):
        return raw
    return default


def hex_to_rgb(value):
    raw = normalize_hex(value, "000000")
    return int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16)


def rgb_to_hex(red, green, blue):
    clamp = lambda v: max(0, min(255, int(round(v))))
    return f"{clamp(red):02X}{clamp(green):02X}{clamp(blue):02X}"


def hsl(value):
    red, green, blue = hex_to_rgb(value)
    hue, lightness, saturation = colorsys.rgb_to_hls(red / 255.0, green / 255.0, blue / 255.0)
    return hue * 360.0, saturation, lightness


def hue_group(value):
    return int(hsl(value)[0] / 30.0) % 12


def spread(value):
    red, green, blue = hex_to_rgb(value)
    return max(red, green, blue) - min(red, green, blue)


def is_neutral(value):
    raw = normalize_hex(value)
    if raw is None:
        return False
    if raw in NEUTRAL_KNOWN:
        return True
    return spread(raw) <= NEUTRAL_MAX_SPREAD


def relative_luminance(value):
    channels = []
    for channel in hex_to_rgb(value):
        srgb = channel / 255.0
        channels.append(srgb / 12.92 if srgb <= 0.04045 else ((srgb + 0.055) / 1.055) ** 2.4)
    red, green, blue = channels
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def contrast_ratio(foreground, background):
    first = relative_luminance(foreground)
    second = relative_luminance(background)
    lighter, darker = max(first, second), min(first, second)
    return (lighter + 0.05) / (darker + 0.05)


def composite(foreground, background, alpha=1.0):
    """把带透明度的前景色叠加到背景色上，返回实际呈现的颜色。"""
    try:
        alpha = float(alpha)
    except (TypeError, ValueError):
        alpha = 1.0
    alpha = max(0.0, min(1.0, alpha))
    fr, fg, fb = hex_to_rgb(foreground)
    br, bg, bb = hex_to_rgb(background)
    return rgb_to_hex(fr * alpha + br * (1 - alpha),
                      fg * alpha + bg * (1 - alpha),
                      fb * alpha + bb * (1 - alpha))


def required_ratio(size_pt=None, bold=False):
    if size_pt is None:
        return RATIO_BODY
    if size_pt >= LARGE_PT or (bold and size_pt >= LARGE_BOLD_PT):
        return RATIO_LARGE
    return RATIO_BODY


def contrast_ok(foreground, background, size_pt=None, bold=False):
    return contrast_ratio(foreground, background) >= required_ratio(size_pt, bold)


def text_safe(value, background="FFFFFF"):
    return contrast_ratio(value, background) >= RATIO_BODY


def generate_scale(base, steps=7):
    """按明度阶梯从基色派生色阶；步数为奇数，中间一档就是基色本身。"""
    base_hex = normalize_hex(base)
    if base_hex is None:
        raise ValueError("需要 6 位十六进制颜色")
    if steps % 2 == 0:
        steps += 1
    hue, saturation, lightness = hsl(base_hex)
    mid = steps // 2
    ceiling = 0.94
    floor = max(0.08, min(lightness, lightness - (ceiling - lightness)))
    scale = []
    for index in range(steps):
        if index == mid:
            scale.append(base_hex)
            continue
        offset = (index - mid) / float(mid)
        if offset > 0:
            target = lightness + offset * (ceiling - lightness)
        else:
            target = lightness + offset * (lightness - floor)
        target = max(0.0, min(1.0, target))
        keep = saturation if abs(offset) < 0.99 else saturation * 0.85
        red, green, blue = colorsys.hls_to_rgb(hue / 360.0, target, max(0.0, min(1.0, keep)))
        scale.append(rgb_to_hex(red * 255, green * 255, blue * 255))
    return tuple(scale)


def describe(value):
    raw = normalize_hex(value, "000000")
    hue, saturation, lightness = hsl(raw)
    tag = " 中性" if is_neutral(raw) else ""
    return (f"#{raw}{tag} 色相 {hue:.0f}度 饱和 {saturation:.2f} 明度 {lightness:.2f} "
            f"白底对比 {contrast_ratio(raw, 'FFFFFF'):.2f}:1"
            f"{' ≥4.5 可作文字' if text_safe(raw) else ' 仅图形/填充'}")


def custom_scale_report(base):
    scale = generate_scale(base)
    lines = [f"从 #{normalize_hex(base)} 派生的 {len(scale)} 档色阶："]
    for item in scale:
        lines.append("  " + describe(item))
    return "\n".join(lines)


def preset_report(name="A"):
    p = get_preset(name)
    lines = [f"预设 {name}：{p['label']}", f"出处：{p['source']}",
             f"{'角色':<14}{'色值':<10}{'白底对比':>9}{'正文4.5':>9}"]
    roles = [("墨 ink", p["ink"]), ("次级灰 muted", p["muted"]), ("发丝线 hair", p["hair"]),
             ("图表灰 chart_neutral", p["chart_neutral"]), ("点缀 accent", p["accent"]),
             ("点缀深 accent_deep", p["accent_deep"])]
    for label, value in roles:
        lines.append(f"{label:<14}#{value:<10}{contrast_ratio(value, 'FFFFFF'):>8.2f}"
                     f"{'是' if text_safe(value) else '否':>8}")
    for value in p.get("series", []):
        lines.append(f"{'序列 series':<14}#{value:<10}{contrast_ratio(value, 'FFFFFF'):>8.2f}"
                     f"{'是' if text_safe(value) else '否':>8}")
    return "\n".join(lines)


def self_check(name="A"):
    assert abs(contrast_ratio("000000", "FFFFFF") - 21.0) < 0.01
    assert is_neutral("F1F5F8") and is_neutral("000000") and is_neutral("808080")
    assert not is_neutral("0072B2") and not is_neutral("A63603")
    # 全部预设的文字 token 必须达到正文对比度门槛（hair/white 除外）
    for key in preset_names():
        p = get_preset(key)
        for role in ("ink", "muted", "accent", "accent_deep"):
            assert text_safe(p[role]), f"预设 {key} 的 {role} #{p[role]} 对比度不足"
    print("配色工具自检通过")
    print(preset_report(name))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="运行自检并打印预设对比度表")
    parser.add_argument("--preset", default="A", choices=tuple(preset_names()))
    parser.add_argument("--report", nargs="?", const=ACCENT_DEFAULT, help="打印预设表（旧参数兼容）")
    parser.add_argument("--contrast", nargs=2, metavar=("前景", "背景"))
    parser.add_argument("--scale", metavar="基色")
    args = parser.parse_args()

    if args.contrast:
        ratio = contrast_ratio(args.contrast[0], args.contrast[1])
        print(f"#{normalize_hex(args.contrast[0])} 与 #{normalize_hex(args.contrast[1])} 对比度 {ratio:.2f}:1")
        print("正文达标（机审口径 ≥4.5:1）" if ratio >= RATIO_BODY else
              ("仅达 3:1 人工判读档（大字/图形边界），机审不放行"
               if ratio >= RATIO_LARGE else "不达标"))
        return 0
    if args.scale:
        print(custom_scale_report(args.scale))
        return 0
    if args.report:
        print(preset_report(args.preset))
        return 0
    self_check(args.preset)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())