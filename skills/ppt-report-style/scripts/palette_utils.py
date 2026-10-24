# -*- coding: utf-8 -*-
"""配色系统工具：中性色判定、主题色阶、对比色、辅助色、WCAG 对比度与色相分组。

权威定义见 references/palette-system.md；本模块是它的可执行实现。

自检:
  python palette_utils.py --check
其他用法:
  python palette_utils.py --contrast 283E4E FFFFFF
  python palette_utils.py --scale 1F4E79
  python palette_utils.py --report 283E4E
"""
import argparse
import colorsys
import sys

sys.stdout.reconfigure(encoding="utf-8")

# —— 配色方案（与 references/palette-system.md 保持一致）——
ACCENT_DEFAULT = "283E4E"
THEME_SCALE = ("142129", "1E303B", "283E4E", "4A6577", "8199A8", "B7C7D2", "E4EBF0")
CONTRAST_SCALE = ("8C5A16", "B4771F", "C98B2E", "E0B15B")
SUPPORT_SCALE = ("2F4E48", "4F7C74", "8DB0A8", "D6E5E1")
NEUTRAL_KNOWN = ("000000", "1A1A1A", "333333", "595959", "808080", "A6A6A6", "D9D9D9",
                 "F1F5F8", "F2F2F2", "FFFFFF")

FAMILIES = ("theme", "contrast", "support")
FAMILY_LABEL = {"neutral": "中性色", "theme": "主题色", "contrast": "对比色",
                "support": "辅助色", "unknown": "方案外颜色"}

NEUTRAL_MAX_SPREAD = 8      # 同一颜色 rgb 最大最小差 <= 8 视为灰阶
FAMILY_TOLERANCE = 60.0     # 与色阶成员的 RGB 距离容差
MAX_HUE_GROUPS = 3          # 单页允许的色相组数（不含黑白灰）

RATIO_BODY = 4.5            # 正文对比度下限
RATIO_LARGE = 3.0           # 大字与图形边界对比度下限
LARGE_PT = 18.0             # WCAG 大字口径
LARGE_BOLD_PT = 14.0

ACCENT_AREA_WARN = 0.10     # 对比色面积提醒线
ACCENT_AREA_MAX = 0.20
SUPPORT_AREA_WARN = 0.20    # 辅助色面积提醒线
SUPPORT_AREA_MAX = 0.35


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


def _distance(first, second):
    a, b, c = hex_to_rgb(first)
    x, y, z = hex_to_rgb(second)
    return ((a - x) ** 2 + (b - y) ** 2 + (c - z) ** 2) ** 0.5


def family_of(value, accent=None, contrast=None, support=None):
    raw = normalize_hex(value)
    if raw is None:
        return "unknown"
    if is_neutral(raw):
        return "neutral"
    palette = {
        "theme": set(THEME_SCALE) | ({normalize_hex(accent, ACCENT_DEFAULT)} if accent else set()),
        "contrast": set(CONTRAST_SCALE) | ({normalize_hex(contrast)} if contrast else set()),
        "support": set(SUPPORT_SCALE) | ({normalize_hex(support)} if support else set()),
    }
    for name in FAMILIES:
        if raw in palette[name]:
            return name
    best, best_distance = "unknown", None
    for name in FAMILIES:
        for member in palette[name]:
            distance = _distance(raw, member)
            if best_distance is None or distance < best_distance:
                best, best_distance = name, distance
    return best if best_distance is not None and best_distance <= FAMILY_TOLERANCE else "unknown"


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
    family = family_of(raw)
    hue, saturation, lightness = hsl(raw)
    return (f"#{raw} {FAMILY_LABEL[family]} 色相 {hue:.0f}度 饱和 {saturation:.2f} 明度 {lightness:.2f} "
            f"白底对比 {contrast_ratio(raw, 'FFFFFF'):.2f}:1")


def report(accent=ACCENT_DEFAULT):
    lines = [f"主题色：#{normalize_hex(accent, ACCENT_DEFAULT)}",
             "主题色阶：" + " ".join(f"#{item}" for item in THEME_SCALE),
             "对比色阶：" + " ".join(f"#{item}" for item in CONTRAST_SCALE),
             "辅助色阶：" + " ".join(f"#{item}" for item in SUPPORT_SCALE),
             "中性色　：黑白灰不计入配色额度",
             "",
             f"{'色值':<9}{'族':<7}{'白底对比':>9}{'正文4.5':>9}{'大字3.0':>9}"]
    for item in list(THEME_SCALE) + list(CONTRAST_SCALE) + list(SUPPORT_SCALE):
        lines.append(f"#{item:<8}{FAMILY_LABEL[family_of(item)]:<7}"
                     f"{contrast_ratio(item, 'FFFFFF'):>8.2f}"
                     f"{'是' if text_safe(item) else '否':>8}"
                     f"{'是' if contrast_ratio(item, 'FFFFFF') >= RATIO_LARGE else '否':>8}")
    return "\n".join(lines)


def custom_scale_report(base):
    scale = generate_scale(base)
    lines = [f"从 #{normalize_hex(base)} 派生的 {len(scale)} 档色阶："]
    for item in scale:
        lines.append("  " + describe(item))
    return "\n".join(lines)


def self_check():
    assert abs(contrast_ratio("000000", "FFFFFF") - 21.0) < 0.01
    assert abs(contrast_ratio("283E4E", "283E4E") - 1.0) < 0.001
    assert is_neutral("F1F5F8") and is_neutral("000000") and is_neutral("808080")
    assert not is_neutral("283E4E") and not is_neutral("E0B15B")
    assert family_of("283E4E") == "theme"
    assert family_of("4A6577") == "theme"
    assert family_of("C98B2E") == "contrast"
    assert family_of("4F7C74") == "support"
    assert family_of("FF00FF") == "unknown"
    print("配色工具自检通过")
    print(report())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="运行自检并打印配色表")
    parser.add_argument("--report", nargs="?", const=ACCENT_DEFAULT, help="打印配色表（可指定主题色）")
    parser.add_argument("--contrast", nargs=2, metavar=("前景", "背景"))
    parser.add_argument("--scale", metavar="基色")
    args = parser.parse_args()

    if args.contrast:
        ratio = contrast_ratio(args.contrast[0], args.contrast[1])
        print(f"#{normalize_hex(args.contrast[0])} 与 #{normalize_hex(args.contrast[1])} 对比度 {ratio:.2f}:1")
        print("正文达标" if ratio >= RATIO_BODY else ("仅大字达标" if ratio >= RATIO_LARGE else "不达标"))
        return 0
    if args.scale:
        print(custom_scale_report(args.scale))
        return 0
    if args.report:
        print(report(args.report))
        return 0
    self_check()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())