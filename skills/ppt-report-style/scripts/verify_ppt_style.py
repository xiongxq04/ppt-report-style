# -*- coding: utf-8 -*-
"""风格校验：字阶、文字配色与对比度、点缀色限额、图表配色、双字体槽、前缀与口径。

阈值唯一来源：references/profiles.json（经 profile_lib 加载）。
规则 ID 见 references/checking.md；本脚本负责 TYP.* / PAL.* / FONT.* 三组。

用法:
  python verify_ppt_style.py <file.pptx> [more.pptx ...] --tier balanced|strict|fast --preset A|B|C|D
  python verify_ppt_style.py deck.pptx --tier strict --preset A
选项（全部可选，缺省取档位/预设值）:
  --min-size / --sizes / --max-steps / --fonts / --latin-fonts / --allow-colors / --accent
  --allow-cjk-latin   容忍西文槽填中文字体（只用于跑旧稿）
退出码 0 = 全部通过。
"""
import sys
import os
import argparse

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

from profile_lib import get_tier, get_preset, preset_colors, accent_family, is_hf_zone
from palette_utils import contrast_ratio, hue_group, is_neutral, normalize_hex

CJK_FACES = {"宋体", "楷体", "黑体", "SimSun", "KaiTi", "SimHei", "MS Mincho", "Microsoft YaHei"}
PREFIXES = ("结论：", "判断：", "收束：")
CAVEAT_WORDS = ("虚构", "示例")
EMU_IN = 914400.0
BASE_NEUTRAL = {"000000", "FFFFFF"}


def _finding(severity, slide, rule, kind, detail, extra=""):
    return (severity, slide, rule, kind, detail, extra)


def has_cjk(s):
    return any("　" <= c <= "鿿" or "＀" <= c <= "￯" for c in s)


def has_latin(s):
    return any((c.isalpha() or c.isdigit()) and not ("　" <= c <= "鿿") for c in s)


def slots(run):
    """返回 (latin 槽, ea 槽)；未显式设置的返回空串。"""
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return (run.font.name or "", "")
    lat = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    return (
        (lat.get("typeface") if lat is not None else None) or (run.font.name or ""),
        ea.get("typeface") if ea is not None else "",
    )


def _run_color(run):
    try:
        raw = run.font.color.rgb
        return normalize_hex(str(raw))
    except Exception:
        return None


def _shape_uses_accent(shape, accent_set):
    """形状（含表格/图表）是否使用了点缀色——按形状计 1 处，不按 run 数。"""
    if getattr(shape, "has_text_frame", False):
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                col = _run_color(r)
                if col and col in accent_set:
                    return True
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        col = _run_color(r)
                        if col and col in accent_set:
                            return True
    if getattr(shape, "has_chart", False):
        try:
            xml = shape.chart._chartSpace.xml.upper()
            if any(("#" + c) in xml or c in xml for c in accent_set):
                return True
        except Exception:
            pass
    return False


def _chart_accent_marks(shape, accent_set):
    """图表里被点缀色点亮的标记数：逐点(dPt)填充逐个计；无点缀点填充时系列底色算 1。"""
    marks = 0
    try:
        cs = shape.chart._chartSpace
        for ser in cs.iter(qn("c:ser")):
            dpt_accent = 0
            for dpt in ser.iter(qn("c:dPt")):
                for sc in dpt.iter(qn("a:srgbClr")):
                    if (sc.get("val") or "").upper() in accent_set:
                        dpt_accent += 1
                        break
            marks += dpt_accent
            if dpt_accent:
                continue  # 点级已点亮时，系列底色不重复计
            for sc in ser.iter(qn("a:srgbClr")):
                if (sc.get("val") or "").upper() in accent_set:
                    marks += 1
                    break
    except Exception:
        return 0
    return marks


def check(path, tier="balanced", preset="A",
          allow=None, fonts=None, latin_fonts=None, sizes=None,
          min_size=None, max_steps=None, allow_cjk_latin=False,
          accent=None, max_per_rule=3):
    """返回 findings 列表：(severity, slide_no, rule, kind, detail, extra)。"""
    T = get_tier(tier)
    P = dict(get_preset(preset))
    if accent:                       # --accent 临时覆盖预设点缀色
        P["accent"] = normalize_hex(accent, P["accent"])
    allowed = preset_colors(P) | set(BASE_NEUTRAL) | {c.upper() for c in (allow or []) if c.strip()}
    accent_set = accent_family(P)
    ea_ok = {f for f in (fonts if fonts is not None else (T["fonts"] or [])) if f}
    lat_ok = {f for f in (latin_fonts if latin_fonts is not None else (T["latin_fonts"] or [])) if f}
    size_ok = {float(s) for s in (sizes if sizes is not None else (T["sizes"] or []))} or None
    floor = float(min_size if min_size is not None else T["min_size"])
    hf_floor = float(T["hf_min_size"])
    steps_cap = int(max_steps if max_steps is not None else (T["max_steps"] or 0)) or None
    err = "ERROR"
    findings = []

    prs = Presentation(path)
    n_slides = len(prs.slides)
    deck_text = []

    for si, slide in enumerate(prs.slides, 1):
        steps = set()
        accent_shapes = 0
        page_colors = set()

        def add(sev, rule, kind, detail, extra=""):
            findings.append(_finding(sev, si, rule, kind, detail, extra))

        def walk_runs(runs, top_in=None, h_in=None, in_table=False):
            nonlocal accent_shapes
            for r in runs:
                if not r.text.strip():
                    continue
                deck_text.append(r.text)
                sz = r.font.size.pt if r.font.size else None
                col = _run_color(r)
                hf_ok = (col == P["hair"]) and is_hf_zone(top_in, h_in)
                local_floor = hf_floor if hf_ok else floor
                if sz is None or sz < local_floor:
                    add(err, "TYP.MINSIZE", "字号过小", r.text[:20],
                        f"{sz}pt < {local_floor}pt" + ("（页眉页脚档）" if hf_ok else ""))
                if sz is not None:
                    steps.add(sz)
                    if size_ok and sz not in size_ok and not hf_ok:
                        add(err, "TYP.LADDER", "字号不在字阶内", r.text[:20],
                            f"{sz}pt，字阶={sorted(size_ok)}")
                    poster = T["poster_pt"]
                    if poster and si != 1 and sz > poster:
                        add(err, "TYP.POSTER", "海报大字", r.text[:16],
                            f"{sz}pt > {poster}pt（非封面页；关键数字句内黑体强调）")
                if col is None:
                    if not hf_ok:
                        add("WARNING", "PAL.INHERIT", "颜色继承无法判定", r.text[:16], "run 未显式设色")
                elif col not in allowed:
                    add(err, "PAL.COLOR", "方案外颜色", r.text[:16], col)
                else:
                    if not is_neutral(col):
                        page_colors.add(col)
                    if col in accent_set:
                        pass  # 点缀按形状计数，不在此处计
                    elif not hf_ok and col != P["white"]:
                        ratio = contrast_ratio(col, "FFFFFF")
                        if ratio < 4.5:
                            add(err, "PAL.CONTRAST", "对比度不足", r.text[:16],
                                f"{col} {ratio:.2f}:1 < 4.5:1")
                lat, ea = slots(r)
                if has_cjk(r.text):
                    if not ea:
                        add(T["font_slot_severity"], "FONT.EA", "中文字体未设置", r.text[:16], "a:ea 空，继承主题字体")
                    elif ea_ok and ea not in ea_ok:
                        add(T["font_slot_severity"], "FONT.EA", "中文字体越界", r.text[:16], ea)
                if has_latin(r.text):
                    if not lat:
                        add(T["font_slot_severity"], "FONT.LATIN", "西文字体未设置", r.text[:16], "a:latin 空")
                    elif lat in CJK_FACES:
                        if not allow_cjk_latin:
                            add(T["font_slot_severity"], "FONT.LATIN_CJK", "西文用中文字体", r.text[:16],
                                f"latin={lat}（英文将用中文字形渲染）")
                    elif lat_ok and lat not in lat_ok:
                        add(T["font_slot_severity"], "FONT.LATIN", "西文字体越界", r.text[:16], lat)
                if ea == "黑体" and r.font.bold:
                    add(T["black_bold_severity"], "FONT.EA_BOLD", "黑体叠加粗体", r.text[:16], "合成粗会糊边；强调改换字体或点缀色")
                if in_table and ea == "楷体":
                    add(err, "FONT.TABLE_KAI", "楷体入表格", r.text[:16], "表格只用宋体/黑体")

        for shp in slide.shapes:
            top_in = shp.top / EMU_IN if shp.top is not None else None
            h_in = shp.height / EMU_IN if shp.height is not None else None
            if _shape_uses_accent(shp, accent_set):
                accent_shapes += 1
            if getattr(shp, "has_text_frame", False):
                walk_runs([r for p in shp.text_frame.paragraphs for r in p.runs], top_in, h_in)
            if getattr(shp, "has_table", False):
                for row in shp.table.rows:
                    for cell in row.cells:
                        walk_runs([r for p in cell.text_frame.paragraphs for r in p.runs],
                                  None, None, in_table=True)
            if getattr(shp, "has_chart", False):
                try:
                    cs = shp.chart._chartSpace
                    for rPr in cs.iter(qn("a:defRPr")):
                        sz_raw = rPr.get("sz")
                        if sz_raw and int(sz_raw) < floor * 100:
                            add(err, "TYP.CHARTSIZE", "图表字号过小", "", int(sz_raw) / 100)
                        sf = rPr.find(qn("a:solidFill"))
                        if sf is not None:
                            sc = sf.find(qn("a:srgbClr"))
                            if sc is not None:
                                val = normalize_hex(sc.get("val"))
                                if val and val not in allowed:
                                    add(err, "PAL.CHARTCOLOR", "图表颜色越界", "", val)
                    marks = _chart_accent_marks(shp, accent_set)
                    limit = T["chart_accent_marks"]
                    if limit is not None and marks > limit:
                        add(err if tier == "strict" else "WARNING", "PAL.CHART_SERIES",
                            "图表点亮超额", f"{marks} 处点缀 > {limit}", "只点亮最重要的一根柱/一行")
                except Exception:
                    pass

        limit = T["accent_per_page"]
        if limit is not None and accent_shapes > limit:
            sev = err if tier == "strict" else "WARNING"
            add(sev, "PAL.ACCENT_PAGE", "点缀色超额", f"{accent_shapes} 个形状 > {limit}",
                "点缀只标图表重点/表格关键行/当前章，不悬空当广告牌")
        hue_max = T["hue_groups_max"]
        if hue_max is not None:
            groups = {hue_group(c) for c in page_colors}
            if len(groups) > hue_max:
                add("WARNING", "PAL.HUE", "色相组过多", f"{len(groups)} 组 > {hue_max}",
                    "非中性色相 ≤3 组")
        if steps_cap and len(steps) > steps_cap:
            add(err, "TYP.STEPS", "字号层级过多", f"{len(steps)} 档 > {steps_cap}", sorted(steps))

    joined = "".join(deck_text)
    if T["ban_prefixes"]:
        for kw in PREFIXES:
            n = joined.count(kw)
            if n:
                findings.append(_finding(err, 0, "TXT.PREFIX", "结论行前缀", kw,
                                         f"出现 {n} 次，应为 0（直接写判断句）"))
    caveat_cap = T["caveat_max"]
    if caveat_cap is not None:
        for kw in CAVEAT_WORDS:
            n = joined.count(kw)
            if n > caveat_cap:
                findings.append(_finding(err, 0, "TXT.CAVEAT", "口径声明超量", kw,
                                         f"「{kw}」出现 {n} 次 > {caveat_cap}（封面副标 + 数据页表注两处声明位）"))

    # 同规则同页折叠
    folded, seen = [], {}
    for item in findings:
        key = (item[1], item[2])
        seen[key] = seen.get(key, 0) + 1
        if seen[key] <= max_per_rule:
            folded.append(item)
    return folded


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--tier", default="balanced", choices=("fast", "balanced", "strict"))
    ap.add_argument("--preset", default="A")
    ap.add_argument("--min-size", type=float, default=None)
    ap.add_argument("--fonts", default=None, help="中文字体（a:ea）白名单覆盖")
    ap.add_argument("--latin-fonts", default=None, help="西文字体（a:latin）白名单覆盖")
    ap.add_argument("--allow-colors", default="")
    ap.add_argument("--allow-cjk-latin", action="store_true")
    ap.add_argument("--sizes", default=None, help="字阶覆盖，逗号分隔")
    ap.add_argument("--max-steps", type=int, default=None)
    ap.add_argument("--accent", default=None, help="临时覆盖预设点缀色")
    ap.add_argument("--max-per-rule", type=int, default=3)
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    allow = [c.strip() for c in a.allow_colors.split(",") if c.strip()]
    sizes = [float(s) for s in a.sizes.split(",")] if a.sizes else None
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        bad = check(path, tier=a.tier, preset=a.preset, allow=allow, fonts=a.fonts,
                    latin_fonts=a.latin_fonts, sizes=sizes, min_size=a.min_size,
                    max_steps=a.max_steps, allow_cjk_latin=a.allow_cjk_latin,
                    accent=a.accent, max_per_rule=a.max_per_rule)
        if bad:
            rc = 1
            print("FAIL", os.path.basename(path), f"({len(bad)} 项)")
            for b in bad[:40]:
                print("   ", b)
            if len(bad) > 40:
                print("    ...另", len(bad) - 40, "项")
        else:
            print("OK  ", os.path.basename(path))
    return rc


if __name__ == "__main__":
    sys.exit(main())
