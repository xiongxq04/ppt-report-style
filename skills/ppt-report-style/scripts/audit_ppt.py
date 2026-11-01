# -*- coding: utf-8 -*-
"""结构审查：几何越界 / 文本溢出 / 重叠 / 字形与字体槽 / 阴影特效 / 配图分辨率 / 布局骨架 / 备注。

阈值唯一来源：references/profiles.json（经 profile_lib 加载）。
规则 ID 见 references/checking.md；本脚本负责 GEO.* / TXT.* / FONT.* / EFF.* / IMG.* / LAY.* / NOTE.*。
字号、文字配色、对比度、点缀限额、图表配色归 verify_ppt_style.py 管，本脚本不重复报。

用法: python audit_ppt.py <file.pptx> [更多文件 ...] --tier balanced|strict|fast
选项:
  --fonts/--latin-fonts  字体白名单覆盖（缺省取档位值）
  --notes-min N          每页备注最少字数覆盖（缺省取档位值）
  --overflow-tol         溢出容差倍数，默认 1.06
  --allow-smartart       保留模板 SmartArt 时放行 MED.SMARTART
退出码 0 = 无 ERROR（警告/建议/例外不影响退出码）。

中英混排三条硬道理（判定依据）：
1. 宽度按字符分类取 em 宽，宁可估宽不漏报；
2. 西文只在空格与中文处断行，长 token 单独报 TXT.NOBREAK；
3. 字体分两槽：a:latin 管西文、a:ea 管中文，两槽都要查。
"""
import sys
import os
import argparse
import unicodedata

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from pptx.util import Length
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

from profile_lib import (get_tier, get_preset, expand_font_whitelist, note_exception_rules,
                         note_exception_unknown, SAFETY_RULES, is_hf_zone, CJK_FACES)
from palette_utils import normalize_hex

EMU_IN = 914400.0
EMU_PT = 12700.0
NARROW_LATIN = set("ijlrtf|.,:;'\"()[]{}/!-\\` ")
WIDE_LATIN = set("mwMW@%&")


def is_cjk(ch):
    return unicodedata.east_asian_width(ch) in ("W", "F")


def char_em(ch):
    if is_cjk(ch):
        return 1.0
    if ch == " " or ch in NARROW_LATIN:
        return 0.30
    if ch in WIDE_LATIN:
        return 0.90
    if ch.isdigit():
        return 0.58
    if ch.isupper():
        return 0.30 if ch in "IL" else 0.72
    if ch.islower():
        return 0.56
    return 0.56


def tokens(s):
    out, cur = [], ""
    for ch in s:
        if is_cjk(ch):
            if cur:
                out.append(cur); cur = ""
            out.append(ch)
        elif ch == " ":
            out.append(cur + " "); cur = ""
        else:
            cur += ch
    if cur:
        out.append(cur)
    return out


def wrap_lines(s, size, avail_pt):
    if avail_pt <= 0:
        return 99, 0.0
    lines, cur, longest = 1, 0.0, 0.0
    for tk in tokens(s):
        w = sum(char_em(c) for c in tk) * size
        longest = max(longest, w)
        if cur > 0 and cur + w > avail_pt + 0.01:
            lines += 1
            cur = w
        else:
            cur += w
    return lines, longest


def slots(run):
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return (run.font.name or "", "")
    lat = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    return (
        (lat.get("typeface") if lat is not None else None) or (run.font.name or ""),
        ea.get("typeface") if ea is not None else "",
    )


def has_latin(s):
    return any(c.isalnum() and not is_cjk(c) for c in s)


def _classify_skeleton(text_boxes, charts, tables, pictures, sw, sh_):
    """保守估算页面骨架；分错的风险交给人工审阅（checking.md 约定）。"""
    if tables and any(t[2] > 0.65 * sw for t in tables):
        return "全宽单栏"
    if pictures or charts:
        for l, t, w, h in charts + pictures:
            if w > 0.30 * sw and (l > 0.45 * sw or l + w < 0.55 * sw):
                return "主图+侧栏"
    blocks = [b for b in text_boxes if b[2] > 1.0 and b[3] > 0.6]   # (l,t,w,h)
    if len(blocks) >= 4:
        xs = sorted({round(b[0] / 1.0) for b in blocks})
        ys = sorted({round(b[1] / 1.0) for b in blocks})
        if len(xs) >= 2 and len(ys) >= 2:
            return "网格/矩阵"
    left = [b for b in blocks if b[0] + b[2] / 2 < sw / 2]
    right = [b for b in blocks if b[0] + b[2] / 2 >= sw / 2]
    if left and right:
        lw = sum(b[2] for b in left)
        if 0.25 * sw < lw < 0.75 * sw:
            return "左右双栏"
    bands = {0: 0, 1: 0, 2: 0}
    for b in text_boxes:            # text_boxes: (l, t, w, h)
        idx = min(2, max(0, int(b[1] / (sh_ / 3))))
        bands[idx] += b[3]
    if sum(1 for v in bands.values() if v > 0.8) >= 3:
        return "上下分区"
    return "全宽单栏"


def _relation(tables, charts, blocks):
    if tables:
        return "矩阵"
    if charts:
        return "对比"
    if blocks >= 2:
        return "并列"
    return "其他"


def line_spacing_factor(paragraph, size):
    """段落的行距倍数。python-pptx 的 Length 是 int 子类，必须先于数值判断，
    否则「固定值行距」（如 Pt(24)）会被当成 30 万倍行距。"""
    spacing = paragraph.line_spacing
    if isinstance(spacing, Length):
        return spacing.pt / size if size else 1.0
    if isinstance(spacing, (int, float)):
        return max(spacing, 1.0)
    return 1.0


def run_color_hex(run):
    """run 的显式颜色（大写 HEX）；继承或主题色返回 None。"""
    try:
        return normalize_hex(str(run.font.color.rgb))
    except Exception:
        return None


def title_placeholder_size(slide):
    """标题占位符的生效字号：按 形状 → 版式 → 母版占位符 → 母版 txStyles 的顺序回落。"""
    try:
        title = slide.shapes.title
    except Exception:
        title = None
    if title is None:
        return None
    candidates = [title]
    try:
        idx = title.placeholder_format.idx
        for placeholder in slide.slide_layout.placeholders:
            if placeholder.placeholder_format.idx == idx:
                candidates.append(placeholder)
        for placeholder in slide.slide_layout.slide_master.placeholders:
            if placeholder.placeholder_format.idx == idx:
                candidates.append(placeholder)
    except Exception:
        pass

    def size_from_defrpr(element):
        if element is None:
            return None
        for defrpr in element.iter(qn("a:defRPr")):
            raw = defrpr.get("sz")
            if raw:
                return int(raw) / 100.0
        return None

    for candidate in candidates:
        try:
            for paragraph in candidate.text_frame.paragraphs:
                if paragraph.font.size:
                    return paragraph.font.size.pt
                for run in paragraph.runs:
                    if run.font.size:
                        return run.font.size.pt
        except Exception:
            pass
        try:
            found = size_from_defrpr(candidate._element.find(".//" + qn("a:lstStyle")))
            if found:
                return found
        except Exception:
            pass
    try:
        styles = slide.slide_layout.slide_master._element.find(qn("p:txStyles"))
        if styles is not None:
            found = size_from_defrpr(styles.find(qn("p:titleStyle")))
            if found:
                return found
    except Exception:
        pass
    return None


def placeholder_title_exempt(slide, page_h=7.5):
    """占位符版的封面/章节页判定：标题字号够大。

    封面（CENTER_TITLE 版式）整体豁免，允许作者/日期等附加文本框；
    其余标题页只要带正文占位符文字、或非占位符正文/图表/表格（HF 页眉页脚区除外），
    就按内容页处理，避免"仅标题版式 + 文本框正文"被误当封面。"""
    size = title_placeholder_size(slide)
    if size is None:
        return None
    try:
        title = slide.shapes.title
        is_cover = (title is not None
                    and title.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE)
    except Exception:
        is_cover = False
    if is_cover:
        return size
    try:
        for placeholder in slide.placeholders:
            try:
                ph_type = placeholder.placeholder_format.type
            except Exception:
                continue
            if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
                if placeholder.has_text_frame and placeholder.text_frame.text.strip():
                    return None
    except Exception:
        pass
    try:
        for shape in slide.shapes:
            if getattr(shape, "is_placeholder", False):
                continue
            if getattr(shape, "has_chart", False) or getattr(shape, "has_table", False):
                return None   # 带图表/表格的页面按内容页处理
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                top = shape.top / EMU_IN if shape.top is not None else None
                height = shape.height / EMU_IN if shape.height is not None else None
                if not is_hf_zone(top, height, page_h):
                    return None   # 非占位符正文（页脚区除外）
    except Exception:
        pass
    return size


def audit(path, tier="balanced", preset="A", fonts=None, latin_fonts=None,
          notes_min=None, tol=1.06, allow_smartart=False):
    """返回 (findings, layout_rows)。
    findings: (severity, slide_no, rule, kind, detail, extra)
    layout_rows: (slide_no, relation, skeleton, means, pure_text)"""
    T = get_tier(tier)
    P = get_preset(preset)
    # 封面/章节页豁免的标题字号档：取档位的海报阈值（fast 档为 null，按 36pt 兜底）
    exempt_pt = T["poster_pt"] if T["poster_pt"] is not None else 36
    issues = []
    layout_rows = []
    note_exc_by_slide = {}
    prs = Presentation(path)
    sw, sh_ = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
    ea_ok = expand_font_whitelist(fonts if fonts is not None else (T["fonts"] or []))
    lat_ok = expand_font_whitelist(latin_fonts if latin_fonts is not None else (T["latin_fonts"] or []))
    n_min = int(notes_min if notes_min is not None else T["notes_min"])
    note_sev = "ERROR" if tier == "strict" else ("WARNING" if tier == "balanced" else None)

    for si, slide in enumerate(prs.slides, 1):
        boxes = []
        tb_boxes, chart_boxes, table_boxes, pic_boxes = [], [], [], []
        graphic_shapes = 0
        smartart = 0
        max_size = 0.0
        page_sizes = set()
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        note_exc = note_exception_rules(notes)
        note_exc_by_slide[si] = note_exc
        for raw in note_exception_unknown(notes):
            issues.append(("SUGGESTION", si, "NOTE.EXCEPTION_UNKNOWN", "例外标记无法识别",
                           f"[例外:{raw}]", "规则 ID 见 checking.md；中文写法仅兼容『骨架重复』"))

        def add(sev, rule, kind, detail, extra=""):
            if rule in note_exc:      # 备注级例外：[例外:规则名] 降级为「例外」，保留可追溯
                if rule in SAFETY_RULES:
                    extra = ((extra + "；") if extra else "") + "例外声明被拒绝（安全规则不可例外）"
                else:
                    sev = "例外"
                    extra = ((extra + "；") if extra else "") + "备注声明 [例外:" + rule + "]"
            issues.append((sev, si, rule, kind, detail, extra))

        def check_runs(runs, where="", hf_zone=False):
            nonlocal max_size
            for r in runs:
                if not r.text.strip():
                    continue
                is_hf_run = hf_zone and run_color_hex(r) == P["hair"]
                if r.font.size and not is_hf_run:
                    max_size = max(max_size, r.font.size.pt)
                    page_sizes.add(r.font.size.pt)
                lat, ea = slots(r)
                if any(is_cjk(c) for c in r.text):
                    if not ea or (ea_ok and ea not in ea_ok):
                        add(T["font_slot_severity"], "FONT.EA", where + "中文字体", r.text[:16],
                            ea or "a:ea 未设置")
                if has_latin(r.text):
                    if lat in CJK_FACES:
                        add(T["font_slot_severity"], "FONT.LATIN_CJK", where + "西文用中文字体", r.text[:16],
                            f"latin={lat}")
                    elif not lat or (lat_ok and lat not in lat_ok):
                        add(T["font_slot_severity"], "FONT.LATIN", where + "西文字体", r.text[:16],
                            lat or "a:latin 未设置")
                for ch in r.text:
                    if ch in " \n\t":
                        continue
                    if "０" <= ch <= "９" or "Ａ" <= ch <= "Ｚ" or "ａ" <= ch <= "ｚ":
                        add("ERROR", "TXT.FULLWIDTH", where + "全角西文", repr(ch),
                            "字母与数字改半角；中文标点不受限")
                    try:
                        ch.encode("gbk")
                    except UnicodeEncodeError:
                        add("WARNING", "TXT.GLYPH", where + "字形缺失", repr(ch),
                            f"U+{ord(ch):04X} {unicodedata.name(ch, '?')} 不在 GBK，可能豆腐块")

        for shp in slide.shapes:
            _el = shp._element
            for _tag in ("a:outerShdw", "a:innerShdw", "a:prstShdw", "a:reflection", "a:glow"):
                if _el.find(".//" + qn(_tag)) is not None:
                    add("ERROR", "EFF.SHADOW", "阴影/特效", str(getattr(shp, "shape_type", shp)),
                        f"{_tag} 存在，必须删除")
                    break
            _st = str(getattr(shp, "shape_type", ""))
            if "AUTO_SHAPE" in _st or "TEXT_BOX" in _st:
                try:
                    if shp.shadow.inherit:
                        add("ERROR", "EFF.SHADOW", "阴影未显式关闭", str(shp.shape_type),
                            "生成时须 shadow.inherit=False（写入空 a:effectLst）")
                except Exception:
                    pass
                _style_el = _el.find(qn("p:style"))
                if _style_el is not None:
                    _eref = _style_el.find(qn("a:effectRef"))
                    if _eref is not None and _eref.get("idx", "0") not in ("0", "none"):
                        add("ERROR", "EFF.STYLE_REF", "主题特效引用", str(shp.shape_type),
                            f"effectRef idx={_eref.get('idx')} 指向主题外阴影，须删除 p:style")
            for graphic_data in _el.iter(qn("a:graphicData")):
                if "diagram" in (graphic_data.get("uri") or ""):      # SmartArt
                    smartart += 1
            if getattr(shp, "has_chart", False):
                try:
                    if "outerShdw" in shp.chart._chartSpace.xml:
                        add("ERROR", "EFF.SHADOW", "阴影/特效", "图表", "chartSpace 含 outerShdw")
                except Exception:
                    pass
            try:
                l, t = shp.left / EMU_IN, shp.top / EMU_IN
                w, h = shp.width / EMU_IN, shp.height / EMU_IN
                if (getattr(shp, "rotation", 0) or 0) == 0:
                    if l < -0.01 or t < -0.01 or l + w > sw + 0.01 or t + h > sh_ + 0.01:
                        add("ERROR", "GEO.OUTSIDE", "越界", f"{shp.shape_type} @({l:.2f},{t:.2f}) {w:.2f}x{h:.2f}",
                            f"页面 {sw:.2f}x{sh_:.2f}")
            except Exception:
                l = t = w = h = None

            if getattr(shp, "has_chart", False):
                chart_boxes.append((l or 0, t or 0, w or 0, h or 0))
                graphic_shapes += 1
            if getattr(shp, "has_table", False):
                table_boxes.append((l or 0, t or 0, w or 0, h or 0))
                graphic_shapes += 1
            if "PICTURE" in _st:
                pic_boxes.append((l or 0, t or 0, w or 0, h or 0))
                graphic_shapes += 1
                try:
                    px_w, px_h = shp.image.size
                except Exception:
                    add("WARNING", "IMG.MISSING", "图片数据无法读取",
                        f"{w:.2f}x{h:.2f}in" if w else str(shp.shape_type),
                        "链接图或文件损坏，分辨率不可判定；换内嵌图")
                else:
                    if w:
                        dpi = px_w / w
                        if dpi < 96:
                            add("ERROR", "IMG.RES", "配图分辨率严重不足", f"{px_w}px / {w:.2f}in",
                                f"有效约 {dpi:.0f} dpi < 96，按放置尺寸换清晰图")
                        elif dpi < 150:
                            add("WARNING", "IMG.RES", "配图分辨率偏低", f"{px_w}px / {w:.2f}in",
                                f"有效约 {dpi:.0f} dpi < 150")
                    long_edge = max(px_w, px_h)
                    if long_edge < 800:
                        add("WARNING", "IMG.QUALITY", "图片长边像素过低", f"{long_edge}px",
                            "放大后可能糊边；按放置尺寸换高清图")
            if "AUTO_SHAPE" in _st and getattr(shp, "fill", None) is not None:
                # 只有"有体量"的形状才算版面结构：面积 >0.3 in²，或填充/描边不是发丝灰。
                # 发丝线、页眉页脚横线太细，不构成视觉手段，否则纯文字页判定全部失效。
                try:
                    area = (w or 0) * (h or 0)
                    fill_hex = str(shp.fill.fore_color.rgb) if shp.fill.type is not None and shp.fill.type == 1 else None
                except Exception:
                    area, fill_hex = 0, None
                meaningful = area > 0.3 or (fill_hex and fill_hex.upper() not in (P["hair"], P["white"]))
                if meaningful:
                    graphic_shapes += 1

            if getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip():
                tf = shp.text_frame
                bw = shp.width / EMU_PT - (tf.margin_left + tf.margin_right) / EMU_PT
                bh = shp.height / EMU_PT - (tf.margin_top + tf.margin_bottom) / EMU_PT
                need = 0.0
                for p in tf.paragraphs:
                    runs = [r.text for r in p.runs]
                    if not runs:
                        continue
                    size = max((r.font.size.pt if r.font.size else 18) for r in p.runs)
                    ls = line_spacing_factor(p, size)
                    n, longest = wrap_lines("".join(runs), size, bw)
                    need += n * size * ls * 1.18
                    if longest > bw * tol:
                        add("ERROR", "TXT.NOBREAK", "西文不可断行", "".join(runs)[:30],
                            f"最长段 {longest:.0f}pt / 栏宽 {bw:.0f}pt")
                check_runs([r for p in tf.paragraphs for r in p.runs],
                           hf_zone=is_hf_zone(t, h, sh_))
                if bh > 0 and need > bh * tol:
                    add("ERROR", "TXT.OVERFLOW", "文本溢出", tf.text[:30].replace("\n", " "),
                        f"需 {need:.0f}pt / 框 {bh:.0f}pt")
                boxes.append((tf.text[:22].replace("\n", " "), l or 0, t or 0, w or 0, h or 0))
                tb_boxes.append((l or 0, t or 0, w or 0, h or 0))

            if getattr(shp, "has_table", False):
                tb = shp.table
                tw = sum(c.width for c in tb.columns) / EMU_IN
                if abs(tw - shp.width / EMU_IN) > 0.05:
                    add("ERROR", "GEO.TABLE_WIDTH", "表宽不符", f"{tw:.2f} vs {shp.width/EMU_IN:.2f}", "")
                colw = [c.width / EMU_PT for c in tb.columns]
                rowhs = [r.height / EMU_PT for r in tb.rows]
                for ri, row in enumerate(tb.rows):
                    for ci, cell in enumerate(row.cells):
                        tf = cell.text_frame
                        if not tf.text.strip():
                            continue
                        if getattr(cell, "is_spanned", False):
                            continue
                        span_w = getattr(cell, "span_width", 1) or 1
                        span_h = getattr(cell, "span_height", 1) or 1
                        avail_w = sum(colw[ci:ci + span_w]) - (cell.margin_left + cell.margin_right) / EMU_PT
                        avail_h = sum(rowhs[ri:ri + span_h]) - (cell.margin_top + cell.margin_bottom) / EMU_PT
                        need = 0.0
                        for p in tf.paragraphs:
                            if not p.runs:
                                continue
                            size = max((r.font.size.pt if r.font.size else 16) for r in p.runs)
                            ls = line_spacing_factor(p, size)
                            n, longest = wrap_lines("".join(r.text for r in p.runs), size, avail_w)
                            need += n * size * ls * 1.18
                            if longest > avail_w * tol:
                                add("ERROR", "TXT.NOBREAK", "表格西文不可断行", tf.text[:24],
                                    f"列{ci} 最长段 {longest:.0f}pt / {avail_w:.0f}pt")
                        if avail_h > 0 and need > avail_h * tol:
                            add("ERROR", "TXT.TABLE_OVERFLOW", "表格文本溢出", tf.text[:24],
                                f"需 {need:.0f}pt / 行高 {avail_h:.0f}pt（会撑高压到下方图形）")
                        check_runs([r for p in tf.paragraphs for r in p.runs], "表格")

        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ox = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
                oy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
                if ox > 0.02 and oy > 0.02:
                    area = ox * oy
                    small = min(a[3] * a[4], b[3] * b[4])
                    if small > 0 and area / small > 0.18:
                        add("ERROR", "TXT.OVERLAP", "文本重叠", f"[{a[0]}] x [{b[0]}]",
                            f"覆盖 {area/small*100:.0f}%")

        for pic in pic_boxes:      # 图片与文字重叠：占了小块面积 18% 以上提醒人工复核
            for b in tb_boxes:
                ox = min(pic[0] + pic[2], b[0] + b[2]) - max(pic[0], b[0])
                oy = min(pic[1] + pic[3], b[1] + b[3]) - max(pic[1], b[1])
                if ox > 0.02 and oy > 0.02:
                    area = ox * oy
                    small = min(pic[2] * pic[3], b[2] * b[3])
                    if small > 0 and area / small > 0.18:
                        add("WARNING", "GEO.OVERLAP_VISUAL", "图片与文字重叠",
                            f"@({pic[0]:.2f},{pic[1]:.2f}) {pic[2]:.2f}x{pic[3]:.2f}",
                            f"覆盖 {area / small * 100:.0f}%，人工复核")

        if smartart and not allow_smartart:
            add("ERROR", "MED.SMARTART", "禁用对象", f"{smartart} 个 SmartArt",
                "用普通形状/表格/图表表达；模板沿用需 --allow-smartart")

        # 封面/章节页（存在 ≥exempt_pt 标题）按实际讲解需要写备注，不适用字数下限；
        # 与 LAY.MEANS / 纯文字页比例的页型豁免同一判定口径。
        page_max = max(max_size, placeholder_title_exempt(slide, sh_) or 0)
        if note_sev and len(notes) < n_min and page_max < exempt_pt:
            add(note_sev, "NOTE.SHORT", "逐字稿过短", f"{len(notes)} 字 < {n_min}",
                "备注写约 1 分钟逐字稿 + 口径备答（封面/章节页豁免）")

        skeleton = _classify_skeleton(tb_boxes, chart_boxes, table_boxes, pic_boxes, sw, sh_)
        pure_text = not (chart_boxes or table_boxes or pic_boxes) and graphic_shapes == 0
        means = ("图表" if chart_boxes else "表格" if table_boxes else "图片" if pic_boxes
                 else "图形结构" if graphic_shapes else "纯排版")
        if pure_text and page_max < exempt_pt and len(page_sizes) < 2:
            add("SUGGESTION", "LAY.MEANS", "纯文字页无字号层级", f"仅 {len(page_sizes)} 档字号",
                "纯文字页至少两级字号拉开主次（visual-means.md）；封面/章节页豁免")
        relation = _relation(table_boxes, chart_boxes, len(tb_boxes))
        # 封面/章节页（存在 ≥36pt 标题）允许纯排版，不计入纯文字页比例
        layout_rows.append((si, relation, skeleton, means, pure_text, page_max >= exempt_pt))

    # ── LAY 规则（按档位分治）──
    if T["check_skeleton_repeat"] and len(prs.slides) >= T["skeleton_repeat_min_slides"]:
        run = 1
        for i in range(1, len(layout_rows) + 1):
            if i < len(layout_rows) and layout_rows[i][2] == layout_rows[i - 1][2]:
                run += 1
            else:
                if run >= 3:
                    if "LAY.REPEAT" in note_exc_by_slide.get(i, set()):
                        issues.append(("例外", i, "LAY.REPEAT", "骨架重复",
                                       f"连续 {run} 页「{layout_rows[i-1][2]}」",
                                       "备注声明 [例外:LAY.REPEAT]"))
                    else:
                        issues.append((T["skeleton_repeat_severity"], i, "LAY.REPEAT", "骨架重复",
                                       f"连续 {run} 页「{layout_rows[i-1][2]}」",
                                       "例外：逐步揭示同一图可写 [例外:LAY.REPEAT]（旧写法 [例外:骨架重复] 仍兼容）"))
                run = 1
    ratio_cap = T["pure_text_ratio"]
    if ratio_cap is not None:
        content = [(si, pt) for si, rel, sk, means, pt, is_title in layout_rows if not is_title]
        if content:
            pure = [si for si, pt in content if pt]
            if len(pure) / len(content) > ratio_cap:
                issues.append(("WARNING", 0, "LAY.PURETEXT", "纯文字页比例偏高",
                               f"{len(pure)}/{len(content)} 页纯文字 > {ratio_cap:.0%}",
                               f"纯文字页：{pure}"))
    layout_rows = [(si, rel, sk, means, pt) for si, rel, sk, means, pt, _ft in layout_rows]
    return issues, layout_rows


def format_findings(rows):
    lines = []
    for sev, si, rule, kind, detail, extra in rows:
        loc = f"p{si}" if si else "全片"
        tail = f" —— {extra}" if extra else ""
        lines.append(f"  [{sev}] {rule} {loc} {kind}: {detail}{tail}")
    return "\n".join(lines)


def format_layout(rows):
    lines = ["  页  关系  骨架  手段  纯文字"]
    for si, rel, sk, means, pt in rows:
        lines.append(f"  {si:>2}  {rel:<4}  {sk:<6}  {means:<6}  {'是' if pt else '否'}")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--tier", default="balanced", choices=("fast", "balanced", "strict"))
    ap.add_argument("--preset", default="A")
    ap.add_argument("--fonts", default=None)
    ap.add_argument("--latin-fonts", default=None)
    ap.add_argument("--notes-min", type=int, default=None)
    ap.add_argument("--overflow-tol", type=float, default=1.06)
    ap.add_argument("--allow-smartart", action="store_true")
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        issues, _ = audit(path, tier=a.tier, preset=a.preset, fonts=a.fonts,
                          latin_fonts=a.latin_fonts, notes_min=a.notes_min,
                          tol=a.overflow_tol, allow_smartart=a.allow_smartart)
        if issues:
            has_error = any(item[0] == "ERROR" for item in issues)
            if has_error:
                rc = 1
            print("FAIL" if has_error else "WARN", os.path.basename(path), f"({len(issues)} 项)")
            print(format_findings(issues[:40]))
            if len(issues) > 40:
                print("    ...另", len(issues) - 40, "项")
        else:
            print("OK  ", os.path.basename(path))
    return rc


if __name__ == "__main__":
    sys.exit(main())
