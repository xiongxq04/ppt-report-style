# -*- coding: utf-8 -*-
"""汇报 PPT 结构审查：越界 / 文本溢出 / 西文断行 / 重叠 / 字形与全角 / 备注 / 表格 /
禁用对象 / 布局骨架 / 配图质量。

字号、配色与字体槽的判定在 verify_ppt_style.py，用 check_ppt.py 汇总，两边不重复同一规则。

用法: python audit_ppt.py <file.pptx> [更多文件 ...] [选项]
选项:
  --min-size N        最小字号(pt)，默认 16（只用于溢出估算）
  --notes-min N       每页备注最少字数，默认 150
  --overflow-tol      溢出容差倍数，默认 1.06
  --mode balanced|strict  平衡档把备注、骨架重复、配图分辨率等列为警告；严格档列错误
  --allow-smartart    沿用原模板时放行已有 SmartArt
  --layout-report     打印每页的关系模型与页面骨架分类
退出码 0 = 没有错误。

中英混排的三条硬道理（本脚本据此判定）：
1. 宽度不能按"西文一律 0.52 倍"估：Arial 里 i/l 约 0.3 em，M/W 约 0.9 em，全大写缩写比小写宽得多。
2. 西文不能逐字断行：换行只发生在空格与中文之间，长 token 比栏宽还长时必然溢出或挤到下一栏。
3. 内嵌图片要按放置尺寸核算有效分辨率，长边像素过小的图放大后必然糊。
"""
import argparse
import math
import os
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from palette_utils import ACCENT_DEFAULT

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn
except ImportError:
    print("需要 python-pptx：pip install python-pptx")
    sys.exit(2)

EMU_IN = 914400.0
EMU_PT = 12700.0
SEVERITY_ORDER = {"ERROR": 0, "WARNING": 1, "SUGGESTION": 2, "例外": 3}
CJK_FACES = {"宋体", "楷体", "黑体", "SimSun", "KaiTi", "SimHei", "Microsoft YaHei"}
NARROW_LATIN = set("ijlrtf|.,:;'\"()[]{}/!-\\` ")
WIDE_LATIN = set("mwMW@%&")
EXCEPTION_TAGS = ("[例外", "【例外")
PURE_TEXT_LIMIT = 0.20
BLOCK_LIMIT = 0.35
MIN_RELATION_MODELS = 4
IMAGE_DPI_WARN = 150.0
IMAGE_DPI_MIN = 96.0
IMAGE_EDGE_WARN = 800

ARROW_SHAPES = {getattr(MSO_SHAPE, item) for item in
                ("CHEVRON", "PENTAGON", "RIGHT_ARROW", "LEFT_ARROW", "UP_ARROW", "DOWN_ARROW",
                 "LEFT_RIGHT_ARROW", "STRIPED_RIGHT_ARROW", "STRIPED_LEFT_ARROW",
                 "NOTCHED_RIGHT_ARROW", "BENT_ARROW", "BENT_UP_ARROW", "U_TURN_ARROW",
                 "CURVED_RIGHT_ARROW", "CURVED_LEFT_ARROW", "CURVED_UP_ARROW", "CURVED_DOWN_ARROW")
                if hasattr(MSO_SHAPE, item)}
ROUND_SHAPES = {getattr(MSO_SHAPE, item) for item in
                ("OVAL", "DONUT", "BLOCK_ARC", "ARC", "PIE", "CHORD", "ROUNDED_RECTANGLE")
                if hasattr(MSO_SHAPE, item)}
TRIANGLE_SHAPES = {getattr(MSO_SHAPE, item) for item in
                   ("ISOSCELES_TRIANGLE", "RIGHT_TRIANGLE", "PYRAMID")
                   if hasattr(MSO_SHAPE, item)}


def is_cjk(ch):
    return unicodedata.east_asian_width(ch) in ("W", "F")


def char_em(ch):
    if is_cjk(ch):
        return 1.0
    if ch == " ":
        return 0.30
    if ch in NARROW_LATIN:
        return 0.30
    if ch in WIDE_LATIN:
        return 0.90
    if ch.isdigit():
        return 0.58
    if ch.isupper():
        return 0.30 if ch in "IL" else 0.72
    return 0.56


def tokens(text):
    """按可断行位置切分：中文逐字可断，西文连续串不可断，空格附着在前一段。"""
    out, current = [], ""
    for ch in text:
        if is_cjk(ch):
            if current:
                out.append(current)
                current = ""
            out.append(ch)
        elif ch == " ":
            out.append(current + " ")
            current = ""
        else:
            current += ch
    if current:
        out.append(current)
    return out


def wrap_lines(text, size, avail_pt):
    if avail_pt <= 0:
        return 99, 0.0
    lines, current, longest = 1, 0.0, 0.0
    for token in tokens(text):
        width = sum(char_em(ch) for ch in token) * size
        longest = max(longest, width)
        if current > 0 and current + width > avail_pt + 0.01:
            lines += 1
            current = width
        else:
            current += width
    return lines, longest


def has_latin(text):
    return any(ch.isalnum() and not is_cjk(ch) for ch in text)


def box_of(shape):
    try:
        return (shape.left / EMU_IN, shape.top / EMU_IN,
                shape.width / EMU_IN, shape.height / EMU_IN)
    except Exception:
        return None


def is_connector(shape):
    return shape._element.tag.endswith("}cxnSp")


def auto_shape_type(shape):
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return shape.auto_shape_type
    except Exception:
        pass
    return None


def is_smartart(shape):
    try:
        element = shape._element
        if element.tag.endswith("}graphicFrame"):
            for node in element.iter():
                if node.tag.endswith("}relIds"):
                    return True
                if node.tag.endswith("}graphicData") and "diagram" in (node.get("uri") or ""):
                    return True
    except Exception:
        pass
    return False


def picture_info(shape):
    try:
        image = shape.image
        return image.size, image.blob
    except Exception:
        return None, None


def cluster_axis(centers, span):
    if not centers:
        return []
    ordered = sorted(centers)
    clusters = [[ordered[0]]]
    for value in ordered[1:]:
        if value - clusters[-1][-1] > 0.08 * span:
            clusters.append([value])
        else:
            clusters[-1].append(value)
    return clusters


def walk_shapes(shapes):
    for shape in shapes:
        yield shape
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)


def analyze_layout(slide, slide_w, slide_h):
    """返回 (关系模型, 页面骨架, 视觉手段, 是否纯文字页, 用到的字号档)。"""
    items = []
    for shape in slide.shapes:
        rect = box_of(shape)
        if rect is None or rect[2] <= 0.05 or rect[3] <= 0.05:
            continue
        items.append((shape, rect))
    all_boxes = [(shape, box_of(shape)) for shape in slide.shapes if box_of(shape) is not None]

    pictures = [item for item in items if getattr(item[0], "shape_type", None) == MSO_SHAPE_TYPE.PICTURE]
    charts = [item for item in items if getattr(item[0], "has_chart", False)]
    tables = [item for item in items if getattr(item[0], "has_table", False)]
    autos = [(shape, rect, auto_shape_type(shape)) for shape, rect in items
             if auto_shape_type(shape) is not None]
    texts = [(shape, rect) for shape, rect in items
             if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip()]

    sizes = set()
    big_number = False
    for shape, _ in texts:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size:
                    sizes.add(round(run.font.size.pt, 1))
                    if run.font.size.pt >= 28 and any(ch.isdigit() for ch in run.text):
                        big_number = True

    header_zone = 0.62
    structural_w = 0.75 * slide_w
    content = [(shape, rect) for shape, rect in items
               if rect[1] >= header_zone and rect[2] < structural_w]
    columns = cluster_axis([rect[0] + rect[2] / 2.0 for _, rect in content], slide_w)
    significant = [cluster for cluster in columns if len(cluster) >= 2]
    two_column = False
    if len(significant) == 2 and len(content) >= 6:
        left, right = significant
        mass = max(0.25 * len(content), 2)
        if (len(left) >= mass and len(right) >= mass and right[0] - left[-1] >= 0.20 * slide_w
                and len(left) + len(right) >= 0.5 * len(content)):
            two_column = True

    cards = [(shape, rect) for shape, rect, _ in autos
             if rect[1] >= header_zone and rect[2] < structural_w
             and rect[2] * rect[3] >= 0.6 and 0.5 <= rect[2] / max(rect[3], 0.01) <= 6.0]
    grid = False
    if len(cards) >= 4:
        widths = sorted(rect[2] for _, rect in cards)
        median_w = widths[len(widths) // 2]
        similar = [(shape, rect) for shape, rect in cards if abs(rect[2] - median_w) <= 0.3 * median_w]
        if len(similar) >= 4:
            rows = cluster_axis([rect[1] + rect[3] / 2.0 for _, rect in similar], slide_h)
            cols = cluster_axis([rect[0] + rect[2] / 2.0 for _, rect in similar], slide_w)
            if len(rows) >= 2 and len(cols) >= 2:
                grid = True

    node_like = [(shape, rect) for shape, rect in content
                 if auto_shape_type(shape) is not None
                 or getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE
                 or getattr(shape, "has_table", False) or getattr(shape, "has_chart", False)]
    center_radial = False
    if len(node_like) >= 4:
        middle = [rect for _, rect in node_like
                  if 0.3 * slide_w <= rect[0] + rect[2] / 2.0 <= 0.7 * slide_w
                  and rect[2] * rect[3] >= 1.0]
        outer = [rect for _, rect in node_like
                 if rect[0] + rect[2] / 2.0 < 0.3 * slide_w or rect[0] + rect[2] / 2.0 > 0.7 * slide_w]
        if len(middle) == 1 and len(outer) >= 3:
            center_radial = True

    full_bands = [rect for _, rect, _ in autos if rect[2] >= 0.75 * slide_w]
    picture_area = sum(rect[2] * rect[3] for _, rect in pictures)
    picture_side = None
    if pictures:
        picture_side = "right" if (pictures[0][1][0] + pictures[0][1][2] / 2.0) > slide_w / 2.0 else "left"

    if picture_side and picture_area >= 0.06 * slide_w * slide_h and texts:
        skeleton = "主图+侧栏"
    elif center_radial:
        skeleton = "中心辐射"
    elif two_column:
        skeleton = "左右双栏"
    elif grid:
        skeleton = "网格/矩阵"
    elif len(full_bands) >= 3:
        skeleton = "上下分区"
    else:
        skeleton = "全宽单栏"

    rounds = [(shape, rect) for shape, rect, kind in autos
              if kind in ROUND_SHAPES and rect[2] < structural_w]
    arrows = [(shape, rect) for shape, rect, kind in autos
              if kind in ARROW_SHAPES and rect[2] < structural_w]
    triangles = [(shape, rect) for shape, rect, kind in autos if kind in TRIANGLE_SHAPES]
    lines = [(shape, rect) for shape, rect in all_boxes
             if rect is not None and (is_connector(shape) or rect[2] <= 0.06 or rect[3] <= 0.06)]

    circular = False
    if len(rounds) >= 3:
        spots = [(rect[0] + rect[2] / 2.0, rect[1] + rect[3] / 2.0) for _, rect in rounds]
        center_x = sum(item[0] for item in spots) / len(spots)
        center_y = sum(item[1] for item in spots) / len(spots)
        radii = sorted(math.hypot(item[0] - center_x, item[1] - center_y) for item in spots)
        groups = [[radii[0]]]
        for value in radii[1:]:
            if abs(value - groups[-1][-1]) <= 0.25 * max(groups[-1][-1], 0.01):
                groups[-1].append(value)
            else:
                groups.append([value])
        ring = max(groups, key=len)
        if len(ring) >= 3 and sum(ring) / len(ring) > 0.5:
            circular = True

    timeline = False
    for _, rect in lines:
        if rect[2] >= 0.5 * slide_w and rect[3] <= 0.06 and rect[1] + rect[3] / 2.0 >= 1.2:
            line_center = rect[1] + rect[3] / 2.0
            attached = [other for _, other in content
                        if abs(other[1] + other[3] / 2.0 - line_center) <= 0.55]
            if len(attached) >= 3 and len(cluster_axis([item[0] + item[2] / 2.0 for item in attached],
                                                       slide_w)) >= 2:
                timeline = True
                break

    hierarchy = False
    if triangles:
        hierarchy = True
    else:
        bars = sorted([rect for _, rect, _ in autos if rect[2] >= 0.35], key=lambda rect: rect[1])
        for index in range(len(bars) - 2):
            first, second, third = bars[index], bars[index + 1], bars[index + 2]
            if first[2] < second[2] < third[2] and first[1] < second[1] < third[1]:
                hierarchy = True
                break

    if tables:
        relation = "矩阵对照"
    elif charts:
        relation = "数据对比"
    elif circular:
        relation = "循环"
    elif len(arrows) >= 2:
        relation = "流程"
    elif timeline:
        relation = "时间线"
    elif hierarchy:
        relation = "层级"
    elif center_radial:
        relation = "辐射"
    elif grid:
        relation = "并列"
    elif two_column:
        relation = "对比"
    else:
        relation = "叙述"

    means = []
    if pictures:
        means.append("图片")
    if charts:
        means.append("图表")
    if tables:
        means.append("表格")
    if autos:
        means.append("图形结构")
    if big_number:
        means.append("大数字")
    if len(sizes) >= 3:
        means.append("多级字号")

    pure_text = (not pictures and not charts and not tables
                 and not [item for item in autos if item[1][2] >= 0.3 and item[1][3] >= 0.3]
                 and bool(texts))
    return relation, skeleton, means, pure_text, sizes


def audit(path, min_size=16.0, notes_min=150, tol=1.06, mode="balanced", allow_smartart=False):
    findings = []
    layout_rows = []
    prs = Presentation(path)
    slide_w, slide_h = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN

    def add(slide_no, severity, rule, kind, detail, extra=""):
        findings.append((severity, slide_no, rule, kind, detail, extra))

    def soft(severity):
        return severity if mode == "strict" else "WARNING"

    previous_skeleton = None
    repeated_run = 0
    pure_text_slides = []
    block_slides = []
    relations = {}

    for slide_no, slide in enumerate(prs.slides, 1):
        text_boxes = []
        visual_boxes = []
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        exempt = any(tag in notes for tag in EXCEPTION_TAGS)

        def check_runs(runs, where):
            for run in runs:
                if not run.text.strip():
                    continue
                for ch in run.text:
                    if ch in " \n\t":
                        continue
                    if "０" <= ch <= "９" or "Ａ" <= ch <= "Ｚ" or "ａ" <= ch <= "ｚ":
                        add(slide_no, soft("ERROR"), "TXT.FULLWIDTH", "全角西文", repr(ch),
                            "字母与数字改半角；中文标点不受限")
                    try:
                        ch.encode("gbk")
                    except UnicodeEncodeError:
                        add(slide_no, soft("ERROR"), "TXT.GLYPH", "字形缺失", repr(ch),
                            f"U+{ord(ch):04X} {unicodedata.name(ch, '?')}；上下标请写成 ASCII")

        for shape in slide.shapes:
            rect = box_of(shape)
            if rect is not None:
                left, top, width, height = rect
                if (getattr(shape, "rotation", 0) or 0) == 0:
                    if (left < -0.01 or top < -0.01 or left + width > slide_w + 0.01
                            or top + height > slide_h + 0.01):
                        add(slide_no, "ERROR", "GEO.OUTSIDE", "越界",
                            f"{shape.shape_type} @({left:.2f},{top:.2f}) {width:.2f}x{height:.2f}",
                            f"页面 {slide_w:.2f}x{slide_h:.2f}")
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                add(slide_no, "WARNING", "GEO.GROUP", "组合形状",
                    "组合形状内部元素未递归做几何审查", "请在 PowerPoint 中复核")
            if is_smartart(shape) and not allow_smartart:
                add(slide_no, soft("ERROR"), "MED.SMARTART", "禁用对象", "页面存在 SmartArt",
                    "改用文本、表格、形状或图表；沿用原模板时用 --allow-smartart")
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                pixels, blob = picture_info(shape)
                if pixels is None or not blob:
                    add(slide_no, "ERROR", "IMG.MISSING", "图片缺失",
                        "图片对象无法读取原始数据", "可能是链接图或文件损坏")
                else:
                    pixel_w, pixel_h = pixels
                    if rect is not None and rect[2] > 0.2:
                        dpi = pixel_w / rect[2]
                        if dpi < IMAGE_DPI_MIN:
                            add(slide_no, "ERROR", "IMG.RES", "图片分辨率",
                                f"实际约 {dpi:.0f} dpi（{pixel_w}x{pixel_h}px）",
                                f"按放置尺寸至少需要 {IMAGE_DPI_MIN:.0f} dpi")
                        elif dpi < IMAGE_DPI_WARN:
                            add(slide_no, "WARNING", "IMG.RES", "图片分辨率",
                                f"实际约 {dpi:.0f} dpi（{pixel_w}x{pixel_h}px）",
                                f"建议不低于 {IMAGE_DPI_WARN:.0f} dpi")
                    if max(pixel_w, pixel_h) < IMAGE_EDGE_WARN:
                        add(slide_no, "WARNING", "IMG.QUALITY", "图片画质",
                            f"长边仅 {max(pixel_w, pixel_h)}px", "放大后可能出现糊边或水印残留")
            if (getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE
                    or getattr(shape, "has_chart", False) or getattr(shape, "has_table", False)):
                if rect is not None:
                    visual_boxes.append((str(shape.shape_type),) + rect)

            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                frame = shape.text_frame
                width_pt = shape.width / EMU_PT - (frame.margin_left + frame.margin_right) / EMU_PT
                height_pt = shape.height / EMU_PT - (frame.margin_top + frame.margin_bottom) / EMU_PT
                needed = 0.0
                for paragraph in frame.paragraphs:
                    runs = [run.text for run in paragraph.runs]
                    if not runs:
                        continue
                    size = max((run.font.size.pt if run.font.size else min_size) for run in paragraph.runs)
                    spacing = paragraph.line_spacing if isinstance(paragraph.line_spacing, (int, float)) else 1.0
                    count, longest = wrap_lines("".join(runs), size, width_pt)
                    needed += count * size * max(spacing, 1.0) * 1.18
                    if longest > width_pt * tol:
                        add(slide_no, "ERROR", "TXT.NOBREAK", "西文不可断行",
                            "".join(runs)[:30], f"最长段 {longest:.0f}pt / 栏宽 {width_pt:.0f}pt")
                check_runs([run for paragraph in frame.paragraphs for run in paragraph.runs], "")
                if height_pt > 0 and needed > height_pt * tol:
                    add(slide_no, "ERROR", "TXT.OVERFLOW", "文本溢出",
                        frame.text[:30].replace("\n", " "),
                        f"约需 {needed:.0f}pt > 可用 {height_pt:.0f}pt")
                if rect is not None:
                    text_boxes.append((frame.text[:22].replace("\n", " "),) + rect)

            if getattr(shape, "has_table", False):
                table = shape.table
                table_width = sum(column.width for column in table.columns) / EMU_IN
                if rect is not None and abs(table_width - rect[2]) > 0.05:
                    add(slide_no, "WARNING", "GEO.TABLE_WIDTH", "表宽不符",
                        f"列宽合计 {table_width:.2f}in vs 表格框 {rect[2]:.2f}in",
                        "列宽被挤压或留白，请在 PowerPoint 中复核")
                column_widths = [column.width / EMU_PT for column in table.columns]
                row_heights = [row.height / EMU_PT for row in table.rows]
                for row_no, row in enumerate(table.rows):
                    for col_no, cell in enumerate(row.cells):
                        frame = cell.text_frame
                        if not frame.text.strip() or getattr(cell, "is_spanned", False):
                            continue
                        span_w = getattr(cell, "span_width", 1) or 1
                        span_h = getattr(cell, "span_height", 1) or 1
                        avail_w = (sum(column_widths[col_no:col_no + span_w])
                                   - (cell.margin_left + cell.margin_right) / EMU_PT)
                        avail_h = (sum(row_heights[row_no:row_no + span_h])
                                   - (cell.margin_top + cell.margin_bottom) / EMU_PT)
                        needed = 0.0
                        for paragraph in frame.paragraphs:
                            if not paragraph.runs:
                                continue
                            size = max((run.font.size.pt if run.font.size else min_size)
                                       for run in paragraph.runs)
                            spacing = paragraph.line_spacing if isinstance(paragraph.line_spacing, (int, float)) else 1.0
                            count, longest = wrap_lines("".join(run.text for run in paragraph.runs), size, avail_w)
                            needed += count * size * max(spacing, 1.0) * 1.18
                            if longest > avail_w * tol:
                                add(slide_no, "ERROR", "TXT.NOBREAK", "表格西文不可断行",
                                    frame.text[:24], f"列{col_no} 最长段 {longest:.0f}pt / {avail_w:.0f}pt")
                        if avail_h > 0 and needed > avail_h * tol:
                            add(slide_no, "ERROR", "TXT.TABLE_OVERFLOW", "表格文本溢出",
                                frame.text[:24],
                                f"约需 {needed:.0f}pt > 行高 {avail_h:.0f}pt（会撑高压到下方图形）")
                        check_runs([run for paragraph in frame.paragraphs for run in paragraph.runs], "表格")

        for index, first in enumerate(text_boxes):
            for second in text_boxes[index + 1:]:
                overlap = rectangle_overlap(first, second)
                if overlap > 0.18:
                    add(slide_no, "ERROR", "TXT.OVERLAP", "文本重叠",
                        f"[{first[0]}] × [{second[0]}]", f"覆盖 {overlap:.0%}")
            for visual in visual_boxes:
                overlap = rectangle_overlap(first, visual)
                if overlap > 0.18:
                    add(slide_no, "WARNING", "GEO.OVERLAP_VISUAL", "文字与视觉元素重叠",
                        f"[{first[0]}] × [{visual[0]}]", f"覆盖 {overlap:.0%}")

        if len(notes) < notes_min:
            add(slide_no, soft("ERROR"), "NOTE.SHORT", "备注缺失或过短", "逐字稿",
                f"{len(notes)} 字 < {notes_min} 字")

        relation, skeleton, means, pure_text, sizes = analyze_layout(slide, slide_w, slide_h)
        relations[relation] = relations.get(relation, 0) + 1
        layout_rows.append((slide_no, relation, skeleton, "、".join(means) or "无", pure_text))

        if skeleton == previous_skeleton and not exempt:
            repeated_run += 1
            if repeated_run == 3:
                add(slide_no, "WARNING", "LAY.REPEAT", "骨架重复",
                    f"连续 3 页使用同一主骨架「{skeleton}」",
                    "换骨架或换关系模型，或在备注写 [例外:骨架重复] 并说明理由")
        else:
            repeated_run = 1
        previous_skeleton = skeleton

        if pure_text:
            pure_text_slides.append(slide_no)
            if len(sizes) < 2:
                add(slide_no, "WARNING", "LAY.PURETEXT_FLAT", "纯文字页",
                    "纯文字页只有一种字号", "至少建立两级字号层级，或用留白与对齐承担视觉任务")
        if skeleton == "网格/矩阵" or relation in ("矩阵对照", "并列"):
            block_slides.append(slide_no)
        if not means:
            add(slide_no, "WARNING", "LAY.MEANS", "视觉手段缺失",
                "本页没有图片、图表、表格或图形结构", "至少安排一种视觉手段；纯文字页需靠排版承担")

    total = max(1, len(prs.slides._sldIdLst))
    if pure_text_slides and len(pure_text_slides) / total > PURE_TEXT_LIMIT:
        add(0, "WARNING", "LAY.PURETEXT", "纯文字页比例",
            f"{len(pure_text_slides)}/{total} 页为纯文字页（{len(pure_text_slides) / total:.0%}）",
            f"建议不超过 {PURE_TEXT_LIMIT:.0%}；页 {pure_text_slides}")
    if block_slides and len(block_slides) / total > BLOCK_LIMIT and total >= 6:
        add(0, "WARNING", "LAY.BLOCKRATIO", "方块矩阵占比",
            f"方块/矩阵类页面 {len(block_slides)}/{total}（{len(block_slides) / total:.0%}）",
            f"建议不超过 {BLOCK_LIMIT:.0%}；页 {block_slides}")
    if total >= 6 and len(relations) < MIN_RELATION_MODELS:
        add(0, "SUGGESTION", "LAY.MODEL", "关系模型覆盖",
            f"全篇只用了 {len(relations)} 种关系模型：{'、'.join(sorted(relations))}",
            f"建议至少 {MIN_RELATION_MODELS} 种，按内容关系换骨架")

    return findings, layout_rows


def rectangle_overlap(first, second):
    _, l1, t1, w1, h1 = first
    _, l2, t2, w2, h2 = second
    overlap_w = min(l1 + w1, l2 + w2) - max(l1, l2)
    overlap_h = min(t1 + h1, t2 + h2) - max(t1, t2)
    if overlap_w <= 0.02 or overlap_h <= 0.02:
        return 0.0
    smaller = min(w1 * h1, w2 * h2)
    return overlap_w * overlap_h / smaller if smaller > 0 else 0.0


def format_findings(findings):
    if not findings:
        return "OK"
    lines = []
    for severity, slide_no, rule, kind, detail, extra in sorted(
            findings, key=lambda item: (item[1], SEVERITY_ORDER.get(item[0], 3), item[2])):
        where = f"第{slide_no}页" if slide_no else "全篇"
        lines.append(f"{severity}  {where}  [{rule}] {kind}: {detail}" + (f" — {extra}" if extra else ""))
    return "\n".join(lines)


def format_layout(rows):
    lines = [f"{'页':<4}{'关系模型':<10}{'页面骨架':<12}{'视觉手段':<24}纯文字", "-" * 66]
    for slide_no, relation, skeleton, means, pure_text in rows:
        lines.append(f"{slide_no:<6}{relation:<12}{skeleton:<14}{means:<26}{'是' if pure_text else '否'}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="*")
    parser.add_argument("--min-size", type=float, default=16)
    parser.add_argument("--notes-min", type=int, default=150)
    parser.add_argument("--overflow-tol", type=float, default=1.06)
    parser.add_argument("--mode", choices=("balanced", "strict"), default="balanced")
    parser.add_argument("--allow-smartart", action="store_true", help="沿用原模板时保留已有 SmartArt")
    parser.add_argument("--layout-report", action="store_true", help="打印每页的关系模型与骨架分类")
    args = parser.parse_args()
    if not args.files:
        print(__doc__)
        return 2

    exit_code = 0
    for path in args.files:
        if not os.path.exists(path):
            print("文件不存在:", path)
            exit_code = 1
            continue
        findings, rows = audit(path, args.min_size, args.notes_min, args.overflow_tol,
                               args.mode, args.allow_smartart)
        errors = [item for item in findings if item[0] == "ERROR"]
        print(("FAIL" if errors else "OK  "), os.path.basename(path), f"({len(findings)} 项)")
        if findings:
            print(format_findings(findings))
        if args.layout_report:
            print(format_layout(rows))
        if errors:
            exit_code = 1
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())