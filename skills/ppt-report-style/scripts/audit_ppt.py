# -*- coding: utf-8 -*-
"""汇报 PPT 自动审查：越界 / 文本溢出 / 元素重叠 / 字形缺失 / 备注缺失 / 表格与图表。

用法: python audit_ppt.py <file.pptx> [更多文件 ...] [选项]
选项:
  --min-size N      最小字号(pt)，默认 16
  --accent HEX      主题色，默认 283E4E
  --fonts "a,b,c"   允许字体，逗号分隔；默认 宋体,楷体,黑体
  --notes-min N     每页备注最少字数，默认 150
  --overflow-tol    溢出容差倍数，默认 1.06
退出码 0 = 全部通过。

要点：中文用 GBK 覆盖即可保证宋体/黑体/楷体不出现豆腐块；因此把"字符能否用 gbk 编码"
作为字形检查，比逐个查字体 cmap 更稳。数学上下标建议直接写成 V^T、σ1 这类 ASCII 形式。
"""
import sys, os, math, argparse, unicodedata
sys.stdout.reconfigure(encoding="utf-8")
try:
    from pptx import Presentation
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

EMU_IN = 914400.0
BASE = {"000000", "FFFFFF"}


def cjk(ch):
    return unicodedata.east_asian_width(ch) in ("W", "F")


def text_w(s, size):
    return sum(size * (1.0 if cjk(c) else 0.52) for c in s)


def audit(path, min_size, accent, fonts, notes_min, tol):
    issues = []
    prs = Presentation(path)
    sw, sh_ = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
    allowed = BASE | {accent.upper()}
    font_ok = {f.strip() for f in fonts.split(",") if f.strip()}

    for si, slide in enumerate(prs.slides, 1):
        boxes = []

        def add(kind, detail, extra=""):
            issues.append((si, kind, detail, extra))

        for shp in slide.shapes:
            try:
                l, t = shp.left / EMU_IN, shp.top / EMU_IN
                w, h = shp.width / EMU_IN, shp.height / EMU_IN
                if (getattr(shp, "rotation", 0) or 0) == 0:
                    if l < -0.01 or t < -0.01 or l + w > sw + 0.01 or t + h > sh_ + 0.01:
                        add("越界", f"{shp.shape_type} @({l:.2f},{t:.2f}) {w:.2f}x{h:.2f}",
                            f"页面 {sw:.2f}x{sh_:.2f}")
            except Exception:
                pass

            if getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip():
                tf = shp.text_frame
                bw = (shp.width / EMU_IN) * 72 - (tf.margin_left + tf.margin_right) / EMU_IN * 72
                bh = (shp.height / EMU_IN) * 72 - (tf.margin_top + tf.margin_bottom) / EMU_IN * 72
                need = 0.0
                for p in tf.paragraphs:
                    runs = [r.text for r in p.runs]
                    if not runs:
                        continue
                    size = max((r.font.size.pt if r.font.size else min_size) for r in p.runs)
                    ls = p.line_spacing if isinstance(p.line_spacing, (int, float)) else 1.0
                    lines = max(1, math.ceil(sum(text_w(x, size) for x in runs) / bw)) if bw > 0 else 99
                    need += lines * size * max(ls, 1.0) * 1.18
                    for r in p.runs:
                        sz = r.font.size.pt if r.font.size else None
                        if sz is None or sz < min_size:
                            add("字号", r.text[:20], sz)
                        try:
                            col = str(r.font.color.rgb)
                        except Exception:
                            col = "?"
                        if col not in allowed:
                            add("颜色", r.text[:20], col)
                        if font_ok and (r.font.name or "") not in font_ok:
                            add("字体", r.text[:16], r.font.name)
                        for ch in r.text:
                            if ch in " \n\t":
                                continue
                            try:
                                ch.encode("gbk")
                            except UnicodeEncodeError:
                                add("字形缺失", repr(ch), f"U+{ord(ch):04X} {unicodedata.name(ch,'?')}")
                if bh > 0 and need > bh * tol:
                    add("文本溢出", tf.text[:30].replace("\n", " "), f"需 {need:.0f}pt / 框 {bh:.0f}pt")
                boxes.append((tf.text[:22].replace("\n", " "), l, t, w, h))

            if getattr(shp, "has_table", False):
                tb = shp.table
                tw = sum(c.width for c in tb.columns) / EMU_IN
                if abs(tw - shp.width / EMU_IN) > 0.05:
                    add("表宽不符", f"{tw:.2f} vs {shp.width/EMU_IN:.2f}", "")
                for row in tb.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                sz = r.font.size.pt if r.font.size else None
                                if sz is None or sz < min_size:
                                    add("表格字号", r.text[:16], sz)
                                try:
                                    col = str(r.font.color.rgb)
                                except Exception:
                                    col = "?"
                                if col not in allowed:
                                    add("表格颜色", r.text[:16], col)
                                for ch in r.text:
                                    if ch in " \n\t":
                                        continue
                                    try:
                                        ch.encode("gbk")
                                    except UnicodeEncodeError:
                                        add("表格字形缺失", repr(ch), f"U+{ord(ch):04X}")

        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ox = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
                oy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
                if ox > 0.02 and oy > 0.02:
                    area = ox * oy
                    small = min(a[3] * a[4], b[3] * b[4])
                    if small > 0 and area / small > 0.18:
                        add("文本重叠", f"[{a[0]}] x [{b[0]}]", f"覆盖 {area/small*100:.0f}%")

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        if len(notes) < notes_min:
            add("备注缺失/过短", "逐字稿", f"{len(notes)} 字 < {notes_min}")
    return issues


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--min-size", type=float, default=16)
    ap.add_argument("--accent", default="283E4E")
    ap.add_argument("--fonts", default="宋体,楷体,黑体")
    ap.add_argument("--notes-min", type=int, default=150)
    ap.add_argument("--overflow-tol", type=float, default=1.06)
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        issues = audit(path, a.min_size, a.accent, a.fonts, a.notes_min, a.overflow_tol)
        if issues:
            rc = 1
            print("FAIL", os.path.basename(path), f"({len(issues)} 项)")
            for it in issues:
                print("   ", it)
        else:
            print("OK  ", os.path.basename(path))
    return rc


sys.exit(main())
