# -*- coding: utf-8 -*-
"""汇报 PPT 自动审查：越界 / 文本溢出 / 元素重叠 / 字形与字体槽 / 备注 / 表格与图表。

用法: python audit_ppt.py <file.pptx> [更多文件 ...] [选项]
选项:
  --min-size N        最小字号(pt)，默认 16
  --accent HEX        主题色，默认 283E4E
  --fonts "a,b,c"     允许的中文字体（a:ea），默认 宋体,楷体,黑体
  --latin-fonts "x,y" 允许的西文字体（a:latin），默认 Arial,Times New Roman
  --notes-min N       每页备注最少字数，默认 150
  --overflow-tol      溢出容差倍数，默认 1.06
退出码 0 = 全部通过。

中英混排的三条硬道理（本脚本据此判定）：
1. 宽度不能按"西文一律 0.52 倍"估：Arial 里 i/l 约 0.28 em，M/W 约 0.9 em，全大写缩写
   （BPM、LUFS、RMS）比小写宽得多。这里按字符分类取 em 宽，宁可估宽也不漏报。
2. 西文不能像中文那样逐字断行：换行只发生在空格与中文之间。一个长 token
   （如 conversation_log.csv）比栏宽还长时，必然溢出或被挤到下一栏——单独报"西文不可断行"。
3. 字体分两槽：a:latin 管西文、a:ea 管中文。只设 font.name=宋体 会让英文用中文字形渲染，
   所以两槽都要查；表格单元格还要按栏宽单独算换行，否则"行高 0.3in"会被撑成两行而压到下方图形。
"""
import sys, os, math, argparse, unicodedata
sys.stdout.reconfigure(encoding="utf-8")
try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

EMU_IN = 914400.0
EMU_PT = 12700.0
BASE = {"000000", "FFFFFF"}
CJK_FACES = {"宋体", "楷体", "黑体", "SimSun", "KaiTi", "SimHei", "Microsoft YaHei"}
NARROW_LATIN = set("ijlrtf|.,:;'\"()[]{}/!-\\` ")
WIDE_LATIN = set("mwMW@%&")


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
    if ch.islower():
        return 0.56
    return 0.56


def tokens(s):
    """按可断行位置切分：中文逐字可断，西文连续串不可断，空格附着在前一段。"""
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


def audit(path, min_size, accent, fonts, latin_fonts, notes_min, tol):
    issues = []
    prs = Presentation(path)
    sw, sh_ = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
    allowed = BASE | {accent.upper()}
    ea_ok = {f.strip() for f in fonts.split(",") if f.strip()}
    lat_ok = {f.strip() for f in latin_fonts.split(",") if f.strip()}

    for si, slide in enumerate(prs.slides, 1):
        boxes = []

        def add(kind, detail, extra=""):
            issues.append((si, kind, detail, extra))

        def check_runs(runs, where):
            for r in runs:
                if not r.text.strip():
                    continue
                sz = r.font.size.pt if r.font.size else None
                if sz is None or sz < min_size:
                    add(where + "字号", r.text[:20], sz)
                try:
                    col = str(r.font.color.rgb)
                except Exception:
                    col = "?"
                if col not in allowed:
                    add(where + "颜色", r.text[:20], col)
                lat, ea = slots(r)
                if any(is_cjk(c) for c in r.text):
                    if not ea or (ea_ok and ea not in ea_ok):
                        add(where + "中文字体", r.text[:16], ea or "a:ea 未设置")
                if has_latin(r.text):
                    if lat in CJK_FACES:
                        add(where + "西文用中文字体", r.text[:16], f"latin={lat}")
                    elif not lat or (lat_ok and lat not in lat_ok):
                        add(where + "西文字体", r.text[:16], lat or "a:latin 未设置")
                for ch in r.text:
                    if ch in " \n\t":
                        continue
                    if "０" <= ch <= "９" or "Ａ" <= ch <= "Ｚ" or "ａ" <= ch <= "ｚ":
                        add(where + "全角西文", repr(ch), "字母与数字改半角；中文标点（，：（））不受限")
                    try:
                        ch.encode("gbk")
                    except UnicodeEncodeError:
                        add(where + "字形缺失", repr(ch), f"U+{ord(ch):04X} {unicodedata.name(ch,'?')}")

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
                bw = shp.width / EMU_PT - (tf.margin_left + tf.margin_right) / EMU_PT
                bh = shp.height / EMU_PT - (tf.margin_top + tf.margin_bottom) / EMU_PT
                need = 0.0
                for p in tf.paragraphs:
                    runs = [r.text for r in p.runs]
                    if not runs:
                        continue
                    size = max((r.font.size.pt if r.font.size else min_size) for r in p.runs)
                    ls = p.line_spacing if isinstance(p.line_spacing, (int, float)) else 1.0
                    n, longest = wrap_lines("".join(runs), size, bw)
                    need += n * size * max(ls, 1.0) * 1.18
                    if longest > bw * tol:
                        add("西文不可断行", "".join(runs)[:30], f"最长段 {longest:.0f}pt / 栏宽 {bw:.0f}pt")
                check_runs([r for p in tf.paragraphs for r in p.runs], "")
                if bh > 0 and need > bh * tol:
                    add("文本溢出", tf.text[:30].replace("\n", " "), f"需 {need:.0f}pt / 框 {bh:.0f}pt")
                boxes.append((tf.text[:22].replace("\n", " "), l, t, w, h))

            if getattr(shp, "has_table", False):
                tb = shp.table
                tw = sum(c.width for c in tb.columns) / EMU_IN
                if abs(tw - shp.width / EMU_IN) > 0.05:
                    add("表宽不符", f"{tw:.2f} vs {shp.width/EMU_IN:.2f}", "")
                colw = [c.width / EMU_PT for c in tb.columns]
                rowhs = [r.height / EMU_PT for r in tb.rows]
                for ri, row in enumerate(tb.rows):
                    rowh = rowhs[ri]
                    for ci, cell in enumerate(row.cells):
                        tf = cell.text_frame
                        if not tf.text.strip():
                            continue
                        if getattr(cell, "is_spanned", False):
                            continue        # 合并单元格的从属格：文本与原点格重复，跳过免误报
                        span_w = getattr(cell, "span_width", 1) or 1
                        span_h = getattr(cell, "span_height", 1) or 1
                        avail_w = sum(colw[ci:ci + span_w]) - (cell.margin_left + cell.margin_right) / EMU_PT
                        avail_h = sum(rowhs[ri:ri + span_h]) - (cell.margin_top + cell.margin_bottom) / EMU_PT
                        need = 0.0
                        for p in tf.paragraphs:
                            if not p.runs:
                                continue
                            size = max((r.font.size.pt if r.font.size else min_size) for r in p.runs)
                            ls = p.line_spacing if isinstance(p.line_spacing, (int, float)) else 1.0
                            n, longest = wrap_lines("".join(r.text for r in p.runs), size, avail_w)
                            need += n * size * max(ls, 1.0) * 1.18
                            if longest > avail_w * tol:
                                add("表格西文不可断行", tf.text[:24], f"列{ci} 最长段 {longest:.0f}pt / {avail_w:.0f}pt")
                        if avail_h > 0 and need > avail_h * tol:
                            add("表格文本溢出", tf.text[:24], f"需 {need:.0f}pt / 行高 {avail_h:.0f}pt（会撑高压到下方图形）")
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
    ap.add_argument("--fonts", default="宋体,楷体,黑体", help="中文字体（a:ea）白名单")
    ap.add_argument("--latin-fonts", default="Arial,Times New Roman", help="西文字体（a:latin）白名单")
    ap.add_argument("--notes-min", type=int, default=150)
    ap.add_argument("--overflow-tol", type=float, default=1.06)
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        issues = audit(path, a.min_size, a.accent, a.fonts, a.latin_fonts, a.notes_min, a.overflow_tol)
        if issues:
            rc = 1
            print("FAIL", os.path.basename(path), f"({len(issues)} 项)")
            for it in issues[:40]:
                print("   ", it)
            if len(issues) > 40:
                print("    ...另", len(issues) - 40, "项")
        else:
            print("OK  ", os.path.basename(path))
    return rc


sys.exit(main())
