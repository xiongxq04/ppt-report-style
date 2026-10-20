# -*- coding: utf-8 -*-
"""校验 pptx 是否符合汇报规范：字号下限、文字配色、字体白名单。

用法:
  python verify_ppt_style.py <file.pptx> [more.pptx ...] [选项]
选项:
  --accent HEX      主题色，默认 283E4E
  --min-size N      最小字号(pt)，默认 16
  --fonts "a,b,c"   允许的字体名，逗号分隔；默认 宋体,楷体,黑体
  --allow-colors    额外允许的文字色（逗号分隔十六进制），用于自定义配色
退出码 0 表示全部通过。
"""
import sys, os, argparse
sys.stdout.reconfigure(encoding="utf-8")
try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

BASE = {"000000", "FFFFFF"}


def check(path, accent, min_size, fonts, allow):
    allowed_colors = BASE | {accent.upper()} | {c.upper() for c in allow}
    font_whitelist = {f.strip() for f in fonts.split(",") if f.strip()}
    bad = []
    prs = Presentation(path)
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False):
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        sz = r.font.size.pt if r.font.size else None
                        if sz is None or sz < min_size:
                            bad.append((si, "字号", r.text[:20], sz))
                        try:
                            col = str(r.font.color.rgb)
                        except Exception:
                            col = "?"
                        if col not in allowed_colors:
                            bad.append((si, "颜色", r.text[:20], col))
                        name = (r.font.name or "").strip()
                        if font_whitelist and name and name not in font_whitelist:
                            bad.append((si, "字体", r.text[:16], name))
            if getattr(sh, "has_chart", False):
                for rPr in sh.chart._chartSpace.iter(qn("a:defRPr")):
                    sz = rPr.get("sz")
                    if sz and int(sz) < min_size * 100:
                        bad.append((si, "图表字号", "", int(sz) / 100))
                    sf = rPr.find(qn("a:solidFill"))
                    if sf is not None:
                        sc = sf.find(qn("a:srgbClr"))
                        if sc is not None and sc.get("val", "").upper() not in allowed_colors:
                            bad.append((si, "图表颜色", "", sc.get("val")))
    return bad


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--accent", default="283E4E")
    ap.add_argument("--min-size", type=float, default=16)
    ap.add_argument("--fonts", default="宋体,楷体,黑体")
    ap.add_argument("--allow-colors", default="")
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    allow = [c.strip() for c in a.allow_colors.split(",") if c.strip()]
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        bad = check(path, a.accent, a.min_size, a.fonts, allow)
        if bad:
            rc = 1
            print("FAIL", os.path.basename(path))
            for b in bad:
                print("   ", b)
        else:
            print("OK  ", os.path.basename(path))
    return rc


sys.exit(main())

