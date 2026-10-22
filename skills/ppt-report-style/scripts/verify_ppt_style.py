# -*- coding: utf-8 -*-
"""校验 pptx 是否符合汇报规范：字号下限、文字配色、中文字体（ea 槽）与西文字体（latin 槽）双白名单。

用法:
  python verify_ppt_style.py <file.pptx> [more.pptx ...] [选项]
选项:
  --accent HEX        主题色，默认 283E4E
  --min-size N        最小字号(pt)，默认 16
  --fonts "a,b,c"     允许的中文字体（写入 a:ea），默认 宋体,楷体,黑体
  --latin-fonts "x,y" 允许的西文字体（写入 a:latin），默认 Arial,Times New Roman
  --allow-colors HEX  额外允许的文字色（逗号分隔）
  --allow-cjk-latin   容忍西文槽填中文字体（只用于跑旧稿，新稿不要加）
  --sizes "a,b,..."   允许的字阶档位，默认 16,18,20,30,36,44
  --max-steps N       每页最多几档字号，默认 5
退出码 0 表示全部通过。

为什么必须分两槽：python-pptx 的 run.font.name 写的是 a:latin（西文槽），中文字符实际取 a:ea。
只写 font.name=宋体 → 英文数字用宋体的西文字形，字距松、字形丑，且本项检查会漏掉；
只写 a:ea → 英文回落到主题字体，跨机器不稳定。
正确做法：两槽都显式写，并按 宋体/楷体 → Times New Roman、黑体 → Arial 配对。
"""
import sys, os, argparse
sys.stdout.reconfigure(encoding="utf-8")
try:
    from pptx import Presentation
    from pptx.oxml.ns import qn
except ImportError:
    print("需要 python-pptx：pip install python-pptx"); sys.exit(2)

BASE = {"000000", "FFFFFF"}
CJK_FACES = {"宋体", "楷体", "黑体", "SimSun", "KaiTi", "SimHei", "MS Mincho", "Microsoft YaHei"}


def has_cjk(s):
    return any("\u3000" <= c <= "\u9fff" or "\uff00" <= c <= "\uffef" for c in s)


def has_latin(s):
    return any(c.isalpha() or c.isdigit() for c in s) and not all(
        "\u3000" <= c <= "\u9fff" or "\uff00" <= c <= "\uffef" or not c.isalnum() for c in s)


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


def check(path, accent, min_size, fonts, latin_fonts, allow, allow_cjk_latin, sizes, max_steps):
    allowed_colors = BASE | {accent.upper()} | {c.upper().strip() for c in allow if c.strip()}
    ea_ok = {f.strip() for f in fonts.split(",") if f.strip()}
    lat_ok = {f.strip() for f in latin_fonts.split(",") if f.strip()}
    size_ok = {float(s) for s in sizes.split(",") if s.strip()}
    bad = []
    prs = Presentation(path)
    for si, slide in enumerate(prs.slides, 1):
        steps = set()
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False):
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        txt = r.text
                        if not txt.strip():
                            continue
                        sz = r.font.size.pt if r.font.size else None
                        if sz is None or sz < min_size:
                            bad.append((si, "字号", txt[:20], sz))
                        if sz is not None:
                            steps.add(sz)
                            if size_ok and sz not in size_ok:
                                bad.append((si, "字号不在字阶内", txt[:20], f"{sz}pt，字阶={sorted(size_ok)}"))
                        try:
                            col = str(r.font.color.rgb)
                        except Exception:
                            col = "?"
                        if col not in allowed_colors:
                            bad.append((si, "颜色", txt[:20], col))
                        lat, ea = slots(r)
                        if has_cjk(txt):
                            if not ea:
                                bad.append((si, "中文字体未设置", txt[:16], "a:ea 空，继承主题字体"))
                            elif ea_ok and ea not in ea_ok:
                                bad.append((si, "中文字体", txt[:16], ea))
                        if has_latin(txt):
                            if not lat:
                                bad.append((si, "西文字体未设置", txt[:16], "a:latin 空"))
                            elif lat in CJK_FACES:
                                if not allow_cjk_latin:
                                    bad.append((si, "西文用中文字体", txt[:16], f"latin={lat}（英文将用中文字形渲染）"))
                            elif lat_ok and lat not in lat_ok:
                                bad.append((si, "西文字体", txt[:16], lat))
                        if ea == "黑体" and r.font.bold:
                            bad.append((si, "黑体加粗", txt[:16], "合成粗会糊边；强调改用字号或主题色"))
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
            if getattr(sh, "has_table", False):
                for row in sh.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if not r.text.strip():
                                    continue
                                sz = r.font.size.pt if r.font.size else None
                                if sz is None or sz < min_size:
                                    bad.append((si, "表格字号", r.text[:16], sz))
                                if sz is not None:
                                    steps.add(sz)
                                    if size_ok and sz not in size_ok:
                                        bad.append((si, "表格字号不在字阶内", r.text[:16], sz))
                                lat, ea = slots(r)
                                if ea == "楷体":
                                    bad.append((si, "楷体入表格", r.text[:16], "表格只用宋体/黑体"))
                                if has_cjk(r.text) and ea_ok and (not ea or ea not in ea_ok):
                                    bad.append((si, "表格中文字体", r.text[:16], ea or "空"))
                                if has_latin(r.text) and lat in CJK_FACES and not allow_cjk_latin:
                                    bad.append((si, "表格西文用中文字体", r.text[:16], lat))
        if len(steps) > max_steps:
            bad.append((si, "字号层级过多", "", f"{len(steps)} 档 > {max_steps}：{sorted(steps)}"))
    return bad


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--accent", default="283E4E")
    ap.add_argument("--min-size", type=float, default=16)
    ap.add_argument("--fonts", default="宋体,楷体,黑体", help="中文字体（a:ea）白名单")
    ap.add_argument("--latin-fonts", default="Arial,Times New Roman", help="西文字体（a:latin）白名单")
    ap.add_argument("--allow-colors", default="")
    ap.add_argument("--allow-cjk-latin", action="store_true", help="容忍西文槽填中文字体（旧稿用）")
    ap.add_argument("--sizes", default="16,18,20,30,36,44", help="允许的字阶档位，逗号分隔")
    ap.add_argument("--max-steps", type=int, default=5, help="每页最多出现几档字号")
    a = ap.parse_args()
    if not a.files:
        print(__doc__); return 2
    allow = [c.strip() for c in a.allow_colors.split(",") if c.strip()]
    rc = 0
    for path in a.files:
        if not os.path.exists(path):
            print("文件不存在:", path); rc = 1; continue
        bad = check(path, a.accent, a.min_size, a.fonts, a.latin_fonts, allow, a.allow_cjk_latin,
                    a.sizes, a.max_steps)
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


sys.exit(main())
