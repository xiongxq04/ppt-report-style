# -*- coding: utf-8 -*-
"""把 pptx 渲染成 HTML 预览（形状、文字、图片、表格、图表近似渲染），用于交付前自检。

用法: python preview_ppt.py <file.pptx> <out.html> [--accent 283E4E] [--slide 0]
说明: 图表按数据近似绘制，用于检查布局/溢出/对齐，不等于 PowerPoint 的最终渲染。
"""
import sys, os, base64, html as _h, argparse
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_FILL

ap = argparse.ArgumentParser()
ap.add_argument("src"); ap.add_argument("out")
ap.add_argument("--accent", default="283E4E")
ap.add_argument("--slide", type=int, default=0)
ap.add_argument("--band", default="F1F5F8")
args = ap.parse_args()
ACCENT = "#" + args.accent.lstrip("#").upper()
BAND = "#" + args.band.lstrip("#").upper()

prs = Presentation(args.src); sl = prs.slides[args.slide]
EMU = 914400.0; DPI = 96.0
W = prs.slide_width/EMU*DPI; H = prs.slide_height/EMU*DPI

def rgb(c, d="000000"):
    try: return str(c.rgb) if c is not None and c.type is not None else d
    except Exception: return d

def chart_html(sh, l, t, w, h):
    ch = sh.chart; ct = ch.chart_type
    try: cats = [str(c) for c in ch.plots[0].categories]
    except Exception: cats = []
    try: vals = [float(v) for v in ch.series[0].values]
    except Exception: vals = []
    body = ""
    if ct == XL_CHART_TYPE.DOUGHNUT:
        v = vals[0] if vals else 0; tot = sum(vals) or 100
        pct = v/tot; r = min(w, h)*0.36; sw = min(w, h)*0.20
        cx, cy = w/2, h/2; circ = 2*3.14159*r; dash = circ*pct
        body = ("<svg width='%.1f' height='%.1f'><circle cx='%.1f' cy='%.1f' r='%.1f' fill='none' stroke='%s' stroke-width='%.1f'/>"
                "<circle cx='%.1f' cy='%.1f' r='%.1f' fill='none' stroke='%s' stroke-width='%.1f' stroke-dasharray='%.1f %.1f' transform='rotate(-90 %.1f %.1f)'/></svg>"
                % (w, h, cx, cy, r, BAND, sw, cx, cy, r, ACCENT, sw, dash, circ-dash, cx, cy))
    elif "SCATTER" in str(ct):
        xs = []
        try:
            ser = ch._chartSpace.find(".//"+qn("c:ser"))
            if ser is not None:
                cache = ser.find(qn("c:xVal")+"/"+qn("c:numRef")+"/"+qn("c:numCache"))
                if cache is not None:
                    for pt in cache.findall(qn("c:pt")):
                        v = pt.find(qn("c:v"))
                        if v is not None and v.text:
                            xs.append(float(v.text))
        except Exception:
            xs = []
        ys = vals
        if xs and ys:
            x0, x1 = min(xs), max(xs); y0, y1 = min(ys), max(ys)
            dx = (x1-x0) or 1.0; dy = (y1-y0) or 1.0
            for x, y in zip(xs, ys):
                px = 34 + (x-x0)/dx*(w-68)
                py = h-28 - (y-y0)/dy*(h-58)
                body += ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:7px;height:7px;"
                         "border-radius:50%%;background:%s'></div>" % (px-3.5, py-3.5, ACCENT))
    else:
        mx = max([abs(v) for v in vals] + [1])
        try:
            _vm = ch.value_axis.maximum_scale
            if _vm: mx = float(_vm)
        except Exception: pass
        if ct == XL_CHART_TYPE.COLUMN_CLUSTERED:
            pad_b, pad_t = 30, 30; avail = h-pad_b-pad_t
            n = max(len(vals), 1)
            bw = w/n*0.52                       # 柱宽：留出间隔，避免出现负宽度
            gap = (w - n*bw)/(n+1)
            for i, (c, v) in enumerate(zip(cats, vals)):
                bh = max(1.0, avail*abs(v)/mx); x = gap+i*(bw+gap); y = pad_t+avail-bh
                body += ("<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:%s'></div>"
                         "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;text-align:center;font:16px bold;color:%s'>%s</div>"
                         "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;text-align:center;font:16px color:#000;white-space:nowrap'>%s</div>"
                         % (x, y, bw, bh, ACCENT, x-6, y-20, bw+12, ACCENT, '%g'%v, x-10, pad_t+avail+8, bw+20, c))
        else:
            lab_w = w*0.34; bar_x = lab_w+6; avail = w-bar_x-8; rowh = h/max(len(vals), 1)
            for i, (c, v) in enumerate(zip(cats, vals)):
                bh = min(rowh*0.5, 44); y = i*rowh+(rowh-bh)/2; bwd = max(2.0, avail*abs(v)/mx)
                body += ("<div style='position:absolute;left:0;top:%.1fpx;width:%.1fpx;text-align:right;font:16px color:#000;white-space:nowrap'>%s</div>"
                         "<div style='position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;background:%s'></div>"
                         "<div style='position:absolute;left:%.1fpx;top:%.1fpx;font:16px bold;color:%s;white-space:nowrap'>%s</div>"
                         % (y+bh/2-10, lab_w, c, bar_x, y, bwd, bh, ACCENT, bar_x+bwd+6, y+bh/2-10, '%g'%v))
    return "<div class=sh style='left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx'>%s</div>" % (l, t, w, h, body)

parts = ["<!DOCTYPE html><html lang=zh-CN><head><meta charset=utf-8><style>",
 "*{box-sizing:border-box;margin:0;padding:0}body{background:#e9edf0;font-family:'Microsoft YaHei',sans-serif}",
 ".slide{position:relative;width:%.0fpx;height:%.0fpx;background:#fff;box-shadow:0 6px 24px rgba(30,50,64,.18);overflow:hidden}" % (W, H),
 ".sh{position:absolute;overflow:hidden}.tx{position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center}",
 ".pg{width:100%}span{white-space:pre-wrap}</style></head><body><div class=slide>"]

for sh in sl.shapes:
    l = sh.left/EMU*DPI; t = sh.top/EMU*DPI; w = sh.width/EMU*DPI; h = sh.height/EMU*DPI
    st = "left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx" % (l, t, w, h)
    if getattr(sh, "has_table", False):
        tb = sh.table; rows = []
        for r in tb.rows:
            cells = []
            for c in r.cells:
                bg = "#ffffff"
                try:
                    if c.fill.type == MSO_FILL.SOLID: bg = "#" + rgb(c.fill.fore_color)
                except Exception: pass
                runs = ""
                for para in c.text_frame.paragraphs:
                    for rr in para.runs:
                        sz = rr.font.size.pt if rr.font.size else 16
                        nm = rr.font.name or ""
                        face = "SimSun" if "宋" in nm else ("KaiTi" if "楷" in nm else "Microsoft YaHei")
                        al = {1:"left",2:"center",3:"right"}.get(int(para.alignment or 1),"left")
                        runs += ("<div style='text-align:%s;font-family:\"%s\";font-size:%.2fpx;color:#%s;font-weight:%s'>%s</div>"
                                 % (al, face, sz*DPI/72, rgb(rr.font.color), "bold" if nm=="黑体" else "normal", _h.escape(rr.text)))
                cells.append("<td style='background:%s;padding:6px 10px;vertical-align:middle'>%s</td>" % (bg, runs))
            rows.append("<tr style='height:%.1fpx'>%s</tr>" % (r.height/EMU*DPI, "".join(cells)))
        parts.append("<div class=sh style='%s'><table style='border-collapse:collapse;width:100%%;table-layout:fixed'>%s</table></div>" % (st, "".join(rows)))
        continue
    if getattr(sh, "has_chart", False):
        parts.append(chart_html(sh, l, t, w, h)); continue
    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
        b64 = base64.b64encode(sh.image.blob).decode()
        nw, nh = sh.image.size; cl, ct2, cr, cb = sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom
        vw = nw*(1-cl-cr); sc = w/vw if vw else 1
        parts.append("<div class=sh style='%s;overflow:hidden'><img style='position:absolute;left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx' src='data:image/jpeg;base64,%s'></div>"
                     % (st, -cl*nw*sc, -ct2*nh*sc, nw*sc, nh*sc, b64)); continue
    fill = None; border = ""
    try:
        if sh.fill.type == MSO_FILL.SOLID: fill = rgb(sh.fill.fore_color)
    except Exception: pass
    try:
        if sh.line.fill.type == MSO_FILL.SOLID: border = "border:1px solid #%s" % rgb(sh.line.color)
    except Exception: pass
    clip = ""
    try:
        at = sh.auto_shape_type
        if at == MSO_SHAPE.OVAL: clip = "border-radius:50%;"
        elif at == MSO_SHAPE.ISOSCELES_TRIANGLE: clip = "clip-path:polygon(50% 0,100% 100%,0 100%);"
        elif at == MSO_SHAPE.HEXAGON: clip = "clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%);"
    except Exception: pass
    rot = getattr(sh, "rotation", 0) or 0
    if fill: st += ";background:#" + fill
    if border: st += ";" + border
    if clip: st += ";" + clip
    if rot: st += ";transform:rotate(%.1fdeg);transform-origin:50%% 50%%" % rot
    if sh.has_text_frame and sh.text_frame.text.strip():
        paras = ""
        for p in sh.text_frame.paragraphs:
            runs = ""
            for rr in p.runs:
                sz = rr.font.size.pt if rr.font.size else 16
                nm = rr.font.name or ""
                face = "SimSun" if "宋" in nm else ("KaiTi" if "楷" in nm else "Microsoft YaHei")
                runs += "<span style='font-family:\"%s\";font-size:%.2fpx;color:#%s;font-weight:%s'>%s</span>" % (
                    face, sz*DPI/72, rgb(rr.font.color), "bold" if nm == "黑体" else "normal", _h.escape(rr.text))
            al = {1: "left", 2: "center", 3: "right"}.get(int(p.alignment or 1), "left")
            paras += "<div class=pg style='line-height:1.25;text-align:%s'>%s</div>" % (al, runs)
        parts.append("<div class=sh style='%s'><div class=tx'>%s</div></div>" % (st, paras))
    else:
        parts.append("<div class=sh style='%s'></div>" % st)

parts.append("</div></body></html>")
open(args.out, "w", encoding="utf-8").write("".join(parts))
print("html", args.out)
