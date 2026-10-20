# -*- coding: utf-8 -*-
"""生成中性示例页 assets/example-slide.pptx（两页：方块+图示页 / 表格页）。
示例内容全部为占位文案与虚构数据，不含任何真实项目材料。"""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_MARK
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "assets", "example-slide.pptx")
NAVY="283E4E"; BLACK="000000"; WHITE="FFFFFF"; BAND="F1F5F8"; LINE="B7C1C8"
HEI="黑体"; KAI="楷体"; SONG="宋体"

def font(run, size, color, face):
    run.font.size=Pt(size); run.font.color.rgb=RGBColor.from_string(color); run.font.name=face
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea","a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag),{}); rPr.insert_element_before(el,"a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:endParaRPr")
        el.set("typeface", face)

def rect(sl,l,t,w,h,fill=None,line=None,lw=0.75,shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape,Inches(l),Inches(t),Inches(w),Inches(h)); s.shadow.inherit=False
    if fill: s.fill.solid(); s.fill.fore_color.rgb=RGBColor.from_string(fill)
    else: s.fill.background()
    if line: s.line.color.rgb=RGBColor.from_string(line); s.line.width=Pt(lw)
    else: s.line.fill.background()
    s.text_frame.text=""; return s

def text(sl,l,t,w,h,segs,size,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,ls=1.0):
    """segs: [(文本, 颜色, 字体), ...]"""
    tb = sl.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    tf = tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    p = tf.paragraphs[0]; p.alignment=align; p.line_spacing=ls
    for seg in segs:
        r = p.add_run(); r.text = seg[0]; font(r, size, seg[1], seg[2])
    return tb

def header(sl, page, title, sub):
    text(sl,6.40,0.26,6.10,0.28,[("示例 · 汇报页骨架",BLACK,SONG)],16,PP_ALIGN.RIGHT)
    text(sl,12.60,0.26,0.91,0.28,[(page,NAVY,HEI)],16,PP_ALIGN.RIGHT)
    text(sl,0.55,0.62,12.96,0.52,[(title,NAVY,HEI)],25)
    text(sl,0.55,1.18,12.96,0.30,[(sub,BLACK,SONG)],16)
    rect(sl,0.55,1.62,12.96,0.012,fill=NAVY)

prs = Presentation(); prs.slide_width=12858750; prs.slide_height=7232650

# ── 第 1 页：方块 + 几何记忆点 ──
s1 = prs.slides.add_slide(prs.slide_layouts[6])
header(s1,"01","示例页：左文右图与几何记忆点",
       "本示例仅演示版式与视觉规范，数值均为虚构占位数据。")
rect(s1,0.55,1.74,7.10,0.95,fill=BAND); rect(s1,0.55,1.74,0.045,0.95,fill=NAVY)
text(s1,0.76,1.84,6.74,0.75,[("这是导语块：演示位置、字号、行距与浅底强调样式，文字内容为占位示意。",BLACK,KAI)],17,anchor=MSO_ANCHOR.MIDDLE,ls=1.18)
rows=[("示例指标 A 达到 ","38.6","（占位数值）"),
      ("示例指标 B 同比变化 ","+12.4","（占位数值）"),
      ("示例指标 C 覆盖范围 ","1,200","（占位数值）")]
y=2.84
for a,b,c in rows:
    rect(s1,0.55,y,0.045,0.66,fill=NAVY)
    text(s1,0.76,y,6.84,0.66,[(a,BLACK,SONG),(b,NAVY,HEI),(c,BLACK,SONG)],17,anchor=MSO_ANCHOR.MIDDLE,ls=1.16)
    y+=0.70
rect(s1,0.55,5.06,7.10,1.05,fill=BAND); rect(s1,0.55,5.06,0.045,1.05,fill=NAVY)
text(s1,0.76,5.16,6.74,0.85,[("示例判断句：用一句话收束左侧信息，主题色加粗，作为口头汇报的落点。",NAVY,HEI)],17,anchor=MSO_ANCHOR.MIDDLE,ls=1.18)
rect(s1,7.85,1.74,5.66,1.72,fill=BAND,line=LINE)
text(s1,7.85,2.30,5.66,0.40,[("占位配图区域：实稿放实景照片，按框裁切填满。",BLACK,SONG)],16,PP_ALIGN.CENTER)
rect(s1,7.85,3.66,1.80,1.80,fill=NAVY,shape=MSO_SHAPE.OVAL)
text(s1,7.95,4.10,1.60,0.60,[("38.6%",WHITE,HEI)],26,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
text(s1,7.95,4.66,1.60,0.36,[("占位数值",WHITE,SONG)],16,PP_ALIGN.CENTER)
text(s1,9.95,3.66,3.56,0.26,[("对比图（示例）",BLACK,SONG)],16,PP_ALIGN.CENTER)
cd = CategoryChartData(); cd.categories=["甲","乙","丙"]; cd.add_series("s",(62,45,28))
gf = s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Inches(9.95),Inches(3.96),Inches(3.56),Inches(2.42),cd)
ch = gf.chart; ch.has_title=False; ch.has_legend=False
ch.font.size=Pt(16); ch.font.name=SONG; ch.font.color.rgb=RGBColor.from_string(BLACK)
cs = ch._chartSpace
if cs.find(qn("c:spPr")) is None:
    cs.find(qn("c:chart")).addnext(parse_xml('<c:spPr %s><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'%nsdecls("c","a")))
pa = cs.find(".//"+qn("c:plotArea"))
if pa is not None and pa.find(qn("c:spPr")) is None:
    pa.append(parse_xml('<c:spPr %s><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'%nsdecls("c","a")))
pl = ch.plots[0]; pl.gap_width=90; pl.has_data_labels=True
dl = pl.data_labels; dl.number_format='#,##0'; dl.number_format_is_linked=False
dl.font.size=Pt(16); dl.font.bold=True; dl.font.name=HEI; dl.font.color.rgb=RGBColor.from_string(NAVY)
dl.position=XL_LABEL_POSITION.OUTSIDE_END
sr = pl.series[0]; sr.format.fill.solid(); sr.format.fill.fore_color.rgb=RGBColor.from_string(NAVY); sr.format.line.fill.background()
ca = ch.category_axis; ca.has_major_gridlines=False; ca.major_tick_mark=XL_TICK_MARK.NONE
ca.tick_labels.font.size=Pt(16); ca.tick_labels.font.name=SONG; ca.tick_labels.font.color.rgb=RGBColor.from_string(BLACK)
ca.format.line.color.rgb=RGBColor.from_string(NAVY); ca.format.line.width=Pt(1)
va = ch.value_axis; va.has_major_gridlines=False
d = va._element.find(qn("c:delete"))
if d is None:
    d = va._element.makeelement(qn("c:delete"),{}); va._element.find(qn("c:scaling")).addnext(d)
d.set("val","1"); va.maximum_scale=80
for rPr in cs.iter(qn("a:defRPr")):
    latin = rPr.find(qn("a:latin")); face = latin.get("typeface") if latin is not None else SONG
    for tag in ("a:ea","a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag),{}); rPr.insert_element_before(el,"a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:endParaRPr")
        el.set("typeface", face)
rect(s1,0.55,6.76,0.045,0.56,fill=NAVY)
text(s1,0.76,6.76,12.60,0.56,[("示例结论条：用一句话写清证据强弱与边界，作为整页的落点。",BLACK,SONG)],16,anchor=MSO_ANCHOR.MIDDLE,ls=1.1)
s1.notes_slide.notes_text_frame.text = "【示例页】仅用于演示版式，不是真实内容。讲稿结构：导语 → 三条实证 → 一句判断 → 结论条。"

# ── 第 2 页：强数据用表格 ──
s2 = prs.slides.add_slide(prs.slide_layouts[6])
header(s2,"02","示例页：强数据内容用表格承载",
       "指标对比、口径对照、多维数据用表格；叙事与机制内容用方块。")
DATA=[["指标","2024","2025","变化","口径"],
      ["指标甲","1,180","1,420","+20.3%","示例口径"],
      ["指标乙","62.1","58.4","−3.7","示例口径"],
      ["指标丙","340","512","+50.6%","示例口径"],
      ["指标丁","9.6","12.8","+33.3%","示例口径"]]
shape = s2.shapes.add_table(len(DATA),5,Inches(0.55),Inches(1.86),Inches(12.96),Inches(3.30))
tbl = shape.table
for i,w in enumerate([3.00,2.10,2.10,2.10,3.66]): tbl.columns[i].width=Inches(w)
for ri,row in enumerate(DATA):
    tbl.rows[ri].height=Inches(0.66)
    for ci,val in enumerate(row):
        cell = tbl.cell(ri,ci)
        cell.margin_left=Inches(0.14); cell.margin_right=Inches(0.10)
        cell.margin_top=Inches(0.05); cell.margin_bottom=Inches(0.05)
        cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        if ri==0: cell.fill.fore_color.rgb=RGBColor.from_string(NAVY)
        else: cell.fill.fore_color.rgb=RGBColor.from_string(BAND if ri%2==0 else WHITE)
        tf=cell.text_frame; tf.word_wrap=True
        p=tf.paragraphs[0]; p.alignment=PP_ALIGN.LEFT if ci in (0,4) else PP_ALIGN.RIGHT
        r=p.add_run(); r.text=val
        if ri==0: font(r,16,WHITE,HEI)
        elif ci in (1,2,3): font(r,16,NAVY,HEI)
        else: font(r,16,BLACK,SONG)
rect(s2,0.55,5.60,0.045,0.90,fill=NAVY)
text(s2,0.76,5.60,12.60,0.90,[("表格页同样遵守字号下限与三色限制：表头主题色底白字，关键数值主题色，正文黑色。",BLACK,SONG)],17,anchor=MSO_ANCHOR.MIDDLE,ls=1.18)
s2.notes_slide.notes_text_frame.text = "【表格示例页】演示强数据用表格的样式：表头深色、交替行底纹、关键数值主题色。"

prs.save(OUT)
print("saved", os.path.normpath(OUT))

