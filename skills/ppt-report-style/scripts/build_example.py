# -*- coding: utf-8 -*-
"""生成 6 页示例 assets/example-slide.pptx（封面 / 章节页 / 左文右图 / 表格 / 四宫格 / 收束页）。
画布 13.33 x 7.5 in（PowerPoint 默认 16:9），页面骨架网格由 14.06 x 7.91 版按 0.948 等比换算。
示例内容为"歌声质量评估模型调研"虚构主题，全部数值为虚构占位数据，不含任何真实项目材料。"""
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
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "assets", "example-slide.pptx")
# 配色：Okabe-Ito（Wong palette，Nature Methods 2011）成套学术色板 + 中性灰脚手架
# 文字对比度全部 ≥4.5:1：黑 21:1 / 次级灰 6E6E6E 5.1:1 / 点缀蓝 0072B2 5.2:1
ACC="0072B2"; BLACK="000000"; MUTED="6E6E6E"; HAIR="D9D9D9"; BAR="BFBFBF"; WHITE="FFFFFF"
HEI="黑体"; KAI="楷体"; SONG="宋体"
# 默认学术风：正文宋体 + 标题黑体；西文正文配 Times New Roman，黑体标题内西文配 Arial。
# 微软雅黑保留在白名单，作为企业/演示场合的可选方案。
LATIN_OF = {HEI: "Arial", SONG: "Times New Roman", KAI: "Times New Roman"}

# 13.33x7.5 画布骨架：左边距 0.52，正文宽 12.29（左右边距对称 0.52）。
# 页眉 T=0.22 H=0.36 -> 标题 T=0.53 H=0.63(30pt) -> 副标 T=1.19 H=0.44(20pt)
# -> 细线 T=1.67 -> 正文区 1.86-6.32 -> 结论条 T=6.45 H=0.60。
ML=0.52; CW=12.29

def font(run, size, color, face, latin=None):
    """face = 中文字形（a:ea / a:cs）；latin = 西文字形（a:latin），默认按 face 配对。"""
    lat = latin or LATIN_OF.get(face, "Arial")
    run.font.size=Pt(size); run.font.color.rgb=RGBColor.from_string(color); run.font.name=lat
    rPr = run._r.get_or_add_rPr()
    for tag, tf in (("a:ea", face), ("a:cs", face)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag),{}); rPr.insert_element_before(el,"a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:endParaRPr")
        el.set("typeface", tf)

def rect(sl,l,t,w,h,fill=None,line=None,lw=0.75,shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape,Inches(l),Inches(t),Inches(w),Inches(h)); s.shadow.inherit=False
    _style = s._element.find(qn("p:style"))          # 拆掉主题样式引用：effectRef idx=2 指向主题外阴影，
    if _style is not None: s._element.remove(_style)  # 空 effectLst 在 WPS/部分 PowerPoint 下压不住，必须整块删除
    if fill: s.fill.solid(); s.fill.fore_color.rgb=RGBColor.from_string(fill)
    else: s.fill.background()
    if line: s.line.color.rgb=RGBColor.from_string(line); s.line.width=Pt(lw)
    else: s.line.fill.background()
    s.text_frame.text=""; return s

def text(sl,l,t,w,h,paras,size,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,ls=1.0):
    """paras: [[(文本, 颜色, 字体), ...], ...] —— 外层列表是段落，内层是同段内的 run。"""
    tb = sl.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    tb.shadow.inherit=False   # 不继承任何阴影效果
    tf = tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    if paras and not isinstance(paras[0], list): paras=[paras]
    for i,segs in enumerate(paras):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.line_spacing=ls
        for seg in segs:
            r = p.add_run(); r.text = seg[0]; font(r, size, seg[1], seg[2])
    return tb

def header(sl, chapter="第一章 · 调研范围与方法"):
    text(sl,ML,0.22,9.00,0.36,[(chapter,HAIR,SONG)],14,anchor=MSO_ANCHOR.MIDDLE)   # 页眉：14pt 极浅灰，退后
    rect(sl,ML,0.66,CW,0.005,fill=HAIR)   # 页眉线：0.36pt 特细发丝线，无任何效果

def footer(sl, page, title="歌声质量评估模型调研汇报"):
    rect(sl,ML,7.00,CW,0.005,fill=HAIR)   # 页脚线：同样特细
    text(sl,ML,7.08,9.00,0.34,[(title,HAIR,SONG)],14,anchor=MSO_ANCHOR.MIDDLE)     # 页脚：14pt 极浅灰
    text(sl,10.32,7.08,2.49,0.34,[(page,HAIR,HEI)],14,PP_ALIGN.RIGHT,anchor=MSO_ANCHOR.MIDDLE)

def title_block(sl, title, sub):
    text(sl,ML,0.98,CW,0.63,[(title,BLACK,HEI)],30)   # 与页眉线保持 0.32in 距离
    text(sl,ML,1.96,CW,0.44,[(sub,BLACK,SONG)],20)

def conclusion(sl, segs):
    rect(sl,ML,6.15,CW,0.01,fill=HAIR)
    text(sl,ML,6.27,CW,0.44,[segs],18,anchor=MSO_ANCHOR.MIDDLE,ls=1.1)   # 与页脚线保持 0.29in 距离

def band_block(sl,l,t,w,h,paras,size,face=SONG,color=BLACK,ls=1.2):
    """浅底强调块：F1F5F8 底 + 左侧主题色竖条。"""
    rect(sl,l,t,w,h,fill=BAND); rect(sl,l,t,0.045,h,fill=ACC)
    text(sl,l+0.21,t,w-0.42,h,paras,size,anchor=MSO_ANCHOR.MIDDLE,ls=ls)
    return

def chart_cn_font(ch):
    cs = ch._chartSpace
    for rPr in cs.iter(qn("a:defRPr")):
        latin = rPr.find(qn("a:latin"))
        if latin is None:
            latin = rPr.makeelement(qn("a:latin"),{}); rPr.insert_element_before(latin,"a:ea","a:cs","a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:endParaRPr")
        latin.set("typeface", latin.get("typeface") or "Arial")
        for tag in ("a:ea","a:cs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = rPr.makeelement(qn(tag),{}); rPr.insert_element_before(el,"a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:endParaRPr")
            el.set("typeface", SONG)

prs = Presentation(); prs.slide_width=12192000; prs.slide_height=6858000
BLANK = prs.slide_layouts[6]

# ══ 第 1 页：封面 ══
s1 = prs.slides.add_slide(BLANK)
text(s1,ML,0.22,6.00,0.36,[("虚构示例 · 仅供版式参考",HAIR,SONG)],14,anchor=MSO_ANCHOR.MIDDLE)
text(s1,6.82,0.22,5.99,0.36,[("组会汇报",HAIR,SONG)],14,PP_ALIGN.RIGHT,anchor=MSO_ANCHOR.MIDDLE)
rect(s1,ML,0.66,CW,0.005,fill=HAIR)
text(s1,ML,2.55,11.80,0.95,[("歌声质量评估模型调研汇报",BLACK,HEI)],44)
text(s1,ML,3.80,11.80,0.45,[("2022-2026 方法谱系与开源现状 · 虚构示例数据",MUTED,SONG)],20)
text(s1,ML,5.10,11.80,0.42,[("汇报人：张三 · 教育技术系 · 语音质量评估组",MUTED,SONG)],18)
text(s1,ML,5.62,11.80,0.42,[("2026 年 10 月 · 课题组组会",MUTED,SONG)],18)
footer(s1,"01 / 06")
s1.notes_slide.notes_text_frame.text = """【封面 · 讲稿】
本文件是版式示例，主题与全部数值均为虚构，不要当作真实调研结论引用。
开场约二十秒：先自我介绍，再用一句话交代调研缘起——歌声质量自动评估近四年方法更替很快，组内需要一份按方法范式梳理的谱系图，这页就是给这件事定题。
封面只保留主标题、点缀短线、副标题与汇报人信息，落款用浅灰退后，不放任何装饰图形。
如果实稿需要放单位 Logo，放在页眉右侧与"组会汇报"同排，不要另起一行。"""

# ══ 第 2 页：章节页（大号章号 + 楷体导语，右侧留白）══
s2 = prs.slides.add_slide(BLANK)
header(s2)
footer(s2,"02 / 06")
text(s2,ML,2.20,3.00,0.60,[("01",MUTED,HEI)],36)
text(s2,ML,3.30,7.20,0.80,[("调研范围与方法",BLACK,HEI)],36)
text(s2,ML,4.50,5.90,1.70,
    [[("检索 Web of Science 与 arXiv（示例口径），时间窗 2022-2026；纳入标准：以歌声质量为评估对象、报告与人工评分一致性指标的实证研究。",BLACK,KAI)]],
    20,ls=1.35)
rect(s2,7.90,2.30,0.01,1.70,fill=HAIR)   # 导航分隔发丝线
text(s2,8.30,2.30,4.50,0.44,[("01  调研范围与方法",ACC,HEI)],18,anchor=MSO_ANCHOR.MIDDLE)
text(s2,8.30,2.85,4.50,0.44,[("02  方法指标对比",MUTED,SONG)],18,anchor=MSO_ANCHOR.MIDDLE)
text(s2,8.30,3.40,4.50,0.44,[("03  结论与计划",MUTED,SONG)],18,anchor=MSO_ANCHOR.MIDDLE)
s2.notes_slide.notes_text_frame.text = """【章节页 · 讲稿】
转场约二十秒：说明本章讲清两件事——去哪些库检索、按什么标准纳入。
右侧是全片章节导航：当前章用点缀色，其余章浅灰，翻页时它是路标不是装饰；导语为楷体，是全篇唯一的楷体位置，检索库、时间窗与纳入标准在实稿中必须替换为真实口径。
备答：如果被问到为什么不收录某类会议，回答口径写进备注即可，不上片。"""

# ══ 第 3 页：左文右图正文页（横向条形图 + 关键数字入组）══
s3 = prs.slides.add_slide(BLANK)
header(s3)
footer(s3,"03 / 06")
title_block(s3,"SSL 预训微调路线一致性最优，LCC 达 0.86",
            "四类方法横向对比")
rows=[("调研文献 ","24"," 篇，时间窗 2022-2026，覆盖四类方法范式"),
      ("SSL 微调路线 LCC ","0.86","、SRCC 0.84，为四类最高"),
      ("24 篇中 ","9"," 篇开源可复现，占比 37.5%")]
y=2.70
for a,b,c in rows:
    text(s3,ML,y,6.73,0.44,[(a,BLACK,SONG),(b,BLACK,HEI),(c,BLACK,SONG)],18,anchor=MSO_ANCHOR.MIDDLE,ls=1.2)
    y+=0.85
rect(s3,ML,5.45,6.73,0.01,fill=HAIR)
text(s3,ML,5.57,6.73,0.44,[("一致性领先约 ",BLACK,HEI),("3-8 个千分点",BLACK,HEI),("，优势明确但并非断层。",BLACK,HEI)],18,anchor=MSO_ANCHOR.MIDDLE)
cd = CategoryChartData(); cd.categories=["LLM 评审","谱图 CNN","多任务融合","SSL 微调"]   # 条形图自下而上，SSL 置顶
cd.add_series("LCC",(0.71,0.78,0.83,0.86))
gf = s3.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,Inches(7.44),Inches(2.70),Inches(5.37),Inches(2.30),cd)
ch = gf.chart; ch.has_title=False; ch.has_legend=False
ch.font.size=Pt(16); ch.font.name="Arial"; ch.font.color.rgb=RGBColor.from_string(BLACK)
cs = ch._chartSpace
if cs.find(qn("c:spPr")) is None:
    cs.find(qn("c:chart")).addnext(parse_xml('<c:spPr %s><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'%nsdecls("c","a")))
pa = cs.find(".//"+qn("c:plotArea"))
if pa is not None and pa.find(qn("c:spPr")) is None:
    pa.append(parse_xml('<c:spPr %s><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'%nsdecls("c","a")))
pl = ch.plots[0]; pl.gap_width=90; pl.has_data_labels=True
dl = pl.data_labels; dl.number_format='0.00'; dl.number_format_is_linked=False
dl.font.size=Pt(16); dl.font.bold=False; dl.font.name="Arial"; dl.font.color.rgb=RGBColor.from_string(BLACK)
dl.position=XL_LABEL_POSITION.OUTSIDE_END
sr = pl.series[0]
for idx,pt in enumerate(sr.points):        # 只点亮最重要的一根：SSL 微调（置顶）
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb=RGBColor.from_string(ACC if idx==3 else BAR)
    pt.format.line.fill.background()
ca = ch.category_axis; ca.has_major_gridlines=False; ca.major_tick_mark=XL_TICK_MARK.NONE
ca.tick_labels.font.size=Pt(16); ca.tick_labels.font.name="Arial"; ca.tick_labels.font.color.rgb=RGBColor.from_string(BLACK)
ca.format.line.color.rgb=RGBColor.from_string(HAIR); ca.format.line.width=Pt(1)
va = ch.value_axis                          # 数值轴窗口取数据范围的截断区间并诚实标注
va.minimum_scale=0.60; va.maximum_scale=0.90; va.major_unit=0.05
va.has_major_gridlines=True
va.major_gridlines.format.line.color.rgb=RGBColor.from_string(HAIR)
va.major_gridlines.format.line.width=Pt(0.75)
va.tick_labels.font.size=Pt(16); va.tick_labels.font.name="Arial"; va.tick_labels.font.color.rgb=RGBColor.from_string(MUTED)
va.tick_labels.number_format='0.00'; va.tick_labels.number_format_is_linked=False
va.format.line.fill.background()
chart_cn_font(ch)
text(s3,7.44,5.20,5.37,0.80,[[("图注：SSL 路线一致性最高（LCC 0.86）。",MUTED,SONG)],
    [("数值轴 0.60-0.90，截断显示 · 示例数据",MUTED,SONG)]],16,ls=1.25)
s3.notes_slide.notes_text_frame.text = """【左文右图正文页 · 讲稿】
本页约一分钟，节奏是：三条实证逐条讲，判断行收束。
三条实证都是"指标名 + 数值 + 口径"写在同一句里：篇数 24、最优 LCC 0.86、开源占比 37.5%。
右栏是横向条形图，按数值排序、SSL 置顶并点亮；数值轴取 0.60-0.90 的截断窗口让差距可读，图注写明结论与截断口径，避免误读。
备答：0.86 是虚构口径，实稿替换为真实测试集结果时，务必把测试集与评测脚本来源写进备注。"""

# ══ 第 4 页：表格页（杂志式三线表 + 关键数字同组）══
s4 = prs.slides.add_slide(BLANK)
header(s4)
footer(s4,"04 / 06")
title_block(s4,"四类方法范式的指标与开源现状对比",
            "同一测试集内对比")
DATA=[["方法范式","论文篇数","LCC","SRCC","开源仓库数"],
      ["谱图 CNN 基线","8","0.78","0.76","3"],
      ["SSL 预训微调","7","0.86","0.84","4"],
      ["LLM-as-critic","5","0.71","0.69","1"],
      ["多任务融合","4","0.83","0.81","1"],
      ["合计","24","—","—","9"]]
shape = s4.shapes.add_table(len(DATA),5,Inches(ML),Inches(2.55),Inches(12.29),Inches(3.12))
tbl = shape.table
for i,w in enumerate([3.60,2.20,1.80,1.80,2.89]): tbl.columns[i].width=Inches(w)
for ri,row in enumerate(DATA):
    tbl.rows[ri].height=Inches(0.52)
    for ci,val in enumerate(row):
        cell = tbl.cell(ri,ci)
        cell.margin_left=Inches(0.14); cell.margin_right=Inches(0.10)
        cell.margin_top=Inches(0.05); cell.margin_bottom=Inches(0.05)
        cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        cell.fill.solid(); cell.fill.fore_color.rgb=RGBColor.from_string(WHITE)
        tf=cell.text_frame; tf.word_wrap=True
        p=tf.paragraphs[0]; p.alignment=PP_ALIGN.LEFT if ci==0 else PP_ALIGN.RIGHT
        r=p.add_run(); r.text=val
        if ri==0: font(r,16,BLACK,HEI)
        elif ri==2 and ci in (1,2,3,4): font(r,16,ACC,HEI)   # 只点亮 SSL 行
        elif ci in (1,2,3,4): font(r,16,BLACK,HEI)
        else: font(r,16,BLACK,SONG)
rect(s4,ML,2.55,12.29,0.014,fill=BLACK)           # 三线表：顶线（1pt 墨色）
rect(s4,ML,3.07,12.29,0.01,fill=BLACK)            # 表头下线
for k in (2,3,4,5):                               # 行间发丝线
    rect(s4,ML,2.55+0.52*k,12.29,0.01,fill=HAIR)
rect(s4,ML,5.67,12.29,0.014,fill=BLACK)           # 底线（1pt 墨色）
text(s4,ML,5.90,CW,0.70,[[("注：LCC 与 SRCC 为与人工评分的一致性指标；整体开源率 ",MUTED,SONG),("37.5%",BLACK,HEI),("（9/24）。",MUTED,SONG)],
    [("数值均为示例数据；表格篇数合计 24、仓库数合计 9。",MUTED,SONG)]],16,ls=1.25)
s4.notes_slide.notes_text_frame.text = """【表格页 · 讲稿】
本页约一分钟，不要逐格念表：先说这张表回答什么问题，再点两三个关键数字。
关键数字三个：SSL 行的 0.86 一致性最高，该行数值用点缀色点亮，其余行保持墨色；合计行 24 篇里只有 9 个开源仓库；整体开源率 37.5% 写进表注并用黑体强调，不做独立大字。
表格是杂志式三线表：1pt 墨色顶线与底线、表头下线，行间发丝线，无底色无竖线，通栏铺满正文宽度。
表注是口径说明的位置，LCC 与 SRCC 的定义、合计口径都放在表注与备注里，不上正文。
备答：若被问"为什么多任务融合篇数少"，答纳入标准里要求报告一致性指标，部分融合类工作只有主观试听，未收录。"""

# ══ 第 5 页：方法对比页（发丝边框卡片 + 功能性分类标记）══
s5 = prs.slides.add_slide(BLANK)
header(s5)
footer(s5,"05 / 06")
title_block(s5,"四类方法范式各有明确的适用边界",
            "每格一句适用性判断")
BLOCKS=[("谱图 CNN 基线","起点路线，特征可解释；一致性上限偏低，适合作对照组。",MSO_SHAPE.OVAL),
        ("SSL 预训微调","当前一致性最高（LCC 0.86）；依赖大规模预训，部署成本高。",MSO_SHAPE.RECTANGLE),
        ("LLM 评审","零样本即可用；与人工一致性待提升，适合快速粗筛场景。",MSO_SHAPE.ISOSCELES_TRIANGLE),
        ("多任务融合","加辅助任务稳态提升 2-3 个千分点；训练成本随任务数上升。",MSO_SHAPE.DIAMOND)]
BW=6.045; CH=1.60
for i,(bt,body,mk) in enumerate(BLOCKS):
    bx = ML + (i%2)*(BW+0.20); by = 2.70 + (i//2)*1.80
    rect(s5,bx,by,BW,CH,fill=None,line=HAIR,lw=1.0)   # 发丝边框卡片，无底色
    rect(s5,bx+0.24,by+0.28,0.26,0.26,shape=mk,line=MUTED,lw=1.5)   # 功能性分类标记：形状区分类别
    text(s5,bx+0.66,by+0.24,BW-0.90,0.40,[(bt,BLACK,HEI)],18)
    text(s5,bx+0.24,by+0.78,BW-0.48,0.75,[(body,BLACK,SONG)],18,ls=1.2)
s5.notes_slide.notes_text_frame.text = """【方法对比页 · 讲稿】
本页约一分钟，每格一句话，按 01 到 04 的顺序讲。
四格是四张发丝边框卡片：左上角的小几何标记（圆/方/三角/菱形）是功能性分类标记，用来区分四类方法，不是装饰；块内直接写内容，无底色。
备答：若被问四类之外还有哪些路线，可口头提无参考评估与域自适应两类，说明未纳入的原因（一致性指标报告不全），口径进备注。"""

# ══ 第 6 页：收束页（纯排版结论 + 发丝线计划块）══
s6 = prs.slides.add_slide(BLANK)
header(s6)
footer(s6,"06 / 06")
title_block(s6,"主要发现：SSL 路线领先，复现与口径是短板",
            "三条结论与一条下阶段计划")
CONC=[[("SSL 预训微调一致性最高（LCC ",BLACK,SONG),("0.86",BLACK,HEI),("、SRCC 0.84），四类中领先。",BLACK,SONG)],
      [("24 篇中仅 ",BLACK,SONG),("9",BLACK,HEI),(" 篇开源可复现，代码与权重双开源的更少。",BLACK,SONG)],
      [("系统级与语句级评估口径混用，跨论文数值对比仍不可靠。",BLACK,SONG)]]
y=2.70
for segs in CONC:
    text(s6,ML,y,CW,0.44,[segs],18,anchor=MSO_ANCHOR.MIDDLE)
    y+=0.68
rect(s6,ML,4.90,CW,0.01,fill=HAIR)
text(s6,ML,5.05,CW,0.44,[("下阶段计划：",BLACK,HEI),("先复现 9 个开源仓库并登记口径，再统一系统级基准，最后扩展 LLM 评审路线。",BLACK,SONG)],18,anchor=MSO_ANCHOR.MIDDLE)
conclusion(s6,[("先立基准、再谈领先，是本次调研最重要的一条操作建议。",BLACK,HEI)])
s6.notes_slide.notes_text_frame.text = """【收束页 · 讲稿】
本页约四十秒：三条结论自上而下各一句，关键数字用黑体强调；计划行放在发丝线之下，结论行只保留最后一句操作建议，不带"结论/收束"前缀。
备答：下阶段计划的三步各自的时间点与人力安排写进备注，页面上只保留一句话版本。
如果被追问"复现 9 个仓库要多久"，口头回答按每周两个估算约需一个多月，具体排期会后单独同步，页面上不展开。"""

prs.save(OUT)
print("saved", os.path.normpath(OUT))
