# -*- coding: utf-8 -*-
"""Build a Persian RTL .pptx that showcases the powerpoint-fa skill.

Topic source: Persian Wikipedia article on the Asiatic cheetah
(یوزپلنگ آسیایی) — https://fa.wikipedia.org/wiki/یوزپلنگ_آسیایی
"""
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Visual system
# ----------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)

# Semantic tokens
SAND      = RGBColor(0xF7, 0xF2, 0xE9)   # background
INK       = RGBColor(0x21, 0x2B, 0x26)   # primary text (deep green-black)
SECONDARY = RGBColor(0x5A, 0x6B, 0x63)   # secondary text
ACCENT    = RGBColor(0xD9, 0x8E, 0x2B)   # cheetah amber
TEAL      = RGBColor(0x2E, 0x7D, 0x5B)   # conservation green
POSITIVE  = RGBColor(0x2E, 0x7D, 0x5B)
WARNING   = RGBColor(0xD9, 0x8E, 0x2B)
NEGATIVE  = RGBColor(0xB3, 0x3A, 0x2E)   # danger red
CARD      = RGBColor(0xFF, 0xFF, 0xFF)
CARDBORDER= RGBColor(0xE3, 0xDA, 0xC8)
FAINT     = RGBColor(0x8A, 0x97, 0x90)
AMBER_SOFT= RGBColor(0xF7, 0xE6, 0xC8)   # light amber tint
GREEN_SOFT= RGBColor(0xE3, 0xEF, 0xE7)   # light green tint
RED_SOFT  = RGBColor(0xF6, 0xE0, 0xDC)   # light red tint
DARKPANEL = RGBColor(0x1B, 0x2A, 0x24)   # deep panel for title slide

FONT = "Vazirmatn"
FARSI = MSO_LANGUAGE_ID.FARSI
EN = MSO_LANGUAGE_ID.ENGLISH_US

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
blank = prs.slide_layouts[6]

cp = prs.core_properties
cp.title = "یوزپلنگ آسیایی — نمایش توانمندی‌های مهارت powerpoint-fa"
cp.subject = "نمایش توانمندی‌های ساخت ارائه فارسی راست‌به‌چپ با python-pptx"
cp.author = "Hermes Agent (powerpoint-fa skill)"
cp.keywords = "pptx, persian, farsi, rtl, powerpoint-fa, یوزپلنگ"

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def set_run_font(run, size=18, bold=False, color=INK, font=FONT, lang=FARSI):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = font
    f.language_id = lang
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)


def add_textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return tb, tf


def add_para(tf, runs, align=PP_ALIGN.RIGHT, rtl=True, space_after=6,
             space_before=0, line=None, first=False, size=None, bold=None,
             color=None, bullet=None):
    """runs: list of (text, dict-overrides) OR a plain string.

    size/bold/color given here are defaults applied to every run; per-run
    overrides in the dict win. bullet=True uses a real buChar marker
    (slide-rules-fa §7.3) instead of typed "• " text; bullet=False clears
    to buNone for single-statement blocks.
    """
    if isinstance(runs, str):
        runs = [(runs, {})]
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    if line is not None:
        p.line_spacing = line
    pPr = p._p.get_or_add_pPr()
    if rtl:
        pPr.set("rtl", "1")
    if bullet is True:
        for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))
    elif bullet is False:
        for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
    for text, ov in runs:
        r = p.add_run()
        r.text = text
        set_run_font(
            r,
            size=ov.get("size", size if size is not None else 18),
            bold=ov.get("bold", bold if bold is not None else False),
            color=ov.get("color", color if color is not None else INK),
            font=ov.get("font", FONT),
            lang=ov.get("lang", FARSI),
        )
    return p


def add_rect(slide, x, y, w, h, fill, line=None, round_=False, radius=0.06):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE, x, y, w, h
    )
    if fill is None:
        shp.fill.background()          # transparent fill
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    if round_:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    return shp


def slide_bg(slide, color=SAND):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, color)


def header(slide, title_text, kicker=None, num=None, total=None, accent=ACCENT):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.14), accent)
    y = Inches(0.42)
    _, tf = add_textbox(slide, MARGIN, y, SLIDE_W - 2 * MARGIN, Inches(1.1))
    if kicker:
        add_para(tf, kicker, size=13, bold=True, color=TEAL,
                 space_after=2, first=True)
    add_para(tf, title_text, size=30, bold=True, color=INK, space_after=0)
    add_rect(slide, MARGIN, Inches(7.06), SLIDE_W - 2 * MARGIN, Pt(1), CARDBORDER)
    _, ftf = add_textbox(slide, MARGIN, Inches(7.12), SLIDE_W - 2 * MARGIN, Inches(0.3))
    add_para(ftf, "یوزپلنگ آسیایی · نمایش مهارت powerpoint-fa",
             size=10, color=FAINT, align=PP_ALIGN.LEFT, rtl=False, first=True)
    if num is not None:
        _, ntf = add_textbox(slide, SLIDE_W - MARGIN - Inches(1.2), Inches(7.12),
                             Inches(1.2), Inches(0.3))
        add_para(ntf, f"{num:02d} / {total}",
                 size=10, color=FAINT, align=PP_ALIGN.RIGHT, rtl=False, first=True)


def add_image_alt(pic, name, alt):
    nv = pic._element.nvPicPr
    nv.cNvPr.set("name", name)
    nv.cNvPr.set("descr", alt)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def set_chart_font(chart, font=FONT):
    """Force every text element in the chart (title, axes, labels, legend) to
    use the Persian font. Charts carry their own theme font, so the deck font
    is NOT inherited — it must be set explicitly on the chart XML.
    We set the typeface on any existing latin/ea/cs node AND inject it into
    every a:defRPr / a:endParaRPr (axis/title defaults) so PowerPoint and
    LibreOffice both render Persian labels in the deck font."""
    cs = chart._chartSpace
    # 1) any existing font nodes
    for tag in ("a:latin", "a:ea", "a:cs"):
        for el in cs.iter(qn(tag)):
            el.set("typeface", font)
    # 2) default run-properties nodes (axis ticks, title, labels)
    for tag in ("a:defRPr", "a:endParaRPr"):
        for rPr in cs.iter(qn(tag)):
            for f in ("a:latin", "a:ea", "a:cs"):
                child = rPr.find(qn(f))
                if child is None:
                    child = rPr.makeelement(qn(f), {})
                    rPr.append(child)
                child.set("typeface", font)


def list_cards(slide, items, accent_fn, y0=Inches(1.6), card_h=Inches(0.94),
               gap=Inches(0.14)):
    """Stack of RTL cards. Each card is drawn FIRST (so its color sits behind
    the text) with a colored accent strip on the RIGHT (the RTL start side).
    items: list of (title, body). accent_fn(i) -> accent color.
    Geometry is chosen so the stack ends above the footer (y < 7.06)."""
    x = MARGIN
    w = SLIDE_W - 2 * MARGIN
    y = y0
    strip_w = Inches(0.16)
    for i, (title, body) in enumerate(items):
        accent = accent_fn(i)
        add_rect(slide, x, y, w, card_h, CARD, line=CARDBORDER, round_=True)
        # accent strip on the RIGHT, inside the card, behind the text zone
        add_rect(slide, x + w - strip_w, y + Inches(0.12), strip_w,
                 card_h - Inches(0.24), accent)
        _, tf = add_textbox(slide, x + Inches(0.5), y + Inches(0.12),
                            w - Inches(0.85), card_h - Inches(0.24))
        add_para(tf, title, size=16, bold=True, color=INK, first=True, space_after=2)
        add_para(tf, body, size=13.5, color=SECONDARY, line=1.05, space_after=0)
        y += card_h + gap
    return y


# ----------------------------------------------------------------------------
# Slide 1 — Title (creative: dark panel + framed image)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
add_rect(s, 0, 0, SLIDE_W, Inches(0.16), INK)
add_rect(s, 0, Inches(7.34), SLIDE_W, Inches(0.16), INK)

# framed image on the right
pic = s.shapes.add_picture("assets/cheetah.jpg", Inches(8.6), Inches(0.95),
                           height=Inches(5.5))
add_image_alt(pic, "یوزپلنگ ایرانی (دلبر)", "تصویر یوزپلنگ ایرانی در اسارت")
add_rect(s, Inches(8.6), Inches(0.95), Inches(4.15), Inches(5.5), None,
         line=ACCENT)  # amber frame (outline only)
add_rect(s, Inches(8.6), Inches(6.45), Inches(4.15), Inches(0.14), ACCENT)

# dark content panel on the left
add_rect(s, MARGIN, Inches(1.35), Inches(7.5), Inches(4.9), DARKPANEL, round_=True,
         radius=0.05)
_, tf = add_textbox(s, MARGIN + Inches(0.6), Inches(1.75), Inches(6.3), Inches(4.3))
add_para(tf, "مهارت powerpoint-fa", size=15, bold=True, color=ACCENT, first=True,
         space_after=12)
add_para(tf, "یوزپلنگ آسیایی", size=46, bold=True, color=SAND, space_after=6)
add_para(tf, "سریع‌ترین جانور خشکی، در آستانهٔ انقراض",
         size=20, color=AMBER_SOFT, bold=True, space_after=18)
add_para(tf, "یک ارائهٔ نمونهٔ فارسی و راست‌به‌چپ که توانمندی‌های این مهارت را نشان "
             "می‌دهد: متن RTL، جدول، نمودار، تصویر، پانویس، فراپیوند و ابرداده.",
         size=15, color=RGBColor(0xC9, 0xD4, 0xCE), line=1.25, space_after=22)
p = add_para(tf, [("برگرفته از ویکی‌پدیا: ", {"color": RGBColor(0xC9, 0xD4, 0xCE)}),
                  ("fa.wikipedia.org/wiki/یوزپلنگ_آسیایی", {"color": ACCENT})],
             size=12, color=FAINT)
p.runs[1].font.underline = True

notes(s, "اسلاید عنوان. این ارائه حاصل مهارت powerpoint-fa است و بر پایهٔ مقالهٔ "
         "ویکی‌پدیای فارسی «یوزپلنگ آسیایی» ساخته شده. تصویر: یوزپلنگ دلبر.")

# ----------------------------------------------------------------------------
# Slide 2 — Conservation status (stat cards with top accent)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "وضعیت حفاظت: در بحران انقراض",
       kicker="فهرست سرخ IUCN · ردهٔ CR", num=2, total=9)

add_rect(s, MARGIN, Inches(1.6), SLIDE_W - 2 * MARGIN, Inches(1.0), INK, round_=True)
_, tf = add_textbox(s, MARGIN + Inches(0.4), Inches(1.75), SLIDE_W - 2 * MARGIN - Inches(0.8),
                    Inches(0.8), anchor=MSO_ANCHOR.MIDDLE)
add_para(tf, "گونه‌ای در بحران انقراض — Critically Endangered (CR)",
         size=22, bold=True, color=SAND, first=True)

stats = [
    ("۲۷", "قلاده برآورد رسمی (۱۴۰۵)", ACCENT),
    ("۲۰", "در طبیعت · ۶ در اسارت (۱۴۰۴)", TEAL),
    ("تنها", "بازماندهٔ این زیرگونه در ایران", NEGATIVE),
    ("۱۹۷۴", "آخرین مشاهدهٔ مستند در هند", WARNING),
]
cw = (SLIDE_W - 2 * MARGIN - 3 * Inches(0.3)) / 4
cx = MARGIN
cy = Inches(3.0)
for i, (num, label, col) in enumerate(stats):
    x = cx + i * (cw + Inches(0.3))
    add_rect(s, x, cy, cw, Inches(2.2), CARD, line=CARDBORDER, round_=True)
    add_rect(s, x, cy, cw, Inches(0.14), col)  # top accent behind content
    _, tf = add_textbox(s, x + Inches(0.2), cy + Inches(0.3), cw - Inches(0.4), Inches(1.75))
    add_para(tf, num, size=40, bold=True, color=col, align=PP_ALIGN.CENTER,
             rtl=False, space_after=10, first=True)
    add_para(tf, label, size=14, color=SECONDARY, align=PP_ALIGN.CENTER)

notes(s, "وضعیت حفاظت: یوزپلنگ آسیایی در فهرست سرخ IUCN در ردهٔ «در بحران انقراض» (CR) "
         "قرار دارد. ایران تنها کشور باقی‌ماندهٔ این زیرگونه است. منابع: فهرست سرخ IUCN، "
         "سازمان حفاظت محیط زیست ایران، ایرنا.")

# ----------------------------------------------------------------------------
# Slide 3 — Taxonomy table
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "رده‌بندی علمی", kicker="جدول دوستونی راست‌به‌چپ با نام لاتین", num=3, total=9)

rows = [
    ("فرمانرو", "جانوران"),
    ("شاخه", "طنابداران"),
    ("رده", "پستانداران"),
    ("راسته", "گوشت‌خوارسانان"),
    ("تیره", "گربه‌سانان"),
    ("سرده", "Acinonyx"),
    ("گونه", "A. jubatus"),
    ("زیرگونه", "Acinonyx jubatus venaticus"),
]
tbl_shape = s.shapes.add_table(len(rows) + 1, 2, MARGIN, Inches(1.75),
                               SLIDE_W - 2 * MARGIN, Inches(4.7))
table = tbl_shape.table
table.columns[0].width = Inches(8.7)
table.columns[1].width = Inches(3.4)
# Explicit column order for a robust RTL visual layout across renderers:
# trait/rank labels on the RIGHT (col 1), values on the LEFT (col 0).
# (LibreOffice ignores a:tblPr rtl, so we do not rely on it.)

def style_cell(cell, text, bold=False, size=16, color=INK, fill=None, rtl=True,
               lang=FARSI, align=None):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    cell.margin_left = Inches(0.12)
    cell.margin_right = Inches(0.12)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align if align is not None else (PP_ALIGN.RIGHT if rtl else PP_ALIGN.LEFT)
    if rtl:
        p._p.get_or_add_pPr().set("rtl", "1")
    r = p.add_run()
    r.text = text
    set_run_font(r, size=size, bold=bold, color=color, lang=lang)

hdr = ["ارزش", "ویژگی"]
for j, htxt in enumerate(hdr):
    style_cell(table.cell(0, j), htxt, bold=True, size=17, color=SAND, fill=INK)
for i, (k, v) in enumerate(rows, start=1):
    is_latin = all(ord(c) < 128 for c in v)
    style_cell(table.cell(i, 0), v, size=15,
               fill=CARD if i % 2 else SAND,
               rtl=not is_latin, lang=FARSI if not is_latin else EN,
               align=PP_ALIGN.LEFT if is_latin else PP_ALIGN.RIGHT)
    style_cell(table.cell(i, 1), k, size=15, color=TEAL, bold=True,
               fill=CARD if i % 2 else SAND)

notes(s, "این اسلاید جدول رده‌بندی علمی را نشان می‌دهد. ستون «ویژگی» در سمت راست و "
         "نام لاتین در سمت چپ قرار دارد؛ ترتیب ستون‌ها به‌صورت صریح تعیین شده تا در همهٔ "
         "رندرها درست دیده شود و نام‌های علمی جهت LTR و زبان انگلیسی دارند.")

# ----------------------------------------------------------------------------
# Slide 4 — Characteristics (RTL cards, green accent on the right)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "ویژگی‌ها", kicker="متن راست‌به‌چپ با واحدها و اعداد لاتین", num=4, total=9)

features = [
    ("سرعت", "سریع‌ترین جانور خشکی جهان؛ در تعقیب شکار تا ۶۰ برابر گرمای حالت استراحت تولید می‌کند."),
    ("جثه", "طول سر و بدن ۱۱۲ تا ۱۳۵ سانتی‌متر؛ دُم ۶۶ تا ۸۴ سانتی‌متر؛ وزن ۴۳ تا ۵۴ کیلوگرم."),
    ("پوشش", "پوشش کرم‌رنگ (حنایی) با لکه‌های سیاه نامنظم؛ خطوط سیاه روی سر و پشت گردن؛ نوک دُم نواری."),
    ("دودیسی جنسی", "نرها کمی بزرگ‌تر از ماده‌ها هستند."),
    ("ریشهٔ واژه", "«یوز» از بن مضارع «یوزیدن» به معنای جَستن؛ «چیتا» از واژهٔ سانسکریت چیتراکا (خال‌خالی)."),
]
list_cards(s, features, lambda i: TEAL)

notes(s, "ویژگی‌های کلیدی یوزپلنگ آسیایی. کارت‌ها پس‌زمینهٔ رنگی پشت متن دارند و نوار رنگی "
         "در سمت راست (شروع RTL) قرار گرفته. اعداد و واحدها داخل متن فارسی run جداگانه دارند.")

# ----------------------------------------------------------------------------
# Slide 5 — Population bar chart (fonts forced to Vazirmatn)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "برآورد جمعیت: فروپاشی و بازگشتِ امید",
       kicker="نمودار ستونی · تعداد قلاده در طول زمان", num=5, total=9)

chart_data = CategoryChartData()
chart_data.categories = ["دههٔ ۱۳۵۰", "۱۳۷۶", "۱۳۸۷", "۱۳۹۹", "۱۴۰۲", "۱۴۰۴"]
chart_data.add_series("تعداد قلاده", (200, 75, 85, 15, 12, 26))

gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                        MARGIN, Inches(1.8), Inches(8.2), Inches(4.7), chart_data)
chart = gf.chart
set_chart_font(chart)          # apply Vazirmatn to every chart text element
chart.has_legend = False
plot = chart.plots[0]
plot.gap_width = 90
series = plot.series[0]
series.format.fill.solid()
series.format.fill.fore_color.rgb = TEAL
series.points[5].format.fill.solid()          # last bar = the return of hope
series.points[5].format.fill.fore_color.rgb = ACCENT
chart.has_title = True
chart.chart_title.text_frame.text = "برآورد تعداد یوزپلنگ ایرانی (قلاده)"
# ensure title font is Vazirmatn too (walker above covers it)

# annotation on the right, inside a soft card
add_rect(s, Inches(9.1), Inches(2.0), Inches(3.6), Inches(3.6), GREEN_SOFT,
         line=CARDBORDER, round_=True)
_, tf = add_textbox(s, Inches(9.35), Inches(2.3), Inches(3.1), Inches(3.1))
add_para(tf, "پیام نمودار", size=15, bold=True, color=TEAL, first=True, space_after=8)
add_para(tf, "جمعیت از حدود ۲۰۰ قلاده در دههٔ ۱۳۵۰ به ۱۲ قلاده در ۱۴۰۲ سقوط کرد؛ "
             "با پروژه‌های حفاظتی در ۱۴۰۴ به ۲۶ قلاده رسید.",
         size=14, color=INK, line=1.25, space_after=12)
add_para(tf, "مقادیر برآوردی‌اند و به‌دلیل استتار قوی یوز ممکن است واقعیت بیش از این باشد.",
         size=12, color=FAINT, line=1.2)

notes(s, "نمودار ستونی سقوط و بازگشت جمعیت یوزپلنگ ایرانی. قلم همهٔ اجزای نمودار (عنوان، "
         "محورها، برچسب‌ها) صریحاً روی وزیرمتن تنظیم شده، چون نمودار قلم تم خود را به‌ارث "
         "می‌برد نه قلم ارائه. آخرین ستون (۱۴۰۴) نارنجی است تا بازگشت امید را نشان دهد.")

# ----------------------------------------------------------------------------
# Slide 6 — Habitat & range
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "زیستگاه و پراکندگی", kicker="از آسیای غربی تا ایرانِ مرکزی", num=6, total=9)

_, tf = add_textbox(s, MARGIN, Inches(1.65), Inches(7.3), Inches(4.9))
add_para(tf, "روزی روزگاری", size=15, bold=True, color=TEAL, first=True, space_after=6)
add_para(tf, "از شبه‌جزیرهٔ عربستان و خاور نزدیک تا دریای خزر، قفقاز جنوبی، قزلقوم و هند.",
         size=15, color=INK, line=1.2, space_after=18)
add_para(tf, "امروز تنها در ایران", size=15, bold=True, color=TEAL, space_after=6)
add_para(tf, "زیستگاه‌های کویری استان‌های سمنان، یزد، خراسان (شمالی، رضوی، جنوبی) و کرمان.",
         size=15, color=INK, line=1.2, space_after=18)
add_para(tf, "مناطق حفاظت‌شدهٔ کلیدی", size=15, bold=True, color=TEAL, space_after=6)
for area in ("پارک ملی توران و میاندشت (سمنان)",
             "پناهگاه حیات وحش نایبندان و دره‌انجیر",
             "پارک ملی کویر و کوه بافق (یزد)"):
    add_para(tf, area, size=14, color=INK, space_after=5, bullet=True)

pic = s.shapes.add_picture("assets/cubs.jpg", Inches(8.25), Inches(1.8), width=Inches(4.5))
add_image_alt(pic, "توله‌های یوز", "توله‌های یوز در قفس، احتمالاً برای فروش")
add_rect(s, Inches(8.25), Inches(1.8), Inches(4.5), Inches(3.0), None, line=ACCENT)
_, cap = add_textbox(s, Inches(8.25), Inches(5.05), Inches(4.5), Inches(0.8))
add_para(cap, "توله‌های یوز — نمادِ شکنندگی نسل", size=12, color=FAINT, first=True,
         align=PP_ALIGN.CENTER)

notes(s, "زیستگاه: زمانی گسترهٔ وسیعی از آسیای غربی و مرکزی تا هند را در بر می‌گرفت؛ امروز "
         "فقط به بیابان‌های مرکزی ایران محدود شده. تصویر توله‌ها با متن جایگزین توصیف شده.")

# ----------------------------------------------------------------------------
# Slide 7 — Threats (RTL cards, red/amber accent on the right)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "تهدیدها", kicker="چرا نسل یوز رو به نابودی است؟", num=7, total=9)

threats = [
    ("تغییر اقلیم و خشکسالی", "تأثیرگذارترین عامل بر زیستگاه یوزپلنگ"),
    ("تکه‌تکه شدن زیستگاه", "جاده‌ها، معادن و توسعهٔ مسکونی و کشاورزی"),
    ("کاهش طعمه", "شکار بی‌رویهٔ آهو، جبیر، قوچ و میش و خرگوش"),
    ("ورود دام و تعارض با انسان", "اشغال آبشخورها و حملهٔ سگ گله"),
    ("جاده و معدن", "ساخت جاده در قلب مناطق حفاظت‌شده (مانند کوه بافق)"),
]
list_cards(s, threats, lambda i: WARNING if i % 2 == 0 else NEGATIVE)

notes(s, "تهدیدهای اصلی: تغییر اقلیم، از دست دادن زیستگاه، کاهش طعمه، تعارض با دامداران "
         "و جاده‌سازی در مناطق حفاظت‌شده.")

# ----------------------------------------------------------------------------
# Slide 8 — Conservation efforts (RTL cards, green accent on the right)
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "تلاش‌ها برای حفاظت", kicker="نشان‌ها و اقدام‌های کلیدی", num=8, total=9)

efforts = [
    ("پروژهٔ حفاظت از ۲۰۰۱", "با همکاری برنامهٔ عمران ملل متحد (UNDP) و سازمان حفاظت محیط زیست"),
    ("تکثیر در اسارت (۱۴۰۱)", "زایش نخستین توله‌ها در اسارت در مرکز توران"),
    ("بیمهٔ محیط‌بانان و یوز", "پوشش بیمهٔ تکمیلی رایگان از ۱۳۹۲ به‌مدت پنج سال"),
    ("گردنبند ردیاب (۲۰۰۷)", "پروژهٔ رادیو تله‌متری با دبلیو.سی.اس برای ردیابی یوز"),
    ("نماد ملی", "نقش یوز بر لباس تیم ملی فوتبال، موتور جستجوی «یوز»، هواپیماهای معراج"),
]
list_cards(s, efforts, lambda i: POSITIVE)

notes(s, "اقدام‌های حفاظتی شامل پروژهٔ ملی حفاظت، تکثیر در اسارت، بیمه، ردیابی ماهواره‌ای "
         "و به‌کارگیری یوز به‌عنوان نماد ملی برای جلب توجه عمومی است.")

# ----------------------------------------------------------------------------
# Slide 9 — Sources & metadata
# ----------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
slide_bg(s)
header(s, "منابع و دربارهٔ این ارائه", kicker="صداقت در ارجاع و ابرداده", num=9, total=9)

add_rect(s, MARGIN, Inches(1.55), SLIDE_W - 2 * MARGIN, Inches(2.7), CARD,
         line=CARDBORDER, round_=True)
_, tf = add_textbox(s, MARGIN + Inches(0.5), Inches(1.85), SLIDE_W - 2 * MARGIN - Inches(1.0),
                    Inches(2.2))
add_para(tf, "منبع اصلی", size=16, bold=True, color=TEAL, first=True, space_after=6)
p = add_para(tf, [("ویکی‌پدیا — یوزپلنگ آسیایی", {})], size=17, color=INK, space_after=2)
r = p.runs[0]
r.hyperlink.address = "https://fa.wikipedia.org/wiki/یوزپلنگ_آسیایی"
r.font.color.rgb = ACCENT
r.font.underline = True
add_para(tf, "https://fa.wikipedia.org/wiki/یوزپلنگ_آسیایی",
         size=11, color=FAINT, space_after=18)
add_para(tf, "ساخته‌شده با مهارت powerpoint-fa", size=16, bold=True, color=TEAL, space_after=6)
add_para(tf, "متن راست‌به‌چپ (rtl)، فونت وزیرمتن، ارقام فارسی، زبان فارسی (FARSI)، "
             "جدول با جهت راست‌به‌چپ، نمودار ستونی، تصویر با متن جایگزین، "
             "پیوند، پانویسِ سخنران و ابرداده.",
         size=14, color=SECONDARY, line=1.25)

badges = ["RTL", "جدول", "نمودار", "تصویر", "پانویس", "فراپیوند", "ابرداده"]
bw = Inches(1.5)
bx = SLIDE_W - MARGIN - Inches(0.3)
by = Inches(4.6)
for i, b in enumerate(badges):
    x = bx - (i + 1) * (bw + Inches(0.15))
    add_rect(s, x, by, bw, Inches(0.55), INK, round_=True)
    _, btf = add_textbox(s, x, by + Inches(0.05), bw, Inches(0.45))
    add_para(btf, b, size=13, bold=True, color=SAND, align=PP_ALIGN.CENTER, first=True)

notes(s, "این ارائه نمونه‌ای از خروجی مهارت powerpoint-fa است. همهٔ محتوا از مقالهٔ "
         "ویکی‌پدیای فارسی «یوزپلنگ آسیایی» (مجوز Creative Commons) اقتباس و بازنویسی شده.")

prs.save("yuzpalang-showcase.pptx")
print("saved yuzpalang-showcase.pptx with", len(prs.slides._sldIdLst), "slides")
