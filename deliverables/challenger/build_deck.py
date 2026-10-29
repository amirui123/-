# -*- coding: utf-8 -*-
"""Builds: Dodge Challenger — Persian (fa-IR) 16:9 deck.

Content is sourced from Dodge/Stellantis press material, Edmunds, Road & Track,
GoodCarBadCar (see closing slide). Layout is measured with real Vazirmatn metrics.
"""
import os
import sys

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import (BG, C, CW, DIM, HAIR, H, INK, M, MUTED, ORANGE, PANEL, PANEL2,
                     RED, SILVER, Deck, Run, W, fa_num)

HERE = os.path.dirname(os.path.abspath(__file__))
A = lambda f: os.path.join(HERE, "assets", f)
R = PP_ALIGN.RIGHT
CENTER = PP_ALIGN.CENTER
LEFT = PP_ALIGN.LEFT

d = Deck()
PAGE = [0]


def pnum(show=True):
    PAGE[0] += 1
    return PAGE[0] if show else None


def T(text, size, bold=False, color=INK, black=False, light=False):
    return Run(text, size, bold, black, light, color)


# ---------------------------------------------------------------------------
# 01 — COVER
# ---------------------------------------------------------------------------
s = d.slide()
pnum(False)
d.pic(s, A("cover-hellcat.jpg"), 0, 0, W, H)
d.grad(s, 0, 0, W, H, [(0.0, "050607", 30), (0.45, "050607", 55), (1.0, "050607", 88)])
d.rect(s, 0, 0, 0.16, 6.94, fill=RED)

d.text(s, 5.6, 3.24, W - M - 5.6, 0.3,
       [[T("DODGE  CHALLENGER    1970 — 2023", 12.5, True, SILVER)]],
       align=R, line_spacing=1.1, label="cover-kicker")
d.text(s, 5.2, 3.55, W - M - 5.2, 1.3, [[T("چلنجر", 76, True, INK, black=True)]],
       align=R, line_spacing=1.06, label="cover-h1")
d.text(s, 5.2, 4.92, W - M - 5.2, 0.62, [[T("پنجاه‌وسه سال وسوسهٔ آمریکایی", 30, True, INK)]],
       align=R, line_spacing=1.2, label="cover-h2")
d.hline(s, W - M - 1.7, 5.68, 1.7, RED, 0.045)
d.text(s, 5.2, 5.88, W - M - 5.2, 0.36,
       [[T("از پونی‌کار ۱۹۷۰ تا آخرین عضله‌کار تمام‌بنزینی دوج", 16.5, False, MUTED)]],
       align=R, line_spacing=1.2, label="cover-sub")
d.text(s, M, 6.5, 4.2, 0.3, [[T("آبان ۱۴۰۵  ·  ۱۳ اسلاید", 12, False, DIM)]],
       align=LEFT, line_spacing=1.1, label="cover-date")
d.text(s, 8.0, 7.06, W - M - 8.0, 0.26,
       [[T("تصویر: آرشیو عمومی", 9.5, False, DIM)]], align=R, line_spacing=1.0)

# ---------------------------------------------------------------------------
# 02 — THREE LIVES
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "تاریخچه", RED)
d.title(s, "چلنجر سه بار متولد شد و هر بار نقش تازه‌ای گرفت", 36)
d.hline(s, M, 3.32, CW, HAIR, 0.02)

cols = [
    ("۱۹۷۰ — ۱۹۷۴", "پونی‌کار",
     "E-body مشترک با باراکودا؛ ۷۶٬۹۳۵ دستگاه در سال اول و در کل ۱۶۵٬۴۳۷ دستگاه.", INK, True),
    ("۱۹۷۸ — ۱۹۸۳", "کوپهٔ کوچک",
     "نسخهٔ بازنشان‌خوردهٔ میتسوبیشی گالانت لامبدا؛ نام چلنجر ۲۵ سال خاموش ماند.", DIM, False),
    ("۲۰۰۸ — ۲۰۲۳", "بازگشت عضله‌کار",
     "پلتفرم LX/LC؛ از ۳۰۳ تا ۱٬۰۲۵ اسب بخار؛ ۴۴٬۹۶۱ دستگاه در آخرین سال.", INK, True),
]
colw, gap = 3.62, 0.40
for i, (years, name, body, inkcol, filled) in enumerate(cols):
    x = W - M - colw - i * (colw + gap)
    d.text(s, x, 2.42, colw, 0.5, [[T(years, 29, True, inkcol if filled else DIM)]],
           align=R, line_spacing=1.1, label=f"y{i}")
    d.rect(s, x + colw - 0.17, 3.245, 0.17, 0.17, shape=MSO_SHAPE.OVAL,
           fill=RED if filled else C("3A4048"))
    d.text(s, x, 3.62, colw, 0.45, [[T(name, 20.5, True, INK)]], align=R, line_spacing=1.2,
           label=f"n{i}")
    d.text(s, x, 4.14, colw, 1.5, [[T(body, 15.5, False, MUTED)]], align=R, line_spacing=1.55,
           label=f"b{i}")

d.text(s, M, 6.05, CW, 0.4,
       [[T("هر سه نسل با یک ایده ساخته شدند: بدنهٔ بزرگ، موتور جلو و دیفرانسیل عقب.", 16, False, SILVER)]],
       align=R, line_spacing=1.3, label="timeline-foot")
d.footer(s, None, True, PAGE[0])

# ---------------------------------------------------------------------------
# 03 — 1970 HERITAGE (image left, text right)
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.pic(s, A("heritage-1970.jpg"), 0, 0, 6.05, H)
d.grad(s, 0, 0, W, H, [(0.0, "050607", 4), (1.0, "050607", 40)])

tx, tw = 6.55, W - M - 6.55
d.kicker(s, "میراث", RED)
d.text(s, tx, 0.95, tw, 1.9, [[T("E-body سال ۱۹۷۰ خطی ساخت که نیم‌قرن بعد هم تکرار می‌شود", 32)]],
       align=R, line_spacing=1.28, label="h1")
facts = [
    ("۴۲۶ HEMI", "بزرگ‌ترین موتور کرایسلر با ۴۲۵ اسب بخار اعلام‌شده؛ کم‌تیراژترین و گران‌ترین گزینه."),
    ("۳۴۰ Six Pack", "نسخهٔ T/A با سه کاربراتور، برای هومولوگیشن مسابقات ترنس‌ام ساخته شد."),
    ("ونیشینگ پوینت ۱۹۷۱", "یک چلنجر R/T سفید در مشهورترین تعقیب سینمای جاده‌ای، نماد نسل شد."),
]
y = 3.15
for i, (lead, body) in enumerate(facts):
    d.text(s, tx, y, tw, 0.34, [[T(lead, 18, True, INK)]], align=R, line_spacing=1.15,
           label=f"fl{i}")
    d.text(s, tx, y + 0.37, tw, 0.62, [[T(body, 15, False, MUTED)]], align=R, line_spacing=1.5,
           label=f"fb{i}")
    if i < 2:
        d.hline(s, tx, y + 1.06, tw, HAIR, 0.01)
    y += 1.2
d.footer(s, "تصویر: آرشیو عمومی (Curbside Classic) · ارقام: Dodge، ویکی‌پدیا", True, PAGE[0], src_x=6.55, src_w=5.93)

# ---------------------------------------------------------------------------
# 04 — DESIGN DNA
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "طراحی", RED)
d.title(s, "چهار نشانهٔ بصری، چلنجر را از هر کوپهٔ دیگر جدا می‌کند", 36)
cards = [
    ("۰۱", "جلوپنجرهٔ تمام‌عرض", "چهار چراغ گرد در یک قاب افقی که عرض کامل دماغه را می‌گیرد."),
    ("۰۲", "شانهٔ عقب سنگین", "خط سقف کم‌شیب و گلگیر عقب حجیم؛ میراث مستقیم بدنهٔ E-body."),
    ("۰۳", "بدنهٔ Widebody", "از ۲۰۱۸ گلگیرهای پهن‌تر با رینگ ۲۰ اینچی و لاستیک ۳۰۵."),
    ("۰۴", "نشان و صدا", "لوگوی Hellcat و نالهٔ سوپرشارژر، امضای نسخه‌های SRT شد."),
]
cw, ch = 5.62, 1.86
for i, (num, head, body) in enumerate(cards):
    row, col = divmod(i, 2)
    x = (W - M - cw) if col == 0 else M
    y = 2.42 + row * (ch + 0.26)
    d.rect(s, x, y, cw, ch, fill=PANEL)
    d.rect(s, x + cw - 0.055, y + 0.3, 0.055, 0.52, fill=RED)
    d.text(s, x + cw - 0.78, y + 0.25, 0.5, 0.26, [[T(num, 12, True, RED)]], align=R,
           line_spacing=1.0, label=f"cn{i}")
    d.text(s, x + 0.3, y + 0.5, cw - 1.45, 0.4, [[T(head, 20.5, True, INK)]], align=R,
           line_spacing=1.15, label=f"ch{i}")
    d.text(s, x + 0.3, y + 1.0, cw - 0.7, 0.72, [[T(body, 14.5, False, MUTED)]], align=R,
           line_spacing=1.5, label=f"cb{i}")
d.footer(s, None, True, PAGE[0])

# ---------------------------------------------------------------------------
# 05 — POWER LADDER
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "نردبان قدرت", RED)
d.title(s, "یک بدنه، هفت مرحلهٔ قدرت: از ۳۰۳ تا ۱٬۰۲۵ اسب بخار", 36)

c1, w1 = 8.42, 4.06
c2, w2 = 5.40, 2.90
c3, w3 = 3.72, 1.54
c4, w4 = 0.85, 2.74
heads = [("تیپ", c1, w1), ("پیشرانه", c2, w2), ("قدرت (hp)", c3, w3), ("یادداشت", c4, w4)]
for h, x, wid in heads:
    d.text(s, x, 2.24, wid, 0.3, [[T(h, 15, True, MUTED)]], align=R, line_spacing=1.0,
           label=f"hh-{h}")
d.hline(s, M, 2.62, CW, C("7A828E"), 0.016)

rows = [
    ("SXT / GT", "۳.۶ لیتری V6", "۳۰۳", "تنها نسخه با گزینهٔ AWD", INK),
    ("R/T", "۵.۷ لیتری HEMI", "۳۷۵", "دستی ۶ سرعته یا خودکار", INK),
    ("R/T Scat Pack", "۶.۴ لیتری ۳۹۲", "۴۸۵", "محبوب‌ترین نسخهٔ V8", INK),
    ("SRT Hellcat", "۶.۲ سوپرشارژ", "۷۱۷", "آغاز عصر Hellcat", INK),
    ("Hellcat Redeye", "۶.۲ سوپرشارژ", "۷۹۷", "سوپرشارژر بزرگ‌تر", INK),
    ("Super Stock", "۶.۲ سوپرشارژ", "۸۰۷", "درگ‌محور اما مجاز برای جاده", INK),
    ("Demon 170", "۶.۲ + E85", "۱٬۰۲۵", "فقط ۳٬۳۰۰ دستگاه، ۲۰۲۳", RED),
]
ry = 2.74
rh = 0.575
for i, (name, eng, hp, note, col) in enumerate(rows):
    d.text(s, c1, ry + 0.1, w1, 0.36, [[T(name, 17, True, INK)]], align=R, line_spacing=1.1,
           label=f"rn{i}")
    d.text(s, c2, ry + 0.13, w2, 0.32, [[T(eng, 14.5, False, SILVER)]], align=R,
           line_spacing=1.1, label=f"re{i}")
    d.text(s, c3, ry + 0.02, w3, 0.44, [[T(hp, 20, True, col)]], align=R, line_spacing=1.1,
           label=f"rh{i}")
    d.text(s, c4, ry + 0.14, w4, 0.32, [[T(note, 13, False, MUTED)]], align=R,
           line_spacing=1.1, label=f"rnt{i}")
    if i < len(rows) - 1:
        d.hline(s, M, ry + rh - 0.04, CW, HAIR, 0.009)
    ry += rh
d.footer(s, "ارقام: مشخصات رسمی دوج برای مدل ۲۰۲۳", True, PAGE[0])

# ---------------------------------------------------------------------------
# 06 — DEMON 170 HERO
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.pic(s, A("demon-burnout.jpg"), 0, 0, 5.75, H)
d.grad(s, 0, 0, W, H, [(0.0, "050607", 4), (1.0, "050607", 40)])

tx, tw = 6.2, W - M - 6.2
d.kicker(s, "دمون ۱۷۰", RED)
d.text(s, tx, 0.92, tw, 1.4, [[T("دمون ۱۷۰ سریع‌ترین ماشینی است که دوج ساخته", 32)]],
       align=R, line_spacing=1.26, label="h1")
stats = [("۱٬۰۲۵", "اسب بخار با E85"), ("۹۴۵", "پوند-فوت گشتاور"),
         ("۱.۶۶", "ثانیه تا ۶۰ مایل"), ("۸.۹۱", "ثانیه یک‌چهارم مایل")]
for i, (val, lab) in enumerate(stats):
    row, col = divmod(i, 2)
    x = (W - M - 3.0) if col == 0 else tx
    y = 2.5 + row * 1.42
    d.text(s, x, y, 3.0, 0.72, [[T(val, 41, True, INK, black=True)]], align=R,
           line_spacing=1.0, label=f"sv{i}")
    d.text(s, x, y + 0.76, 3.0, 0.3, [[T(lab, 13.5, False, MUTED)]], align=R,
           line_spacing=1.0, label=f"sl{i}")
d.text(s, tx, 5.45, tw, 1.25,
       [[T("۸.۹۱ ثانیه با سرعت ۱۵۱.۱۷ مایل بر ساعت و تأیید NHRA ثبت شد. قیمت پایه ۹۶٬۶۶۶ دلار و تیراژ ۳٬۳۰۰ دستگاه (۳٬۰۰۰ آمریکا + ۳۰۰ کانادا).", 14.5, False, MUTED)]],
       align=R, line_spacing=1.55, label="note")
d.footer(s, None, True, PAGE[0])

# ---------------------------------------------------------------------------
# 07 — FUEL / POWER CHART
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "سوخت", RED)
d.title(s, "قدرت دمون ۱۷۰ را اکتان تعیین می‌کند، نه موتور", 36)

base_y = 5.62
d.text(s, 7.9, 2.3, 4.58, 0.28, [[T("قدرت اعلام‌شده، اسب بخار", 12.5, False, MUTED)]],
       align=R, line_spacing=1.0, label="axis-title")
bars = [("بنزین ۹۱ اکتان", 880, C("454C56")), ("E10", 900, C("69737E")), ("E85", 1025, RED)]
bw, bgap = 1.9, 0.72
for i, (lab, val, col) in enumerate(bars):
    x = W - M - bw - i * (bw + bgap)
    h = val / 1025 * 2.55
    d.rect(s, x, base_y - h, bw, h, fill=col)
    d.text(s, x, base_y - h - 0.5, bw, 0.44, [[T(fa_num(f"{val:,}".replace(",", "٬")), 26, True, INK)]],
           align=CENTER, line_spacing=1.0, label=f"bv{i}")
    d.text(s, x - 0.25, base_y + 0.14, bw + 0.5, 0.34, [[T(lab, 14.5, False, MUTED)]],
           align=CENTER, line_spacing=1.0, label=f"bl{i}")
d.hline(s, W - M - (3 * bw + 2 * bgap), base_y, 3 * bw + 2 * bgap, C("5A636E"), 0.016)

nx, nw = M, 3.9
d.text(s, nx, 2.34, nw, 0.34, [[T("چرا عدد ثابت نیست", 17.5, True, INK)]], align=R,
       line_spacing=1.1, label="note-h")
d.text(s, nx, 2.78, nw, 1.9,
       [[T("حسگر درصد اتانول، نقشهٔ موتور را تغییر می‌دهد: با بنزین ۹۱ اکتان ۸۸۰ اسب، با E10 نود اسب بیشتر و با E85 به ۱٬۰۲۵ اسب می‌رسد.", 14, False, MUTED)]],
       align=R, line_spacing=1.55, label="note-b")
d.text(s, nx, 4.75, nw, 0.9,
       [[T("گشتاور هم پابه‌پای سوخت بالا می‌رود: ۸۱۰ پوند-فوت با E10 و ۹۴۵ با E85.", 13, False, DIM)]],
       align=R, line_spacing=1.5, label="note-b2")
d.footer(s, "منبع: بیانیهٔ مطبوعاتی Dodge، مارس ۲۰۲۳", True, PAGE[0])

# ---------------------------------------------------------------------------
# 08 — REALITY CHECK
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "واقعیت", RED)
d.title(s, "۱.۶۶ ثانیه در خیابان تکرار نمی‌شود؛ عدد آزمایشگاهی است", 36)

items = [
    ("۰۱", "سطح آماده", "آسفالت چسبندهٔ پیست درگ، نه آسفالت شهری یا لاستیک سرد."),
    ("۰۲", "لاستیک ویژه", "درگ رادیال Mickey Thompson که اختصاصی این مدل ساخته شد."),
    ("۰۳", "گیربکس و ترمز", "۸ سرعتهٔ TorqueFlite با TransBrake 2.0 و لانچ کنترل."),
    ("۰۴", "مجوز", "NHRA اجرای زیر ۹ ثانیه را بدون رول‌کیج و چتر ممنوع اعلام کرد."),
]
ix, iw = 4.85, 7.63
y = 2.36
for num, head, body in items:
    d.text(s, ix + iw - 0.62, y + 0.02, 0.62, 0.3, [[T(num, 13, True, RED)]], align=R,
           line_spacing=1.0, label="rn")
    d.text(s, ix, y, iw - 0.8, 0.34, [[T(head, 18, True, INK)]], align=R, line_spacing=1.1,
           label="rh")
    d.text(s, ix, y + 0.4, iw - 0.8, 0.5, [[T(body, 14.5, False, MUTED)]], align=R,
           line_spacing=1.45, label="rb")
    if num != "۰۴":
        d.hline(s, ix, y + 0.96, iw, HAIR, 0.009)
    y += 1.13

d.rect(s, M, 2.36, 3.55, 2.62, fill=PANEL)
d.text(s, M + 0.25, 2.58, 3.05, 0.82, [[T("۲.۰۰۴g", 43, True, RED, black=True)]], align=R,
       line_spacing=1.0, label="g")
d.text(s, M + 0.25, 3.42, 3.05, 0.9,
       [[T("بیشترین شتاب G ثبت‌شده برای یک خودروی تولیدی", 14, False, SILVER)]],
       align=R, line_spacing=1.45, label="g-l")
d.text(s, M + 0.25, 4.38, 3.05, 0.5, [[T("۶۰ فوت در ۱.۲۴ ثانیه", 13, False, DIM)]], align=R,
       line_spacing=1.3, label="g-60")
d.footer(s, "تصویر: آرشیو عمومی (HotCars) · منبع: Dodge، مارس ۲۰۲۳", True, PAGE[0], src_x=6.2, src_w=6.28)

# ---------------------------------------------------------------------------
# 09 — RIVALS (rivals win an axis)
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "رقیب", RED)
d.title(s, "قدرت را دوج می‌برد، وزن و چابکی را رقیب", 36)

c1, w1 = 7.55, 4.93
c2, w2 = 5.35, 2.05
c3, w3 = 3.25, 1.95
c4, w4 = 0.85, 2.25
for h, x, wid in [("خودرو", c1, w1), ("قدرت", c2, w2), ("وزن", c3, w3), ("جایگاه", c4, w4)]:
    d.text(s, x, 2.3, wid, 0.3, [[T(h, 15, True, MUTED)]], align=R, line_spacing=1.0, label="rh")
d.hline(s, M, 2.68, CW, C("7A828E"), 0.016)

rivals = [
    ("دوج چلنجر Demon 170", "۱٬۰۲۵", "۴٬۲۵۵ پوند", "قدرت خط مستقیم", INK, INK),
    ("فورد موستانگ Shelby GT500", "۷۶۰", "۴٬۱۸۲ پوند", "بهترین تعادل", INK, INK),
    ("شورولت کامارو ZL1 1LE", "۶۵۰", "۳٬۸۸۳ پوند", "سبک‌ترین", INK, RED),
]
ry, rh = 2.82, 0.86
for i, (name, hp, wt, pos, hpcol, wtcol) in enumerate(rivals):
    d.text(s, c1, ry + 0.12, w1, 0.42, [[T(name, 18.5, True, INK)]], align=R, line_spacing=1.1,
           label=f"rv{i}")
    d.text(s, c2, ry + 0.1, w2, 0.44, [[T(hp, 22, True, hpcol)]], align=R, line_spacing=1.1,
           label=f"rvh{i}")
    d.text(s, c3, ry + 0.14, w3, 0.34, [[T(wt, 16, True, wtcol)]], align=R, line_spacing=1.1,
           label=f"rvw{i}")
    d.text(s, c4, ry + 0.16, w4, 0.34, [[T(pos, 13.5, False, MUTED)]], align=R,
           line_spacing=1.1, label=f"rvp{i}")
    if i < 2:
        d.hline(s, M, ry + rh - 0.06, CW, HAIR, 0.009)
    ry += rh

d.rect(s, M, 5.62, CW, 0.98, fill=PANEL)
d.text(s, M + 0.4, 5.62, CW - 0.8, 0.98,
       [[T("چلنجر در وزن، تعادل پیست و دید راننده می‌بازد؛ در قدرت، فضای کابین و صدای موتور می‌برد.", 15.5, False, SILVER)]],
       align=R, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.45, label="verdict")
d.footer(s, "ارقام: مشخصات رسمی سازندگان؛ وزن‌ها به پوند", True, PAGE[0])

# ---------------------------------------------------------------------------
# 10 — SALES CHART
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "فروش", RED)
d.title(s, "چلنجر در ۲۰۲۱ و ۲۰۲۲ پرفروش‌ترین عضله‌کار آمریکا بود", 36)

base_y = 5.5
d.text(s, M, 2.34, 6.4, 0.3, [[T("فروش سالانه در آمریکا، دستگاه", 12.5, False, MUTED)]],
       align=R, line_spacing=1.0, label="axis-title")
sales = [("۲۰۱۹", 60997), ("۲۰۲۰", 52955), ("۲۰۲۱", 54315), ("۲۰۲۲", 55245), ("۲۰۲۳", 44961)]
bw, bgap = 1.22, 0.52
total = len(sales) * bw + (len(sales) - 1) * bgap
x0 = W - M - total
for i, (yr, val) in enumerate(sales):
    x = x0 + i * (bw + bgap)
    h = val / 60997 * 2.42
    col = RED if yr == "۲۰۲۳" else C("3E454F")
    d.rect(s, x, base_y - h, bw, h, fill=col)
    d.text(s, x - 0.15, base_y - h - 0.44, bw + 0.3, 0.34,
           [[T(fa_num(f"{val:,}".replace(",", "٬")), 14.5, True, INK)]], align=CENTER,
           line_spacing=1.0, label=f"sv{i}")
    d.text(s, x - 0.15, base_y + 0.12, bw + 0.3, 0.3, [[T(yr, 14, False, MUTED)]],
           align=CENTER, line_spacing=1.0, label=f"sy{i}")
d.hline(s, x0, base_y, total, C("5A636E"), 0.016)

d.text(s, M, 5.95, 7.6, 0.75,
       [[T("در ۲۰۲۲ موستانگ ۴۷٬۵۶۶ و کامارو ۲۴٬۶۵۲ دستگاه فروختند؛ سال ۲۰۲۳ سال خداحافظی بود و آمارش افت کرد.", 13.5, False, MUTED)]],
       align=R, line_spacing=1.5, label="cap")
d.footer(s, "منبع: GoodCarBadCar، آمار فروش سالانهٔ آمریکا", True, PAGE[0])

# ---------------------------------------------------------------------------
# 11 — END OF AN ERA
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.pic(s, A("demon-170-front.jpg"), 0, 0, 4.95, H)
d.grad(s, 0, 0, W, H, [(0.0, "050607", 4), (1.0, "050607", 40)])
d.text(s, 0.3, 0.3, 4.4, 0.28, [[T("۲۰۲۳ Challenger SRT Demon 170", 9.5, False, C("E4E7EB"))]],
       align=LEFT, line_spacing=1.0)

tx, tw = 5.4, W - M - 5.4
d.kicker(s, "پایان دوره", RED)
d.text(s, tx, 0.95, tw, 1.45, [[T("دوج دسامبر ۲۰۲۳ خط تولید را بست و برقی شد", 31)]],
       align=R, line_spacing=1.28, label="h1")
steps = [
    ("مارس ۲۰۲۳", "سفارش دمون ۱۷۰ باز شد و تیراژ روی ۳٬۳۰۰ دستگاه ثابت ماند."),
    ("۲۰۲۳", "هفت نسخهٔ Last Call از Shakedown تا Demon 170 معرفی شد."),
    ("دسامبر ۲۰۲۳", "آخرین چلنجر از خط برامپتون در انتاریو بیرون آمد و خط بسته شد."),
    ("۲۰۲۴", "جانشین برقی: چارجر دایتونا با ۴۹۶ تا ۶۷۰ اسب بخار."),
]
y = 2.6
for i, (when, what) in enumerate(steps):
    d.text(s, tx, y, tw, 0.34, [[T(when, 17, True, RED)]], align=R, line_spacing=1.1,
           label=f"sw{i}")
    d.text(s, tx, y + 0.38, tw, 0.6, [[T(what, 14.5, False, MUTED)]], align=R,
           line_spacing=1.45, label=f"sd{i}")
    if i < 3:
        d.hline(s, tx, y + 0.98, tw, HAIR, 0.009)
    y += 1.12
d.footer(s, "تصویر: آرشیو عمومی · منابع: Stellantis، دسامبر ۲۰۲۳", True, PAGE[0], src_x=5.4, src_w=7.08)

# ---------------------------------------------------------------------------
# 12 — BUYING GUIDE
# ---------------------------------------------------------------------------
s = d.slide()
pnum()
d.kicker(s, "راهنمای خرید", RED)
d.title(s, "چلنجر کارکرده را بدون این پنج بررسی نخر", 36)

checks = [
    ("۰۱", "سوپرشارژر", "نشتی روغن، صدای یاتاقان و تاریخ تعویض کولر را با متخصص بررسی کن."),
    ("۰۲", "تاریخچهٔ درگ", "رد لاستیک نیمه‌سلیک، نیم‌محور تقویت‌شده و تیون ECU را ببین."),
    ("۰۳", "اصالت تیپ", "VIN و کد موتور را با برگهٔ ساخت تطبیق بده؛ Widebody و Jailbreak بازار جدا دارند."),
    ("۰۴", "سرویس", "دفترچهٔ سرویس روغن و شمع؛ نسخه‌های ۷۱۷ اسبی به نگهداری دقیق حساس‌اند."),
    ("۰۵", "بدنه", "گلگیر Widebody یک‌تکه است؛ تعویض بی‌دقت درز و رنگ را لو می‌دهد."),
]
y = 2.38
for i, (num, head, body) in enumerate(checks):
    d.text(s, W - M - 0.72, y + 0.06, 0.72, 0.34, [[T(num, 13.5, True, RED)]], align=R,
           line_spacing=1.0, label=f"bn{i}")
    d.text(s, M, y, CW - 0.85, 0.34, [[T(head, 18.5, True, INK)]], align=R, line_spacing=1.1,
           label=f"bh{i}")
    d.text(s, M, y + 0.4, CW - 0.85, 0.5, [[T(body, 14.5, False, MUTED)]], align=R,
           line_spacing=1.45, label=f"bb{i}")
    if i < 4:
        d.hline(s, M, y + 1.0, CW, HAIR, 0.009)
    y += 0.92
d.footer(s, "این فهرست عمومی است و جای بازرسی فنی تخصصی را نمی‌گیرد.", True, PAGE[0])

# ---------------------------------------------------------------------------
# 13 — CLOSING
# ---------------------------------------------------------------------------
s = d.slide()
pnum(False)
d.pic(s, A("cover-hellcat.jpg"), 0, 0, W, H)
d.grad(s, 0, 0, W, H, [(0.0, "050607", 70), (1.0, "050607", 92)])
d.rect(s, 0, 0, 0.16, 6.94, fill=RED)
d.text(s, M, 1.0, CW, 1.8, [[T("چلنجر آخرین نفس موتور بنزینی آمریکا بود", 46, True, INK, black=True)]],
       align=R, line_spacing=1.28, label="close-h")
d.hline(s, W - M - 2.1, 2.98, 2.1, RED, 0.045)
d.text(s, M, 3.22, CW, 0.85,
       [[T("دوج دسامبر ۲۰۲۳ خط برامپتون را بست تا برای نسل برقی جا باز کند؛ آنچه ماند یک صدا و یک عدد بود: ۱٬۰۲۵ اسب بخار.", 17.5, False, SILVER)]],
       align=R, line_spacing=1.55, label="close-sub")
d.text(s, M, 6.02, CW, 0.86,
       [[T("منابع: بیانیهٔ مطبوعاتی Dodge / Stellantis (مارس ۲۰۲۳)؛ Edmunds؛ Road & Track؛ GoodCarBadCar؛ ویکی‌پدیای Dodge Challenger.", 10.5, False, DIM)],
        [T("تصاویر: آرشیو عمومی (Carscoops، HotCars، Curbside Classic) — حقوق تصاویر متعلق به صاحبان اصلی است.", 10.5, False, DIM)],
        [T("فونت Vazirmatn داخل پوشهٔ fonts همراه فایل است؛ پیش از باز کردن دک آن را نصب کنید.", 10.5, False, DIM)]],
       align=R, line_spacing=1.45, space_after=2, label="credits")


NOTES = {
    1: "این ارائه یک مرور کامل بر دوج چلنجر است: از تولد ۱۹۷۰ تا آخرین نسخهٔ ۲۰۲۳ و جانشین برقی. مخاطب هدف کسی است که یا قصد خرید دارد یا می‌خواهد تاریخچه و جایگاه این خودرو را بداند.",
    2: "نکتهٔ کلیدی این اسلاید این است که چلنجر سه نسل کاملاً متفاوت با سه نقش متفاوت بوده؛ نامش دو بار خاموش و دوباره زنده شد. مکث روی تفاوت نسل دوم (کوپهٔ کوچک میتسوبیشی) معمولاً بیشترین واکنش را می‌گیرد.",
    3: "خط طراحی ۱۹۷۰ را با دماغهٔ پهن، جلوپنجرهٔ تمام‌عرض و شانهٔ عقب حجیم توصیف کنید. ۴۲۶ HEMI موتور افسانه‌ای است؛ ۳۴۰ Six Pack نسخهٔ هومولوگیشن مسابقه بود.",
    4: "چهار نشانهٔ بصری که در هر چهار نسل تکرار شد. اگر مخاطب طراحی‌محور است، روی نسبت عرض به ارتفاع و چراغ‌های گرد تأکید کنید.",
    5: "اینجا نردبان قدرت را نشان می‌دهیم؛ پیام این است که همان بدنه با موتورهای متفاوت از ۳۰۳ تا ۱٬۰۲۵ اسب بخار عرضه شد. نسخه‌های Hellcat و Redeye پرفروش‌ترین نسخه‌های پرهزینه بودند.",
    6: "دیوانگی مهندسی: ۱٬۰۲۵ اسب بخار از یک موتور ۶.۲ لیتری با E85. اگر از ارقام می‌پرسند، بگویید همهٔ این اعداد رسمی و اعلام‌شدهٔ دوج هستند.",
    7: "پیام اسلاید: عدد قدرت ثابت نیست و به سوخت بستگی دارد. حسگر اتانول درصد E85 را می‌خواند و نقشهٔ موتور را عوض می‌کند؛ پس ۸۸۰ تا ۱٬۰۲۵ اسب بسته به پمپ بنزین.",
    8: "صداقت فنی: ۱.۶۶ ثانیه روی آسفالت آمادهٔ پیست و با لاستیک درگ مخصوص به دست آمده. همین موضوع باعث شد NHRA اجرای زیر ۹ ثانیه را بدون تجهیزات ایمنی ممنوع اعلام کند.",
    9: "مقایسه با رقبا: چلنجر در قدرت اول است، اما در وزن و تعادل پیست عقب‌تر. اگر مخاطب پیست‌رو است، این اسلاید نقطهٔ تصمیم است.",
    10: "داده‌ها نشان می‌دهد چلنجر در ۲۰۲۱ و ۲۰۲۲ پرفروش‌ترین عضله‌کار آمریکا بود و در سال آخر افت کرد. منبع، آمار سالانهٔ فروش آمریکاست.",
    11: "پایان دوره: دسامبر ۲۰۲۳ خط برامپتون بسته شد. جانشین برقی است و دیگر صدای V8 ندارد؛ به همین دلیل امروز این خودرو موضوع علاقهٔ مجموعه‌دارها شده.",
    12: "راهنمای خرید: پنج نکتهٔ عملی. تأکید کنید این فهرست جای بازرسی فنی تخصصی را نمی‌گیرد.",
    13: "جمع‌بندی در یک جمله: پایان یک دوره، با یک عدد به‌یادماندنی. سپس به پرسش‌ها بروید.",
}
for idx, text in NOTES.items():
    sl = d.prs.slides[idx - 1]
    sl.notes_slide.notes_text_frame.text = text

out = os.path.join(HERE, "Dodge-Challenger-FA.pptx")
d.save(out)
import fit
d.warnings += fit.fit_presentation(out)
print("saved:", out, "slides:", len(d.prs.slides.__iter__.__self__._sldIdLst))
print("--- fit warnings:", len(d.warnings))
for w in d.warnings:
    print("  ", w)
