# فهرست راهنمای اسکیل‌های «ارائهٔ زنده» و پاورپوینت متحرک

> نسخهٔ اسنپ‌شات: **۲۰۲۶-۱۰-۲۹** · مجموع: **۱۶ اسکیل / ۶۸۶ فایل / ۱۹ مگابایت**
> جزئیات منبع، کامیت دقیق و مجوز هر پوشه در [`SOURCES.md`](./SOURCES.md) آمده است.

---

## ⚠️ هشدار مصرف توکن — مهم‌ترین بخش این فایل

هر `SKILL.md` در این مجموعه بین **۳ تا ۷۵ کیلوبایت** است و چند تا از آن‌ها
(`officecli-pptx` = ۴۴KB، `officecli-pitch-deck` = ۶۶KB، `genoffice` = ۷۵KB)
تنها در یک بار خواندن چند ده هزار توکن مصرف می‌کنند.

**قانون:** برای هر تسک، فقط **۱ اسکیل اصلی** (و در صورت نیاز **حداکثر ۱ اسکیل
مکمل**) را فعال/بارگذاری کنید. هرگز کل پوشه را یک‌جا به مدل ندهید.
اگر شک دارید کدام را بردارید، از «درخت تصمیم» پایین همین صفحه استفاده کنید.

---

## درخت تصمیم سریع

```
فایل واقعی .pptx می‌خواهی؟ (قابل ویرایش در PowerPoint)
├── بله، با انیمیشن Morph بین اسلایدها ....... morph-ppt   (یا morph-ppt-3d برای مدل سه‌بعدی)
├── بله، دک سرمایه‌گذاری / VC ............... officecli-pitch-deck
├── بله، عمومی و همه‌کاره ................... officecli-pptx
├── بله، علمی/دانشگاهی ..................... academic-pptx-skill  (+ officecli-pptx/pptagent برای ساخت فایل)
├── بله، مدیریتی «نتیجه‌اول» ................ genoffice-executive-deck  (+ genoffice برای اجرا)
└── بله، فارسی و راست‌چین ................... powerpoint-fa

اسلاید تعاملی تحت وب (کد زنده، دمو)؟
├── ارائهٔ فنی/دولوپری با Markdown + Vue ..... slidev
├── HTML زیبا و مستقل، بدون build .......... frontend-slides
├── ۸۴ استایل بصری + خروجی HTML و PPTX ..... huashu-design
└── Markdown + نمودار SVG (Marp) ........... marp-slides

ویرایش زندهٔ فایلِ باز در PowerPoint؟ ....... powerpoint-mcp (+ review-slide-visual برای بازبینی)
اسلاید HTML با رندر و تبدیل خودکار به PPTX ... pptagent
```

---

## جدول کامل اسکیل‌ها

| # | اسکیل (پوشه) | خروجی | مخزن مبدأ | لایسنس | پیش‌نیاز اجرایی | چه زمانی فراخوانی شود |
|---|---|---|---|---|---|---|
| ۱ | `officecli-pptx` | `.pptx` قابل ویرایش | iOfficeAI/OfficeCLI | Apache-2.0 | باینری `officecli` (GitHub Releases) — **اسکیل پایهٔ کل خانوادهٔ OfficeCLI** | هر زمان یک فایل `.pptx` در ورودی یا خروجی تسک هست: ساخت دک، خواندن/استخراج متن، ویرایش ارائهٔ موجود، ادغام/تفکیک اسلاید، اسپیکرنوت |
| ۲ | `morph-ppt` | `.pptx` + ترنزیشن Morph | iOfficeAI/OfficeCLI | Apache-2.0 | `officecli` + **باید `officecli-pptx` را هم خوانده باشی** (لایهٔ صحنه روی آن است) | وقتی کاربر «مورف»، «انتقال نرم»، «انیمیشن پیوسته بین اسلایدها»، «مثل Keynote» می‌خواهد؛ شکل‌هایی که با رفتن به اسلاید بعد رشد/جابه‌جا/چرخش می‌کنند |
| ۳ | `morph-ppt-3d` | `.pptx` + Morph سه‌بعدی | iOfficeAI/OfficeCLI | Apache-2.0 | `officecli` + **`morph-ppt`** (این اسکیل توسعه‌دهندهٔ `morph-ppt` است، جایگزین نیست) | وقتی علاوه بر Morph، مدل سه‌بعدی (`.glb`)، دوربین سینمایی یا چیدمان مدل+محتوا خواسته شده باشد |
| ۴ | `officecli-pitch-deck` | `.pptx` دک سرمایه‌گذاری | iOfficeAI/OfficeCLI | Apache-2.0 | `officecli` + `officecli-pptx` | فقط دک جذب سرمایه (Seed / Series A/B/C / SAFE / convertible note)؛ برای دک عمومی یا بورد ریویو از `officecli-pptx` استفاده کن |
| ۵ | `genoffice` | pptx / docx / xlsx / pdf | genspark-ai/genoffice | Apache-2.0 | CLI `genoffice` (نصب با اپ GenOffice) — **بدون نصب، اسکیل عملاً بی‌فایده است** | تبدیل فرمت (pdf↔docx↔pptx↔md↔html↔csv)، خواندن ساختار فایل آفیس موجود، ساخت دک با خط لولهٔ outline→spec→render |
| ۶ | `genoffice-executive-deck` | `.pptx` مدیریتی | genspark-ai/genoffice | Apache-2.0 | `genoffice` CLI | گزارش مدیریتی/بورد/وضعیت که «عنوان هر اسلاید = نتیجه» و ساختار هرمی می‌خواهد |
| ۷ | `genoffice-course-slides` | `.pptx` آموزشی | genspark-ai/genoffice | Apache-2.0 | `genoffice` CLI | تبدیل جزوه/طرح درس به اسلاید آموزشی با اسکریپت گفتار مدرس در notes |
| ۸ | `pptagent` | `.pptx` (از منبع HTML) | icip-cas/PPTAgent | MIT | پایتون ۳.۱۱+ و `.venv` داخل پوشهٔ اسکیل؛ `node` برای بخش HTML | وقتی می‌خواهی اسلایدها را با HTML بنویسی، رندر و بازبینی بصری کنی و PPTX قابل‌ویرایش تحویل بگیری |
| ۹ | `academic-pptx-skill` | راهنمای محتوا/ساختار | Gabberflast/academic-pptx-skill | MIT* | **بدون پیش‌نیاز** (اسکیل محتوایی است، فایل نمی‌سازد) | سخنرانی کنفرانس، سمینار، دفاع تز، ژورنال کلاب، گزارش گرنت. برای ساخت فایل، مکمل `officecli-pptx` یا `pptagent` را بگیر |
| ۱۰ | `slidev` | سایت اسلاید تحت وب | slidevjs/slidev | MIT | `pnpm` + Node (برای PDF: `playwright-chromium`) | ارائهٔ فنی/دولوپری با Markdown + کامپوننت Vue، هایلایت کد، دموی تعاملی، Monaco |
| ۱۱ | `frontend-slides` | HTML تک‌فایل، بدون وابستگی | zarazhangrui/frontend-slides | MIT | **بدون پیش‌نیاز** برای ساخت؛ برای PDF: `bash scripts/export-pdf.sh` (نیازمند Playwright) | دک HTML انیمیشن‌دار بدون build، یا تبدیل `.pptx` موجود به وب (`scripts/extract-pptx.py`) |
| ۱۲ | `huashu-design` | HTML / MP4 / GIF / PPTX | alchaincyf/huashu-design | MIT | Node + `playwright`, `pptxgenjs`, `sharp`, `pdf-lib` (فقط برای خروجی PPTX/PDF/MP4) | وقتی «طراحی خوب» و انتخاب سبک بصری از میان ۸۴ استایل مهم است؛ پروتوتایپ، اسلاید، انیمیشن، ویژوالیزیشن (**اسکیل چینی‌زبان است**) |
| ۱۳ | `marp-slides` | `.md` + PDF/HTML | robonuggets/marp-slides | **نامشخص** ⚠️ | افزونهٔ VS Code «Marp for VS Code» یا `npx @marp-team/marp-cli` | اسلاید Markdown با نمودار SVG و کامپوننت داشبورد؛ **پیش از استفاده تجاری `LICENSE-STATUS.md` را بخوان** |
| ۱۴ | `powerpoint-mcp` | تغییر زندهٔ فایل باز | kzarzycki/powerpoint-mcp | MIT | MCP server `powerpoint-mcp` + افزونهٔ Office.js + **PowerPoint روی macOS** | وقتی ارائه **همین حالا در PowerPoint باز است** و می‌خواهی زنده شکل/جدول/متن اضافه کنی یا اسکرین‌شات بگیری |
| ۱۵ | `review-slide-visual` | بازبینی بصری | kzarzycki/powerpoint-mcp | MIT | `powerpoint-mcp` (از ابزارهای `screenshot_slide`) | **فقط پس از** ویرایش اسلاید، برای کنترل کیفیت بصری؛ بازبین مستقل با کانتکست جدا spawn می‌کند |
| ۱۶ | `powerpoint-fa` | `.pptx` فارسی RTL | eledah/hermes-pptx-skill-fa | MIT (در frontmatter) | `python-pptx==1.0.2` + `lxml` + `Pillow` (پایتون ۳.۱۰+) | هر ارائهٔ فارسی: راست‌چین، نیم‌فاصله، فونت Vazirmatn، چک‌لیست کیفیت RTL (`scripts/check_deck_fa.py`) |

`*` در `academic-pptx-skill` فایل `SKILL.md` عبارت `license: Proprietary` را در
frontmatter دارد، اما فایل `LICENSE` مخزن **MIT** است. تناقض دارد؛ پیش از استفادهٔ
تجاری بررسی حقوقی کنید.

---

## ترکیب‌های پیشنهادی (حداکثر ۲ اسکیل)

| سناریوی واقعی | اسکیل‌ها |
|---|---|
| دک فارسی با انیمیشن Morph | `powerpoint-fa` + `morph-ppt` (و در صورت نیاز `officecli-pptx`) |
| دک سرمایه‌گذاری با حرکت پیوسته | `officecli-pitch-deck` + `morph-ppt` |
| ارائهٔ علمی با فایل قابل‌ویرایش | `academic-pptx-skill` + `pptagent` |
| گزارش بورد به سبک نتیجه‌اول | `genoffice-executive-deck` + `genoffice` |
| دک فنی با کد زنده | `slidev` |
| دک HTML با کیفیت طراحی بالا | `frontend-slides` (یا `huashu-design` اگر انتخاب سبک بصری مهم است) |
| اصلاح ارائه‌ای که همین حالا باز است | `powerpoint-mcp` + `review-slide-visual` |

---

## نکات امنیتی و اجرایی

1. **هیچ نصبی در این پوشه اجرا نشده است.** همهٔ فایل‌ها فقط «متن» هستند.
   اسکریپت‌های موجود در `scripts/` عمداً اجرا نشدند.
2. **بسیاری از `SKILL.md`ها دستور نصب خودکار دارند** (مثلاً `officecli-pptx` می‌گوید
   `curl -fsSL https://d.officecli.ai/install.sh | bash`). این دستور را **بی‌قید و شرط
   اجرا نکنید**؛ باینری را از GitHub Releases بگیرید و هش/امضا را بررسی کنید.
3. `powerpoint-mcp/scripts/fetch-logo.sh` و `huashu-design/scripts/design-gate-hook.sh`
   اسکریپت‌های شبکه‌ای/هوک هستند و اجرا نشده‌اند.
4. فایل‌های نمونهٔ سنگین حذف شدند: PDF ۱۹۰KB در `academic-pptx-skill`، حدود ۲۷MB
   فایل صوتی `.mp3` در `huashu-design`، و تصاویر/PPTX نمونهٔ `powerpoint-fa`.
5. مجوز دو پوشه (`marp-slides`، `powerpoint-fa`) در مخزن مبدأ فایل `LICENSE` ندارد —
   `LICENSE-STATUS.md` داخل همان پوشه‌ها را ببینید.
