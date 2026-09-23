# Handoff

## معماری منتخب
**هسته ماژولار VBA + صفحه ابزار وب برای کشف/انتخاب/کپی + مسیر ارتقا به .dotm و Office Add-in**

این معماری برای Word دسکتاپ سریع و عملی است، کدها را قابل ممیزی نگه می‌دارد و در عین حال کاربر غیر فنی می‌تواند از ghelegh.ir ماژول مناسب را انتخاب و کپی کند.

## وضعیت
- آرشیو ورودی ۲۰ فایل داشت: VBA، JavaScript مرورگر، PowerShell و دو تصویر مرجع.
- دو فایل `ToFoot...Ok` بایت‌به‌بایت یکسان بودند.
- `combined-vba-scripts.txt` نام‌های procedure تکراری دارد و برای import مستقیم مناسب نیست.
- فایل‌های دارای `Ok` یا `Test good` به‌عنوان user-tested legacy شناخته می‌شوند؛ تاریخ تست در آرشیو نبود.
- Production جدید با پیشوند `RVA_` و فایل‌های BAS تماماً ASCII ساخته شده است.
- محیط فعلی Microsoft Word ندارد؛ بنابراین وضعیت فعلی ماژول جدید «Static checked» است.

## ادامه کار
1. اولین تست انسانی: `SmartHeadingDetector.bas`.
2. در false positive، heuristic ضعیف‌تر شود؛ جا افتادن یک تیتر مشکوک بهتر از Heading شدن متن بدنه است.
3. هر نشریه/دانشگاه Preset مستقل داشته باشد.
4. تغییرات Page Setup، Section، Footnote و Reference باید backup/undo strategy داشته باشند.
