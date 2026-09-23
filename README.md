# Research VBA — ابزارهای Word برای پژوهشگران فارسی‌زبان

این مخزن مجموعه‌ای ماژولار از ماکروهای VBA برای آماده‌سازی مقاله، رساله و متن پژوهشی در Microsoft Word است. فایل‌های `.bas` عمداً فقط ASCII هستند تا مشکل کاراکتر فارسی در VBA ایجاد نشود.

## شروع سریع
1. از سند اصلی کپی بگیرید.
2. Word: `Alt+F11` → `File > Import File`.
3. فایل BAS موردنیاز را وارد کنید.
4. با `Alt+F8` ماکرو را اجرا کنید.
5. خروجی را قبل از ذخیره نهایی مرور کنید.

## ماژول‌های اصلی
| ماژول | کاربرد | وضعیت |
|---|---|---|
| [SmartHeadingDetector.bas](modules/article/SmartHeadingDetector.bas) | تشخیص تیتر و اعمال Heading 1–3 استاندارد برای TOC/Navigation | Static checked |
| [FERSDArticleFormat.bas](modules/article/FERSDArticleFormat.bas) | قالب و Audit شیوه‌نامه فصلنامه | Static checked |
| [LatinParenthesesToFootnotes.bas](modules/footnotes/LatinParenthesesToFootnotes.bas) | انتقال معادل لاتین داخل پرانتز به پانویس | rewrite of user-tested legacy |
| [ManualHeadingSelector.bas](modules/article/ManualHeadingSelector.bas) | fallback دستی برای تیترهای مبهم | based on user-tested legacy |
| [ThesisFontPreset.bas](modules/thesis/ThesisFontPreset.bas) | قلم‌های اصلی رساله بر اساس مرجع ارسالی | Static checked |
| [InlineComments.bas](modules/comments/InlineComments.bas) | Comment → inline markers | destructive |
| [OneNoteCleanup.bas](modules/cleanup/OneNoteCleanup.bas) | پاک‌سازی محافظه‌کارانه خروجی OneNote | Static checked |

## اسناد
- [RULES.md](RULES.md)
- [معماری و ماتریس ۵ راهکار](docs/ARCHITECTURE.md)\n- [کاتالوگ و امتیازدهی](docs/CODE_CATALOG.md)\n- [ممیزی استاتیک آرشیو قدیمی](docs/LEGACY_AUDIT.md)
- [پرامپت‌های مکمل](docs/PROMPTS.md)
- [BACKLOG.md](BACKLOG.md)
- [HANDOFF.md](HANDOFF.md)
- [Legacy inventory](legacy/README.md)

## کنترل کیفیت
`python scripts/check_vba_repo.py` بررسی می‌کند فایل‌های BAS فقط ASCII باشند، `Option Explicit` داشته باشند، نام Public procedure تکراری نباشد و کامنت `Tested on` فرمت تاریخ درست داشته باشد.

**Static check جای تست واقعی در Word را نمی‌گیرد.** کامنت `Tested on YYYY-MM-DD` فقط بعد از تأیید اجرای موفق توسط کاربر اضافه می‌شود.
