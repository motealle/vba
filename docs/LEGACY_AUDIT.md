# ممیزی استاتیک آرشیو اولیه

> این جدول نتیجه تحلیل فایل‌های داخل ZIP ورودی است. «تست‌شده توسط کاربر» از نام فایل‌های `Ok` یا `Test good` استنباط شده؛ تاریخ تست موجود نبود. این ممیزی جای اجرای واقعی در Microsoft Word را نمی‌گیرد.

| فایل | نوع | خط | Procedure | Option Explicit | Resume Next | Selection | وضعیت/ریسک |
|---|---:|---:|---:|---:|---:|---:|---|
| 01-SwichStyles-ParantestoFootNote-and Aculad-Test good.txt | VBA | 720 | 9 | ✓ | 7 | 3 | user-tested legacy؛ چند مسئولیت در یک فایل |
| 01-SwichStyles-ParantestoFootNote-Ok.txt | VBA | 509 | 9 | ✓ | 7 | 3 | user-tested legacy؛ نیازمند تفکیک ماژول‌ها |
| 02-MultiLevelList.bas.txt | VBA | 236 | 2 | ✗ | 7 | 0 | ایده خوب، error suppression زیاد |
| 03-Inlining comments.txt | VBA | 48 | 1 | ✗ | 0 | 0 | ساده ولی مخرب؛ Comment حذف می‌شود |
| 04-Thesis-Style-Set.txt | VBA | 620 | 10 | ✓ | 8 | 3 | پایه مناسب Preset رساله؛ بیش از حد یکپارچه |
| aistudio paste to word farsi text problem VBA .txt | VBA | 97 | 2 | ✗ | 1 | 5 | encoding مشکل‌دار + Selection-heavy |
| collect-vba-scripts.ps1 | PowerShell | 75 | — | — | — | — | Utility؛ باید خروجی deduplicate شود |
| combined-vba-scripts.txt | VBA | 3838 | 70 | ✓ | 38 | 24 | **برای Import مستقیم نامناسب**؛ بیش از ۲۰ نام procedure تکراری |
| ToFoot.bas--Pattern (Latin) Ok.txt | VBA | 182 | 3 | ✓ | 0 | 0 | user-tested legacy؛ پایه خوب |
| ToFoot.bas-Pattern (Latin) {گیومه} Ok.txt | VBA | 182 | 3 | ✓ | 0 | 0 | بایت‌به‌بایت تکراری فایل قبلی |
| ToFoot.bas.txt | VBA | 164 | 2 | ✓ | 0 | 0 | جد نسخه تست‌شده |
| اسکریپت باز کردن چکیده در سایت انسانی | JavaScript | 1 | — | — | — | — | DOM-dependent |
| اسکریپت تیک زدن در کنسول سامانه گنج | JavaScript | 39 | — | — | — | — | Browser-specific و وابسته به DOM |
| اسکریپت پهن کردن پنجره پاسخ پرامپت | JavaScript | 9 | — | — | — | — | UI tweak؛ شکننده در برابر تغییر selector |
| تنظیم استایل سطح عناوین داخلی با علامت کوچک/بزرگ | VBA | 90 | 1 | ✗ | 0 | 0 | ورودی خاص؛ بهتر است fallback باشد |
| ست کردن هدینگ به تیترهای ساختاری مقاله-Ok | VBA | 83 | 2 | ✗ | 0 | 2 | user-tested legacy؛ مبنای ManualHeadingSelector |
| کد تنظیم فونت قالب مقاله vba-Ok | VBA | 175 | 3 | ✗ | 5 | 0 | user-tested legacy؛ باید به Preset نشریه تفکیک شود |
| مرتب کردن نسبی فایل خروجی یافته از وان نوت | VBA | 212 | 1 | ✓ | 3 | 6 | تغییرات تهاجمی؛ نسخه Production محافظه‌کارانه‌تر شد |
| استاندارد قلم متون رساله.png | تصویر | — | — | — | — | — | مرجع Preset رساله |
| نمونه پرامپت ساختار.png | تصویر | — | — | — | — | — | مرجع طراحی Prompt |

## یافته‌های ماشینی مهم

- دو فایل `ToFoot...Ok` SHA-256 یکسان دارند و duplicate واقعی‌اند.
- `combined-vba-scripts.txt` دارای ۷۰ procedure، ۳۸ مورد `On Error Resume Next` و ۲۴ اشاره به `Selection` است؛ چندین نام procedure در خود فایل تکرار شده است.
- چند فایل قدیمی `Option Explicit` ندارند؛ Production جدید این مورد را اجباری می‌کند.
- وجود کاراکتر غیر ASCII در فایل‌های قدیمی دلیل دیگری برای نگهداری آن‌ها در Legacy و بازنویسی Production به صورت ASCII-only است.
