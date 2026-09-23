# معماری محصول Research VBA

## ماتریس تصمیم

امتیاز کل از ۱۰۰ بر پایه سهولت برای کاربر غیر فنی، ایمنی و اعتماد، نگهداری‌پذیری، پوشش پلتفرم و سرعت رسیدن به ارزش است.

| رتبه | راهکار | امتیاز | مزیت اصلی | ریسک / اشتباه رایج |
|---:|---|---:|---|---|
| 1 | VBA ماژولار + Web Workbench زنده + مسیر .dotm و Office Add-in | 96 | ارزش فوری در Word و تجربه ساده بدون Git، با مسیر رشد روشن | اگر Web UI با کد repo همگام نباشد اعتماد از بین می‌رود؛ بنابراین صفحه کد را مستقیم از main می‌خواند |
| 2 | قالب Word امضاشده .dotm با Ribbon | 92 | نصب و اجرای بسیار ساده برای Windows Word | انتشار/اعتماد Macro و به‌روزرسانی نسخه‌ها نیازمند امضای دیجیتال و فرآیند Release است |
| 3 | Office Add-in با Task Pane از روز اول | 88 | UI حرفه‌ای و cross-platform | هزینه توسعه بیشتر و همه قابلیت‌های VBA/Word desktop الزاماً یک‌به‌یک در Office.js نیست |
| 4 | فقط مخزن GitHub با فایل‌های BAS | 78 | ساده برای توسعه‌دهنده و شفاف | برای پژوهشگر ناآشنا با Git اصطکاک زیاد دارد |
| 5 | فقط HTML مولد/نمایش‌دهنده کد | 72 | دسترسی سریع و بدون نصب | بدون لایه نصب/اعتماد در Word، محصول نهایی ناقص می‌ماند |

## انتخاب

راهکار رتبه ۱ مبناست: هسته Production در `modules/`، مستندات و CI در مخزن VBA، و صفحه `/tools/research-vba/` در ghelegh.ir که آخرین source را مستقیم از شاخه main می‌خواند. فاز بعدی برای Windows یک `.dotm` امضاشده با Ribbon است و سپس Office Add-in برای تجربه چندسکویی.

Microsoft، Word Add-ins را مبتنی بر HTML/CSS/JavaScript و قابل اجرا روی Word در وب، Windows، Mac و iPad مستند کرده است:
https://learn.microsoft.com/en-us/office/dev/add-ins/word/

برای توزیع عمومی VBA نیز Code Signing باید بخشی از Release engineering باشد:
https://support.microsoft.com/en-us/office/digitally-sign-your-vba-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01

## خطاهای رایج که در این مخزن ممنوع/کاهش داده می‌شوند

- یک ماکروی بسیار بزرگ برای چند مسئولیت نامرتبط.
- استفاده زیاد از `Selection` به جای `Range`.
- `On Error Resume Next` گسترده که خطا را پنهان می‌کند.
- ساخت Style سفارشی به جای Heading استاندارد Word برای TOC/Navigation.
- ترکیب قوانین چند دانشگاه/نشریه در یک Preset.
- تغییرات مخرب بدون هشدار، backup یا مسیر بازگشت.
- اعلام «Tested» بر اساس lint/CI به جای اجرای واقعی Word.
- hard-code متن فارسی داخل BAS؛ Production باید ASCII-only باشد.
- تشخیص تهاجمی Heading که false positive متن بدنه تولید کند.
- انتشار Macro عمومی بدون برنامه امضای دیجیتال و اعتماد کاربر.

## مسیر محصول

نسخه 0.x: BAS ماژولار + Audit + صفحه ابزار زنده.  
نسخه 1.0: تست Word روی اسناد واقعی، Release ZIP، نمونه قبل/بعد و .dotm امضاشده.  
نسخه 2.x: Task Pane با Office.js، Preset marketplace، sync نسخه‌ها و telemetry صرفاً opt-in و بدون ارسال متن سند.
