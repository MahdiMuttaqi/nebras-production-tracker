# راه‌اندازی رایگان سامانه نبراس

1. فایل Google Sheet «داده‌های سامانه پیگیری تولید نبراس» را باز کنید.
2. از منوی «افزونه‌ها» وارد Apps Script شوید.
3. محتوای فایل `apps-script/Code.gs` را جایگزین کد پیش‌فرض کنید و ذخیره کنید.
4. در بالای ویرایشگر، تابع `setupAdmin` را انتخاب و اجرا کنید؛ سپس رمز دلخواه مدیر را در پنجره بازشده وارد کنید.
5. از Deploy > New deployment نوع Web app را انتخاب کنید:
   - Execute as: Me
   - Who has access: Anyone
6. نشانی پایان‌یافته به `/exec` را کپی و در `web/config.js` قرار دهید.
7. در GitHub، Settings > Pages، منبع را روی GitHub Actions قرار دهید.
