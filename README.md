# DLRobo-IDM

یک ابزار هوشمند برای استخراج خودکار لینک‌های دانلودی از وب‌سایت‌ها و اضافه کردن آن‌ها به نرم‌افزار مدیریت دانلود IDM.

## ویژگی‌ها
- استخراج لینک‌های با پسوندهای مشخص (ZIP, RAR, PDF, MP3, MP4, EXE)
- پشتیبانی از پروکسی برای دسترسی به سایت‌های محدودشده
- ادغام مستقیم با IDM
- نمایش لاگ عملیات و نوار پیشرفت

## نیازمندی‌ها
- Python 3.x
- کتابخانه‌های: `tkinter`, `requests`, `beautifulsoup4`, `pywin32`, `urllib3`

## نحوه نصب
```bash
pip install requests beautifulsoup4 pywin32
python dlrobo-gui.py
