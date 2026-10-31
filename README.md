# FoodOrder - سیستم سفارش غذا

یک برنامه مدیریت سفارش غذا مبتنی بر Django برای مدیریت سفارشات هفتگی کارمندان.

## ویژگی‌ها

- **مدیریت کاربران**: احراز هویت و مدیریت پروفایل کاربران
- **سفارش هفتگی**: امکان ثبت سفارش غذا برای روزهای هفته
- **پنل ادمین**: مدیریت منوها، سفارشات و کاربران
- **رپورت‌گیری**: گزارش‌های روزانه و هفتگی سفارشات
- **پشتیبانی از RTL**: رابط کاربری کاملاً فارسی و راست‌چین

## تکنولوژی‌ها

- **Backend**: Django 4.x
- **Database**: SQLite (پیش‌فرض) / PostgreSQL
- **Frontend**: HTML, CSS, Bootstrap 5, JavaScript
- **Authentication**: Django Auth System

## نصب و اجرا

### پیش‌نیازها
- Python 3.8+
- pip

### مراحل نصب

```bash
# کلون کردن مخزن
git clone https://github.com/khanfini/FoodOrder.git
cd FoodOrder

# ایجاد محیط مجازی
python -m venv venv
venv\Scripts\activate  # در ویندوز
# source venv/bin/activate  # در لینوکس/مک

# نصب وابستگی‌ها
pip install -r requirements.txt

# اجرای میگریشن‌ها
python manage.py migrate

# ایجاد سوپریوزر (اختیاری)
python manage.py createsuperuser

# اجرای سرور توسعه
python manage.py runserver
```

## ساختار پروژه

```
FoodOrder/
├── appfoodorder/          # اپ اصلی برنامه
│   ├── models.py          # مدل‌های دیتابیس
│   ├── views.py           # ویوهای برنامه
│   ├── urls.py            # مسیریابی
│   ├── admin.py           # تنظیمات ادمین
│   └── management/        # دستورات مدیریت
├── templates/             # قالب‌های HTML
├── static/                # فایل‌های استاتیک
├── media/                 # فایل‌های آپلود شده
└── manage.py              # نقطه ورود جنگو
```

## مدل‌های اصلی

- **UserProfile**: پروفایل کاربری با اطلاعات اضافه
- **FoodItem**: آیتم‌های منوی غذا
- **WeeklyOrder**: سفارشات هفتگی کاربران
- **DailyMenu**: منوی روزانه

## دستورات مدیریت

```bash
# فعال‌سازی ادمین پیش‌فرض
python manage.py activate_admin
```

## مشارکت

1. Fork کنید
2. Branch جدید بسازید (`git checkout -b feature/amazing-feature`)
3. تغییرات را Commit کنید (`git commit -m 'Add amazing feature'`)
4. Push کنید (`git push origin feature/amazing-feature`)
5. Pull Request باز کنید

## لایسنس

این پروژه تحت لایسنس MIT منتشر شده است.

## تماس

- GitHub: [@khanfini](https://github.com/khanfini)
- Repository: [FoodOrder](https://github.com/khanfini/FoodOrder)