import jdatetime
from django.shortcuts import render, redirect
from django.urls import reverse
from django.contrib.auth import authenticate, login as auth_login, logout as auth_logout
from .models import Department, Member, CustomUser, Location, Order, Floor, get_order_settings, Catering, Campaign, Survey
from django import forms
from django.contrib.auth import get_user_model
from django.http import HttpResponse
from django.template.loader import render_to_string
import os
from django.conf import settings
import pandas as pd
import re
from jdatetime import timedelta
from django.views.decorators.csrf import ensure_csrf_cookie
from django.contrib.admin.views.decorators import staff_member_required
from django.http import HttpResponseBadRequest
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from glob import glob
# from .utils import load_menu_dict, get_weekday_name 
# ✅ تابع جدید: تبدیل رشته تاریخ به jdatetime.date
def parse_persian_date(date_str):
    """تبدیل رشتهای مثل '1405/02/07' به jdatetime.date"""
    if not isinstance(date_str, str):
        date_str = str(date_str)
    # تبدیل اعداد فارسی به انگلیسی (اگر وجود داشته باشد)
    trans = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
    date_str = date_str.translate(trans).strip()
    # جدا کردن سه بخش با جداکننده / یا -
    parts = re.split(r'[/\-]', date_str)
    if len(parts) != 3:
        return None
    try:
        year, month, day = map(int, parts)
        return jdatetime.date(year, month, day)
    except:
        return None


# ✅ تابع جدید: خواندن منو و ایجاد دیکشنری تاریخ → (Menu1, Menu2)
def load_menu_dict():
    excel_path = os.path.join(settings.BASE_DIR, 'media', 'Menu_01.xlsx')  # مسیر مطمئن: ریشه پروژه
    debug_info = []
    debug_info.append(f"مسیر فایل: {excel_path}")
    if not os.path.exists(excel_path):
        debug_info.append("❌ فایل وجود ندارد!")
        return {}, debug_info

    try:
        df = pd.read_excel(excel_path)
        debug_info.append(f"✅ فایل خوانده شد. تعداد سطرها: {len(df)}")
        debug_info.append(f"ستون‌ها: {list(df.columns)}")

        menu_dict = {}
        for idx, row in df.iterrows():
            raw_date = row['Date']
            debug_info.append(f"🗓 سطر {idx}: raw_date = {repr(raw_date)}")
            d = parse_persian_date(raw_date)
            if d is None:
                debug_info.append(f"   ❌ تاریخ نامعتبر → رد شد")
                continue
            menu1 = str(row['Menu1']).strip() if pd.notna(row['Menu1']) else ""
            menu2 = str(row['Menu2']).strip() if pd.notna(row['Menu2']) else ""
            menu_dict[d] = (menu1, menu2)
            debug_info.append(f"   ✅ تاریخ شمسی: {d} → ({menu1}, {menu2})")

        debug_info.append(f"🔑 کلیدهای نهایی: {list(menu_dict.keys())}")
        return menu_dict, debug_info
    except Exception as e:
        debug_info.append(f"💥 خطا در خواندن فایل: {e}")
        return {}, debug_info


# ✅ تابع جدید: خواندن منوهای هفتگی (Menu_YYYYMM-wN.xlsx)
def load_weekly_menu_dict(week_number=None, year_month=None):
    """
    بارگذاری منو از فایل‌های هفتگی
    نام فایل: Menu_140509-w1.xlsx
    """
    media_dir = os.path.join(settings.BASE_DIR, 'media')
    debug_info = []

    # اگر week_number و year_month داده نشده باشند، بر اساس تاریخ جاری محاسبه می‌کنیم
    if week_number is None or year_month is None:
        today = jdatetime.date.today()
        year_month = today.strftime("%Y%m")
        # محاسبه هفته ماه (۱ تا ۴)
        day_of_month = today.day
        week_number = (day_of_month - 1) // 7 + 1
        if week_number > 4:
            week_number = 4

    # جستجوی فایل
    pattern = os.path.join(media_dir, f'Menu_{year_month}-w{week_number}.xlsx')
    files = glob(pattern)

    if not files:
        debug_info.append(f"❌ فایل برای {year_month} هفته {week_number} یافت نشد")
        return {}, debug_info, year_month, week_number

    excel_path = files[0]
    debug_info.append(f"✅ فایل یافت شد: {excel_path}")

    try:
        df = pd.read_excel(excel_path)
        debug_info.append(f"✅ فایل خوانده شد. تعداد سطرها: {len(df)}")
        debug_info.append(f"ستون‌ها: {list(df.columns)}")

        menu_dict = {}
        for idx, row in df.iterrows():
            raw_date = row['Date']
            debug_info.append(f"🗓 سطر {idx}: raw_date = {repr(raw_date)}")
            d = parse_persian_date(raw_date)
            if d is None:
                debug_info.append(f"   ❌ تاریخ نامعتبر → رد شد")
                continue
            menu1 = str(row['Menu1']).strip() if pd.notna(row['Menu1']) else ""
            menu2 = str(row['Menu2']).strip() if pd.notna(row['Menu2']) else ""
            menu_dict[d] = (menu1, menu2)
            debug_info.append(f"   ✅ تاریخ شمسی: {d} → ({menu1}, {menu2})")

        debug_info.append(f"🔑 کلیدهای نهایی: {list(menu_dict.keys())}")
        return menu_dict, debug_info, year_month, week_number

    except Exception as e:
        debug_info.append(f"💥 خطا در خواندن فایل: {e}")
        return {}, debug_info, year_month, week_number


# ✅ تابع جدید: دریافت کمپین فعال
def get_active_campaign():
    today = jdatetime.date.today()
    today_greg = today.togregorian()
    campaign = Campaign.objects.filter(
        is_active=True,
        start_date__lte=today_greg,
        end_date__gte=today_greg
    ).select_related('catering').first()
    return campaign


def get_current_week_campaign():
    """کمپین هفته شمسی جاری، حتی اگر امروز جمعه باشد و بازه کمپین تمام شده باشد."""
    today = jdatetime.date.today()
    week_start = today - timedelta(days=today.weekday())   # شنبه
    week_end = week_start + timedelta(days=5)              # پنجشنبه
    return Campaign.objects.filter(
        is_active=True,
        start_date__lte=week_end.togregorian(),
        end_date__gte=week_start.togregorian(),
    ).select_related('catering').order_by('start_date').first()


def get_campaign_menu_path(campaign):
    """مسیر فایل اکسل منوی هفته کمپین."""
    start_jalali = jdatetime.date.fromgregorian(date=campaign.start_date)
    year_month = start_jalali.strftime("%Y%m")
    return os.path.join(settings.BASE_DIR, 'media', f'Menu_{year_month}-w{campaign.week_number}.xlsx')


def save_campaign_menu(campaign, rows):
    """ذخیره نام غذاهای ویرایش‌شده در فایل اکسل همان هفته."""
    excel_path = get_campaign_menu_path(campaign)
    if not os.path.exists(excel_path):
        return False
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    headers = {
        str(cell.value).strip(): index
        for index, cell in enumerate(sheet[1], start=1)
        if cell.value is not None
    }
    required = {'Date', 'Menu1', 'Menu2'}
    if not required.issubset(headers):
        return False

    row_by_date = {}
    for row_number in range(2, sheet.max_row + 1):
        parsed = parse_persian_date(sheet.cell(row=row_number, column=headers['Date']).value)
        if parsed is not None:
            row_by_date[parsed.strftime("%Y%m%d")] = row_number

    for item in rows:
        row_number = row_by_date.get(item['date_str'])
        if row_number is None:
            continue
        sheet.cell(row=row_number, column=headers['Menu1']).value = item['menu1'] or None
        sheet.cell(row=row_number, column=headers['Menu2']).value = item['menu2'] or None
    workbook.save(excel_path)
    return True


# سفارش‌ها و وضعیت نظرسنجی هفته جاری
def get_survey_orders(member, campaign):
    """غذاهای سفارش‌داده‌شده از ابتدای هفته جاری تا امروز."""
    today = jdatetime.date.today()
    menu_dates = [date for date in get_campaign_menu_dict(campaign) if date <= today]
    return Order.objects.filter(
        member=member,
        date__in=[date.togregorian() for date in menu_dates],
    ).exclude(choice_food=0).order_by('date')


def check_survey_completed(member, campaign):
    """همه غذاهای قابل‌نظرسنجی تا امروز امتیاز گرفته‌اند."""
    if not campaign:
        return True

    member_orders = get_survey_orders(member, campaign)
    if not member_orders.exists():
        return True

    return Survey.objects.filter(
        member=member,
        campaign=campaign,
        date__in=[order.date for order in member_orders],
    ).count() >= member_orders.count()


def get_next_campaign(campaign):
    """کمپین بعدی می‌تواند متعلق به کترینگ متفاوتی باشد."""
    return Campaign.objects.filter(
        is_active=True,
        start_date__gt=campaign.start_date,
    ).select_related('catering').order_by('start_date').first()


# ✅ تابع جدید: دریافت تاریخ‌های هفته برای کمپین
def get_week_dates_for_campaign(campaign):
    """دریافت تمام تاریخ‌های واقعی موجود در فایل منوی کمپین."""
    return sorted(get_campaign_menu_dict(campaign))


# ✅ تابع جدید: دریافت منو برای هفته کمپین
def get_campaign_menu_dict(campaign):
    """دریافت منو از فایل هفته کمپین"""
    start_jalali = jdatetime.date.fromgregorian(date=campaign.start_date)
    year_month = start_jalali.strftime("%Y%m")
    menu_dict, _, _, _ = load_weekly_menu_dict(campaign.week_number, year_month)
    return menu_dict


def get_daily_menu_names(jalali_date=None):
    """نام غذاهای امروز از فایل هفته کمپین؛ در صورت نبود فایل هفتگی از منوی عمومی."""
    if jalali_date is None:
        jalali_date = jdatetime.date.today()
    today_greg = jalali_date.togregorian()
    campaign = Campaign.objects.filter(
        is_active=True,
        start_date__lte=today_greg,
        end_date__gte=today_greg,
    ).select_related('catering').order_by('start_date').first()
    if campaign:
        menu_dict = get_campaign_menu_dict(campaign)
        if jalali_date in menu_dict:
            return menu_dict[jalali_date]
    menu_dict, _ = load_menu_dict()
    return menu_dict.get(jalali_date, ("", ""))


def get_request_member_id(request):
    """شناسه کارمند انتخاب‌شده از POST یا GET."""
    member_id = request.POST.get('member_id') if request.method == 'POST' else request.GET.get('member_id')
    if member_id and str(member_id).isdigit():
        return int(member_id)
    return None


def get_member_for_user(user, member_id=None):
    """
    کارمندی که کاربر مجاز است برایش سفارش/نظرسنجی ثبت کند.
    کاربران عادی حساب دپارتمان هستند (مثل IT یا hr) و نامشان با نام کارمند یکی نیست،
    پس فقط اعضای دپارتمان خودشان قابل انتخاب است.
    """
    if user.is_superuser or user.is_staff:
        members = Member.objects.all()
    elif user.department_id:
        members = Member.objects.filter(department_id=user.department_id)
    else:
        return None

    if member_id:
        return members.filter(id=member_id).select_related('department').first()
    # سازگاری با حالتی که نام کاربری دقیقاً برابر نام کارمند باشد
    return members.filter(name=user.username).select_related('department').first()


def to_persian_digits(value):
    return str(value).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))


def get_campaign_promo():
    """اطلاعات بنر تبلیغاتی کمپین نظرسنجی برای صفحه ورود."""
    campaign = get_active_campaign()
    if not campaign:
        return None
    weeks = []
    for number in range(1, 5):
        if number < campaign.week_number:
            state = 'done'
        elif number == campaign.week_number:
            state = 'current'
        else:
            state = 'upcoming'
        weeks.append({'label': to_persian_digits(number), 'state': state})
    return {
        'catering': campaign.catering.name,
        'week': to_persian_digits(campaign.week_number),
        'weeks': weeks,
    }


@ensure_csrf_cookie
def login(request):
    menu_today, debug_info = get_menu_for_today()
    promo = get_campaign_promo()
    if request.method == "POST":
        username = request.POST.get("username")
        password = request.POST.get("password")
        user = authenticate(request, username=username, password=password)
        if user is not None:
            auth_login(request, user)

            # نظرسنجی هر کارمند در صفحه سفارش هفتگی، بعد از انتخاب کارمند نمایش داده می‌شود.
            return redirect("weekly_order")
        else:
            return render(request, "login.html", {"error": "نام کاربری یا رمز عبور اشتباه است", "menu_today": menu_today, "promo": promo})
    return render(request, "login.html", {"menu_today": menu_today, "promo": promo})


def logout_view(request):
    auth_logout(request)
    return redirect("login")


# ✅ ویو نظرسنجی
def survey(request):
    if not request.user.is_authenticated:
        return redirect("login")

    user = request.user

    # انتخاب عضو (ادمین: همه اعضا، کاربر دپارتمان: فقط اعضای دپارتمان خودش)
    member = get_member_for_user(user, get_request_member_id(request))
    if member is None:
        return redirect("weekly_order")

    campaign = get_active_campaign()
    if not campaign:
        return redirect("weekly_order")

    # بررسی تکمیل نظرسنجی غذاهای سفارش‌داده‌شده هفته جاری
    if check_survey_completed(member, campaign):
        return redirect("weekly_order")

    # دریافت منو هفته جاری (برای نظرسنجی)
    current_menu_dict = get_campaign_menu_dict(campaign)
    current_week_dates = get_week_dates_for_campaign(campaign)

    # دریافت کمپین بعدی؛ کترینگ هر هفته متفاوت است.
    next_campaign = get_next_campaign(campaign)

    next_menu_dict = {}
    next_week_dates = []
    if next_campaign:
        next_menu_dict = get_campaign_menu_dict(next_campaign)
        next_week_dates = get_week_dates_for_campaign(next_campaign)

    # هر غذای سفارش‌داده‌شده در هفته جاری یک نظرسنجی مستقل دارد.
    member_orders = get_survey_orders(member, campaign)

    # بررسی نظرسنجی‌های قبلی
    existing_surveys = Survey.objects.filter(
        member=member,
        campaign=campaign,
        date__in=[o.date for o in member_orders]
    )
    survey_dict = {s.date: s for s in existing_surveys}

    # پردازش POST ثبت نظرسنجی
    if request.method == 'POST':
        saved = False
        for order in member_orders:
            if order.date in survey_dict:
                continue

            # نام فیلد در قالب با تاریخ شمسی ساخته می‌شود (مثل rating_14050902)
            jalali_date = jdatetime.date.fromgregorian(date=order.date)
            rating_key = f'rating_{jalali_date.strftime("%Y%m%d")}'
            rating_value = request.POST.get(rating_key)
            if rating_value not in dict(Survey.RATING_CHOICES):
                continue

            Survey.objects.create(
                member=member,
                campaign=campaign,
                date=order.date,
                menu_choice=order.choice_food,
                rating=rating_value,
            )
            saved = True

        if saved:
            # هدایت به صفحه سفارش با حفظ کارمند انتخاب‌شده
            params = []
            dept_id = request.POST.get('department_id')
            if is_order_admin(user) and dept_id:
                params.append(f'department_id={dept_id}')
            params.append(f'member_id={member.id}')
            return redirect(reverse('weekly_order') + '?' + '&'.join(params))

    # آماده‌سازی داده‌ها برای قالب
    survey_data = []
    for order in member_orders:
        jd = jdatetime.date.fromgregorian(date=order.date)
        menu1, menu2 = current_menu_dict.get(jd, ("", ""))
        menu_name = menu1 if order.choice_food == 1 else menu2 if order.choice_food == 2 else "حضور ندارد"
        existing_survey = survey_dict.get(order.date)

        survey_data.append({
            'date': jd,
            'weekday': get_weekday_name(jd.weekday()),
            'date_display': jd.strftime("%Y/%m/%d"),
            'date_str': jd.strftime("%Y%m%d"),
            'menu_choice': order.get_choice_food_display(),
            'menu_name': menu_name,
            'existing_rating': existing_survey.rating if existing_survey else None,
            'is_completed': existing_survey is not None,
        })

    # آماده‌سازی داده‌های هفته بعد برای سفارش
    next_week_data = []
    if next_campaign and next_week_dates:
        next_gregorian_dates = [d.togregorian() for d in next_week_dates]
        next_existing_orders = Order.objects.filter(member=member, date__in=next_gregorian_dates)
        next_order_dict = {o.date: o for o in next_existing_orders}

        for d in next_week_dates:
            menu1, menu2 = next_menu_dict.get(d, ("", ""))
            existing_order = next_order_dict.get(d.togregorian())

            next_week_data.append({
                'date': d,
                'weekday': get_weekday_name(d.weekday()),
                'weekday_num': d.weekday(),
                'date_display': d.strftime("%Y/%m/%d"),
                'date_str': d.strftime("%Y%m%d"),
                'menu1': menu1 if menu1 else "—",
                'menu2': menu2 if menu2 else "—",
                'existing_choice': existing_order.choice_food if existing_order else None,
            })

    context = {
        'campaign': campaign,
        'campaign_start_display': jdatetime.date.fromgregorian(date=campaign.start_date).strftime("%Y/%m/%d"),
        'campaign_end_display': jdatetime.date.fromgregorian(date=campaign.end_date).strftime("%Y/%m/%d"),
        'survey_data': survey_data,
        'next_week_data': next_week_data,
        'next_campaign': next_campaign,
        'RATING_CHOICES': Survey.RATING_CHOICES,
    }
    return render(request, "survey.html", context)


# ✅ تابع جدید: دریافت منو بر اساس تاریخ دقیق
def get_menu_for_today():
    today = jdatetime.date.today()
    menu1_name, menu2_name = get_daily_menu_names(today)
    return (menu1_name, menu2_name), []


# ✅ تابع جدید: نام روز هفته
def get_weekday_name(weekday):
    return ["شنبه", "یکشنبه", "دوشنبه", "سه شنبه", "چهارشنبه", "پنجشنبه", "جمعه"][weekday]


# ✅ فرم ثبت سفارش (تغییر نکرده)
class OrderForm(forms.Form):
    def __init__(self, *args, **kwargs):
        user = kwargs.pop('user', None)
        super().__init__(*args, **kwargs)
        if is_order_admin(user):
            self.fields['member'].queryset = Member.objects.all()
        elif user.department:
            self.fields['member'].queryset = Member.objects.filter(department=user.department)
        else:
            self.fields['member'].queryset = Member.objects.none()

    member = forms.ModelChoiceField(queryset=Member.objects.none(), label="نام پرسنل")
    choice_food = forms.ChoiceField(
        choices=[(1, "منوی غذایی ۱"), (2, "منوی غذایی ۲"), (3, "حضور ندارد")],
        label="گزینه غذا"
    )



def order(request):
    if not request.user.is_authenticated:
        return redirect("login")

    user = request.user
    today = jdatetime.date.today()
    today_weekday = today.weekday()
    today_date_str = today.strftime("%Y/%m/%d")
    # تبدیل تاریخ امروز به میلادی برای جستجو در Order
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # بارگذاری کامل منوها از اکسل
    menu_dict, debug_info = load_menu_dict()

    # منوی امروز (از منوی هفته کمپین)
    menu1_name, menu2_name = get_daily_menu_names(today)
    if not menu1_name:
        menu1_name = "منوی امروز ثبت نشده است"
    if not menu2_name:
        menu2_name = "منوی امروز ثبت نشده است"
    if today_weekday == 6:
        menu1_name = ""
        menu2_name = ""

    # منوهای آینده (فردا به بعد) – مثلاً ۷ روز آینده
    upcoming_menus = []
    for d in sorted(menu_dict.keys()):
        if d > today:
            m1, m2 = menu_dict[d]
            upcoming_menus.append({
                'date': d.strftime("%Y/%m/%d"),
                'weekday': get_weekday_name(d.weekday()),
                'menu1': m1 if m1 else "—",
                'menu2': m2 if m2 else "—",
            })
    upcoming_menus = upcoming_menus[:7]

    # دسته‌بندی اعضا
    if is_order_admin(user):
        departments = Department.objects.all()
        members_by_department = []
        for dept in departments:
            members = Member.objects.filter(department=dept)
            # برای هر عضو، سفارش امروز را از Order می‌خوانیم و به یک attribute موقت اضافه می‌کنیم
            for m in members:
                try:
                    order_obj = Order.objects.get(member=m, date=today_greg)
                    m.today_choice = order_obj.choice_food
                except Order.DoesNotExist:
                    m.today_choice = None  # یعنی ثبت نشده
            members_by_department.append((dept, members))
    else:
        user_department = user.department
        if not user_department:
            return render(request, "order.html", {"error": "شما دپارتمانی ندارید"})
        departments = Department.objects.filter(id=user_department.id)
        members = Member.objects.filter(department=user_department)
        for m in members:
            try:
                order_obj = Order.objects.get(member=m, date=today_greg)
                m.today_choice = order_obj.choice_food
            except Order.DoesNotExist:
                m.today_choice = None
        members_by_department = [(user_department, members)]

    # پردازش POST
    if request.method == 'POST':
        form = OrderForm(user=user, data=request.POST)
        if form.is_valid():
            member = form.cleaned_data['member']
            choice_food = form.cleaned_data['choice_food']  # اینجا 1 یا 2 یا 3 می‌آید

            if not is_order_admin(user) and member.department != user.department:
                return render(request, 'order.html', {
                    'error': 'شما فقط مجاز به ثبت سفارش برای اعضای دپارتمان خود هستید.',
                    'form': form,
                    'members': members_by_department,
                    'departments': departments,
                    'today_date_str': today_date_str,
                    'today_weekday': today_weekday,
                    'menu1_name': menu1_name,
                    'menu2_name': menu2_name,
                    'upcoming_menus': upcoming_menus,
                })

            # ذخیره در جدول Order
            Order.objects.update_or_create(
                member=member,
                date=today_greg,
                defaults={'choice_food': choice_food}
            )
            return redirect('order')
    else:
        form = OrderForm(user=user)

    user_location = user.department.location if user.department else None

    # آمار: تعداد سفارش‌های امروز از جدول Order
    orders_today = Order.objects.filter(date=today_greg)
    total_members = Member.objects.count()
    count_menu1 = orders_today.filter(choice_food=1).count()
    count_menu2 = orders_today.filter(choice_food=2).count()
    count_absent = orders_today.filter(choice_food=0).count()  # توجه: 0 یعنی غایب

    # توجه: اگر عضوی اصلاً سفارش نداده باشد، در آمار غایب محسوب نمی‌شود.
    # برای محاسبه تعداد افرادی که هیچ سفارشی ندارند (Not ordered) می‌توانید:
    members_with_order = orders_today.values_list('member', flat=True).distinct()
    not_ordered_count = Member.objects.exclude(id__in=members_with_order).count()
    # ولی فعلاً همان count_absent را نشان می‌دهیم (تعداد غایب‌هایی که صراحتاً انتخاب کرده‌اند)

    context = {
        'user': user,
        'departments': departments,
        'members': members_by_department,
        'form': form,
        'today_date_str': today_date_str,
        'today_weekday': today_weekday,
        'weekday_str': get_weekday_name(today_weekday),
        'menu1_name': menu1_name,
        'menu2_name': menu2_name,
        'location': user_location,
        'total_members': total_members,
        'count_menu1': count_menu1,
        'count_menu2': count_menu2,
        'count_absent': count_absent,
        'upcoming_menus': upcoming_menus,
        'debug_info': debug_info,
    }
    return render(request, 'order.html', context)


def printable_list(request):
    if not is_order_admin(request.user):
        return redirect("order")

    # تاریخ امروز شمسی و میلادی
    today = jdatetime.date.today()
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # دریافت نام واقعی غذاهای امروز از منوی هفته کمپین
    menu1_name, menu2_name = get_daily_menu_names(today)

    # ----- دریافت پارامترهای فیلتر از GET -----
    selected_location_id = request.GET.get('location')
    selected_department_id = request.GET.get('department')
    selected_floor_id = request.GET.get('floor')

    # ----- ساخت کوئری اولیه سفارش‌های امروز -----
    orders_qs = Order.objects.filter(date=today_greg).select_related('member__department__location', 'member__department__floor')

    # اعمال فیلتر ساختمان (Location)
    if selected_location_id and selected_location_id.isdigit():
        orders_qs = orders_qs.filter(member__department__location_id=int(selected_location_id))

    # اعمال فیلتر طبقه (Floor)
    if selected_floor_id and selected_floor_id.isdigit():
        orders_qs = orders_qs.filter(member__department__floor_id=int(selected_floor_id))

    # اعمال فیلتر دپارتمان (Department)
    if selected_department_id and selected_department_id.isdigit():
        orders_qs = orders_qs.filter(member__department_id=int(selected_department_id))

    # ----- گروه‌بندی بر اساس ساختمان و دپارتمان (با در نظر گرفتن فیلترها) -----
    departments_data = {}
    for order in orders_qs:
        member = order.member
        dept = member.department
        if not dept:
            continue
        location = dept.location
        location_name = location.name if location else "بدون ساختمان"
        dept_name = dept.name

        key = (location_name, dept_name)
        if key not in departments_data:
            departments_data[key] = {
                'location': location_name,
                'department': dept_name,
                'members': []
            }
        departments_data[key]['members'].append({
            'name': member.name,
            'code': member.code,           # ✅ کد پرسنلی اضافه شد
            'choice_food': order.choice_food
        })

    # تبدیل به لیست مرتب
    members_data = []
    for (loc, dept), data in sorted(departments_data.items()):
        data['members'].sort(key=lambda x: x['name'])
        members_data.append(data)

    # ----- داده‌های مورد نیاز برای فیلترها (جهت نمایش در dropdown) -----
    all_locations = Location.objects.all().order_by('name')
    all_floors = Floor.objects.all().order_by('name')
    
    if selected_location_id and selected_location_id.isdigit():
        all_departments = Department.objects.filter(location_id=int(selected_location_id)).order_by('name')
    else:
        all_departments = Department.objects.all().order_by('name')

    today_date_str = today.strftime("%Y/%m/%d")
    weekday_str = get_weekday_name(today.weekday())

    # ==========  محاسبه آمار تجمیعی ==========
    total_menu1 = orders_qs.filter(choice_food=1).count()
    total_menu2 = orders_qs.filter(choice_food=2).count()
    total_absent = orders_qs.filter(choice_food=0).count()
    total_orders = orders_qs.exclude(choice_food=0).count()
    total_members_with_choice = orders_qs.count()
    # =============================================

    html = render_to_string('printable_list.html', {
        'members_data': members_data,
        'date_str': today_date_str,
        'weekday_str': weekday_str,
        'menu1_name': menu1_name,
        'menu2_name': menu2_name,
        'all_locations': all_locations,
        'all_departments': all_departments,
        'all_floors': all_floors,
        'selected_location_id': selected_location_id,
        'selected_department_id': selected_department_id,
        'selected_floor_id': selected_floor_id,
        'total_menu1': total_menu1,
        'total_menu2': total_menu2,
        'total_absent': total_absent,
        'total_orders': total_orders,
        'total_members_with_choice': total_members_with_choice,
    })
    response = HttpResponse(html, content_type='text/html; charset=utf-8')
    response['Content-Disposition'] = f'inline; filename="لیست_غذا_{today_date_str}.html"'
    return response


def export_printable_list_excel(request):
    if not is_order_admin(request.user):
        return redirect("order")

    # ========== تاریخ امروز ==========
    today = jdatetime.date.today()
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # ========== دریافت نام واقعی غذاهای امروز از منوی هفته کمپین ==========
    menu1_name, menu2_name = get_daily_menu_names(today)

    # ========== دریافت پارامترهای فیلتر از GET ==========
    selected_location_id = request.GET.get('location')
    selected_department_id = request.GET.get('department')
    selected_floor_id = request.GET.get('floor')

    # ========== ساخت کوئری سفارش‌های امروز با فیلترها ==========
    orders_qs = Order.objects.filter(date=today_greg).select_related(
        'member__department__location', 'member__department__floor'
    )

    if selected_location_id and selected_location_id.isdigit():
        orders_qs = orders_qs.filter(member__department__location_id=int(selected_location_id))
    if selected_floor_id and selected_floor_id.isdigit():
        orders_qs = orders_qs.filter(member__department__floor_id=int(selected_floor_id))
    if selected_department_id and selected_department_id.isdigit():
        orders_qs = orders_qs.filter(member__department_id=int(selected_department_id))

    # ========== محاسبه آمار ==========
    total_menu1 = orders_qs.filter(choice_food=1).count()
    total_menu2 = orders_qs.filter(choice_food=2).count()
    total_absent = orders_qs.filter(choice_food=0).count()
    total_orders = orders_qs.exclude(choice_food=0).count()

    # ========== گروه‌بندی برای جزئیات ==========
    departments_data = {}
    for order in orders_qs:
        member = order.member
        dept = member.department
        if not dept:
            continue
        location = dept.location
        location_name = location.name if location else "بدون ساختمان"
        dept_name = dept.name
        key = (location_name, dept_name)
        if key not in departments_data:
            departments_data[key] = {
                'location': location_name,
                'department': dept_name,
                'members': []
            }
        departments_data[key]['members'].append({
            'name': member.name,
            'code': member.code,           # ✅ کد پرسنلی اضافه شد
            'choice_food': order.choice_food
        })

    # مرتب‌سازی
    members_data = []
    for (loc, dept), data in sorted(departments_data.items()):
        data['members'].sort(key=lambda x: x['name'])
        members_data.append(data)

    # ========== ایجاد فایل اکسل ==========
    wb = openpyxl.Workbook()

    # ---------- برگه خلاصه ----------
    ws_summary = wb.active
    ws_summary.title = "خلاصه سفارشات"
    ws_summary['A1'] = f"گزارش سفارش غذا - {today.strftime('%Y/%m/%d')}"
    ws_summary['A1'].font = Font(size=14, bold=True)
    ws_summary.merge_cells('A1:D1')

    # ردیف اطلاعات فیلترها
    row = 3
    ws_summary[f'A{row}'] = "فیلتر اعمال شده:"
    loc_name = "همه"
    if selected_location_id and selected_location_id.isdigit():
        try:
            loc_name = Location.objects.get(id=int(selected_location_id)).name
        except:
            pass
    dept_name = "همه"
    if selected_department_id and selected_department_id.isdigit():
        try:
            dept_name = Department.objects.get(id=int(selected_department_id)).name
        except:
            pass
    floor_name = "همه"
    if selected_floor_id and selected_floor_id.isdigit():
        try:
            floor_name = Floor.objects.get(id=int(selected_floor_id)).name
        except:
            pass

    ws_summary[f'B{row}'] = f"ساختمان: {loc_name}"
    ws_summary[f'C{row}'] = f"دپارتمان: {dept_name}"
    ws_summary[f'D{row}'] = f"طبقه: {floor_name}"

    # جدول آمار
    row += 2
    headers = ["نوع", "تعداد", "نام غذا", "توضیحات"]
    for col, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=row, column=col, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    row += 1

    data_stats = [
        ("منوی ۱", total_menu1, menu1_name or "منوی ۱", ""),
        ("منوی ۲", total_menu2, menu2_name or "منوی ۲", ""),
        ("حضور ندارد", total_absent, "", ""),
        ("کل سفارش غذا", total_orders, "", "بدون احتساب غایب‌ها"),
    ]
    for stat in data_stats:
        ws_summary.cell(row=row, column=1, value=stat[0])
        ws_summary.cell(row=row, column=2, value=stat[1])
        ws_summary.cell(row=row, column=3, value=stat[2])
        ws_summary.cell(row=row, column=4, value=stat[3])
        row += 1

    # ---------- برگه جزئیات (با ستون کد پرسنلی) ----------
    ws_detail = wb.create_sheet("جزئیات سفارشات")
    ws_detail['A1'] = f"لیست سفارش غذا - {today.strftime('%Y/%m/%d')}"
    ws_detail['A1'].font = Font(size=14, bold=True)
    ws_detail.merge_cells('A1:E1')   # ۵ ستون شد

    # هدر جدول (新增 ستون کد پرسنلی)
    row = 3
    headers_detail = ["ساختمان", "دپارتمان", "نام عضو", "کد پرسنلی", "وضعیت"]
    for col, h in enumerate(headers_detail, 1):
        cell = ws_detail.cell(row=row, column=col, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="E0F2E9", end_color="E0F2E9", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    row += 1

    # پر کردن داده‌های جزئیات
    for item in members_data:
        # ردیف گروه (ساختمان + دپارتمان) – ستون‌های 1 و 2 پر می‌شوند، بقیه خالی یا یکسان‌سازی
        cell_loc = ws_detail.cell(row=row, column=1, value=item['location'])
        cell_dept = ws_detail.cell(row=row, column=2, value=item['department'])
        for cell in [cell_loc, cell_dept]:
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
        # یکسان‌سازی رنگ برای ستون‌های 3 تا 5 در ردیف گروه
        for col in [3, 4, 5]:
            cell = ws_detail.cell(row=row, column=col)
            cell.fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
        row += 1

        # ردیف‌های اعضا
        for member in item['members']:
            ws_detail.cell(row=row, column=1, value="")   # ساختمان خالی
            ws_detail.cell(row=row, column=2, value="")   # دپارتمان خالی
            ws_detail.cell(row=row, column=3, value=member['name'])
            # ✅ نمایش کد پرسنلی (در صورت None یا خالی، خط تیره)
            ws_detail.cell(row=row, column=4, value=member['code'] if member['code'] is not None else "-")
            # وضعیت غذا
            if member['choice_food'] == 1:
                status = menu1_name or "منوی ۱"
            elif member['choice_food'] == 2:
                status = menu2_name or "منوی ۲"
            else:
                status = "حضور ندارد"
            ws_detail.cell(row=row, column=5, value=status)
            row += 1

    # تنظیم عرض ستون‌ها
    for ws in [ws_summary, ws_detail]:
        # برای برگه خلاصه عرض 4 ستون، برای جزئیات 5 ستون
        max_col = 5 if ws == ws_detail else 4
        for col in range(1, max_col+1):
            ws.column_dimensions[get_column_letter(col)].width = 22

    # ========== خروجی نهایی ==========
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    filename = f"گزارش_غذا_{today.strftime('%Y%m%d')}.xlsx"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response



def get_week_dates(target_date=None):
    """بازگرداندن لیستی از ۵ تاریخ (شنبه تا پنجشنبه) برای هفته‌ای که target_date در آن قرار دارد."""
    if target_date is None:
        target_date = jdatetime.date.today()
    # روز هفته در jdatetime: 0=شنبه, 1=یک‌شنبه, ..., 6=جمعه
    start_week = target_date - timedelta(days=target_date.weekday())  # شروع هفته = شنبه
    week_dates = []
    for i in range(6):  # شنبه (0) تا پنجشنبه (4)
        week_dates.append(start_week + timedelta(days=i))
    return week_dates



def get_month_end_date(start_date):
    year = start_date.year
    month = start_date.month
    if month in [1, 2, 3, 4, 5, 6]:
        month_days = 31
    elif month in [7, 8, 9, 10, 11]:
        month_days = 30
    else:
        if (year % 33 in [1, 5, 9, 13, 17, 22, 26, 30]) or (year % 4 == 0 and year % 100 != 0):
            month_days = 30
        else:
            month_days = 29
    return jdatetime.date(year, month, month_days)


def get_dates_until_month_end(start_date=None):
    """بازگرداندن لیست تاریخ‌ها از start_date تا آخر ماه جاری، به جز جمعه‌ها"""
    if start_date is None:
        start_date = jdatetime.date.today()
    end_date = get_month_end_date(start_date)
    date_list = []
    current = start_date
    while current <= end_date:
        # جمعه = 6 (در jdatetime: شنبه 0، یکشنبه 1، ...، جمعه 6)
        if current.weekday() != 6:
            date_list.append(current)
        current += timedelta(days=1)
    return date_list


def get_next_month_menu_dates(menu_dict, start_date=None):
    if start_date is None:
        start_date = jdatetime.date.today()
    if start_date != get_month_end_date(start_date):
        return []
    if start_date.month == 12:
        next_year, next_month = start_date.year + 1, 1
    else:
        next_year, next_month = start_date.year, start_date.month + 1
    return sorted(
        date for date in menu_dict
        if date.year == next_year and date.month == next_month and date.weekday() != 6
    )






def to_gregorian_date(jalali_date):
    if hasattr(jalali_date, 'togregorian'):
        return jalali_date.togregorian()
    if hasattr(jalali_date, 'to_gregorian'):
        return jalali_date.to_gregorian()
    return jalali_date.todate()


def is_locked_day(order, date, settings_obj):
    if not settings_obj.lock_menu_after_order:
        return False
    if date.weekday() == 5:
        return False
    if order is None:
        return False
    return order.menu_locked or order.choice_food in (1, 2)


def is_order_admin(user):
    return user.is_superuser or user.is_staff


def get_member_month_history(member, reference_date=None):
    """تاریخچه انتخاب غذای کارمند در ماه شمسی جاری، بر اساس همه منوهای منتشرشده."""
    if reference_date is None:
        reference_date = jdatetime.date.today()
    year_month = reference_date.strftime("%Y%m")
    menu_dict = {}
    for week_number in range(1, 5):
        weekly_menu, _, _, _ = load_weekly_menu_dict(week_number, year_month)
        menu_dict.update(weekly_menu)

    menu_dates = sorted(d for d in menu_dict if d.weekday() != 6)
    if not menu_dates:
        return [], {'menu1': 0, 'menu2': 0, 'absent': 0, 'unregistered': 0}

    orders = Order.objects.filter(
        member=member,
        date__in=[d.togregorian() for d in menu_dates],
    )
    order_dict = {order.date: order.choice_food for order in orders}
    surveys = Survey.objects.filter(
        member=member,
        date__in=[d.togregorian() for d in menu_dates],
    )
    survey_dict = {survey.date: survey for survey in surveys}
    rating_styles = {
        'good': {'label': 'خوب', 'icon': '👍', 'class': 'good'},
        'average': {'label': 'متوسط', 'icon': '😐', 'class': 'average'},
        'bad': {'label': 'ضعیف', 'icon': '👎', 'class': 'bad'},
    }
    history = []
    stats = {'menu1': 0, 'menu2': 0, 'absent': 0, 'unregistered': 0}
    for d in menu_dates:
        menu1, menu2 = menu_dict[d]
        choice = order_dict.get(d.togregorian())
        survey = survey_dict.get(d.togregorian())
        rating = rating_styles.get(survey.rating) if survey else None
        history.append({
            'date': d,
            'date_display': d.strftime("%Y/%m/%d"),
            'weekday': get_weekday_name(d.weekday()),
            'choice': choice,
            'menu1': menu1 or 'ثبت نشده',
            'menu2': menu2 or 'ثبت نشده',
            'menu1_selected': choice == 1,
            'menu2_selected': choice == 2,
            'menu1_unselected': choice == 2,
            'menu2_unselected': choice == 1,
            'is_absent': choice == 0,
            'is_unregistered': choice is None,
            'rating': rating,
        })
        if choice in (1, 2):
            stats[f'menu{choice}'] += 1
        elif choice == 0:
            stats['absent'] += 1
        else:
            stats['unregistered'] += 1
    return history, stats


def weekly_order(request):
    if not request.user.is_authenticated:
        return redirect("login")

    user = request.user
    settings_obj = get_order_settings()

    # دریافت کمپین فعال
    campaign = get_active_campaign()

    # انتخاب عضو (برای ادمین از selector، برای کاربر عادی از username)
    selected_member = None
    selected_dept_id = None
    all_members = []
    departments = []

    if is_order_admin(user):
        # مدیریت انتخاب دپارتمان
        if request.method == 'POST':
            selected_dept_id = request.POST.get('department_id')
        elif request.method == 'GET':
            selected_dept_id = request.GET.get('department_id')
        if selected_dept_id and selected_dept_id.isdigit():
            selected_dept_id = int(selected_dept_id)

        # انتخاب عضو
        if request.method == 'POST':
            member_id = request.POST.get('member_id')
            if member_id and member_id.isdigit():
                try:
                    selected_member = Member.objects.get(id=member_id)
                except Member.DoesNotExist:
                    pass
        else:
            member_id = request.GET.get('member_id')
            if member_id and member_id.isdigit():
                try:
                    selected_member = Member.objects.get(id=member_id)
                except Member.DoesNotExist:
                    pass

        if selected_dept_id:
            all_members = Member.objects.filter(department_id=selected_dept_id).select_related('department')
        else:
            all_members = Member.objects.select_related('department').all()
        departments = Department.objects.all()
    else:
        # کاربر دپارتمان: انتخاب از بین اعضای دپارتمان خود
        if not user.department_id:
            return render(request, 'weekly_order.html', {
                'is_admin': False,
                'error': 'برای این کاربر دپارتمانی تعریف نشده است. با مدیر سامانه تماس بگیرید.',
                'today_display': jdatetime.date.today().strftime("%Y/%m/%d"),
            })
        all_members = Member.objects.filter(department_id=user.department_id).select_related('department')
        selected_member = get_member_for_user(user, get_request_member_id(request))

    # اگر کارمندی انتخاب نشده، نمایش selector
    if not selected_member:
        context = {
            'is_admin': is_order_admin(user),
            'departments': departments,
            'all_members': all_members,
            'selected_dept_id': selected_dept_id,
            'selected_member': None,
            'today_display': jdatetime.date.today().strftime("%Y/%m/%d"),
        }
        return render(request, 'weekly_order.html', context)

    # ========== سیستم کمپین برای کارمند انتخاب شده ==========
    if not campaign:
        context = {
            'is_admin': is_order_admin(user),
            'departments': departments,
            'all_members': all_members,
            'selected_dept_id': selected_dept_id,
            'selected_member': selected_member,
            'campaign': None,
            'survey_data': [],
            'next_week_data': [],
            'error': 'هیچ کمپین فعالی وجود ندارد',
        }
        return render(request, 'weekly_order.html', context)

    # بررسی نظرسنجی
    survey_completed = check_survey_completed(selected_member, campaign)

    # دریافت منو هفته جاری
    current_menu_dict = get_campaign_menu_dict(campaign)
    current_week_dates = get_week_dates_for_campaign(campaign)

    # کمپین هفته بعد می‌تواند کترینگ متفاوتی داشته باشد.
    next_campaign = get_next_campaign(campaign)

    next_menu_dict = {}
    next_week_dates = []
    if next_campaign:
        next_menu_dict = get_campaign_menu_dict(next_campaign)
        next_week_dates = get_week_dates_for_campaign(next_campaign)

    # هفته جاری - برای نظرسنجی (تنها اگر نظرسنجی تکمیل نشده)
    survey_data = []
    if campaign:  # کارت‌های ثبت‌شده هم (به‌صورت قفل) نمایش داده می‌شوند
        member_orders = get_survey_orders(selected_member, campaign)

        existing_surveys = Survey.objects.filter(
            member=selected_member,
            campaign=campaign,
            date__in=[o.date for o in member_orders]
        )
        survey_dict = {s.date: s for s in existing_surveys}

        for order in member_orders:
            jd = jdatetime.date.fromgregorian(date=order.date)
            menu1, menu2 = current_menu_dict.get(jd, ("", ""))
            menu_name = menu1 if order.choice_food == 1 else menu2 if order.choice_food == 2 else "حضور ندارد"
            existing_survey = survey_dict.get(order.date)

            survey_data.append({
                'date': jd,
                'weekday': get_weekday_name(jd.weekday()),
                'date_display': jd.strftime("%Y/%m/%d"),
                'date_str': jd.strftime("%Y%m%d"),
                'menu_choice': order.get_choice_food_display(),
                'menu_name': menu_name,
                'existing_rating': existing_survey.rating if existing_survey else None,
                'is_completed': existing_survey is not None,
            })

    # هفته جاری: ادمین می‌تواند انتخاب روزانه کارمند را اصلاح کند.
    current_week_data = []
    current_order_dict = {}
    if is_order_admin(user) and current_week_dates:
        current_gregorian_dates = [d.togregorian() for d in current_week_dates]
        current_orders = Order.objects.filter(member=selected_member, date__in=current_gregorian_dates)
        current_order_dict = {o.date: o for o in current_orders}
        for d in current_week_dates:
            menu1, menu2 = current_menu_dict.get(d, ("", ""))
            existing_order = current_order_dict.get(d.togregorian())
            current_week_data.append({
                'date': d,
                'weekday': get_weekday_name(d.weekday()),
                'weekday_num': d.weekday(),
                'date_display': d.strftime("%Y/%m/%d"),
                'date_str': d.strftime("%Y%m%d"),
                'menu1': menu1 if menu1 else "ثبت نشده",
                'menu2': menu2 if menu2 else "ثبت نشده",
                'existing_choice': existing_order.choice_food if existing_order else None,
            })

    if request.method == 'POST' and is_order_admin(user) and request.POST.get('edit_current_week') == '1':
        for day_info in current_week_data:
            date_obj = day_info['date']
            choice_value = request.POST.get(f"current_choice_{day_info['date_str']}")
            if choice_value not in ('0', '1', '2'):
                continue
            # ادمین برخلاف کارمند محدود به قفل سفارش نیست و می‌تواند اشتباه را اصلاح کند.
            Order.objects.update_or_create(
                member=selected_member,
                date=date_obj.togregorian(),
                defaults={'choice_food': int(choice_value), 'menu_locked': True},
            )
        params = []
        if selected_dept_id:
            params.append(f'department_id={selected_dept_id}')
        params.append(f'member_id={selected_member.id}')
        params.append('orders_saved=1')
        return redirect(reverse('weekly_order') + '?' + '&'.join(params))

    # هفته بعد - برای سفارش
    next_week_data = []
    next_order_dict = {}
    if next_campaign and next_week_dates:
        next_gregorian_dates = [d.togregorian() for d in next_week_dates]
        next_existing_orders = Order.objects.filter(member=selected_member, date__in=next_gregorian_dates)
        next_order_dict = {o.date: o for o in next_existing_orders}

        for d in next_week_dates:
            menu1, menu2 = next_menu_dict.get(d, ("", ""))
            existing_order = next_order_dict.get(d.togregorian())

            next_week_data.append({
                'date': d,
                'weekday': get_weekday_name(d.weekday()),
                'weekday_num': d.weekday(),
                'date_display': d.strftime("%Y/%m/%d"),
                'date_str': d.strftime("%Y%m%d"),
                'menu1': menu1 if menu1 else "—",
                'menu2': menu2 if menu2 else "—",
                'existing_choice': existing_order.choice_food if existing_order else None,
                'is_locked': is_locked_day(existing_order, d, settings_obj),
            })

    # پردازش POST ثبت سفارش برای هفته بعد (اصلاح هفته جاری توسط ادمین جداگانه است)
    if request.method == 'POST' and next_campaign and request.POST.get('submit_next_week') == '1':
        for day_info in next_week_data:
            date_obj = day_info['date']
            choice_key = f'choice_{date_obj.strftime("%Y%m%d")}'
            choice_value = request.POST.get(choice_key)
            if choice_value is None or not choice_value.isdigit():
                continue
            choice_int = int(choice_value)
            existing_order = next_order_dict.get(date_obj.togregorian())
            locked = is_locked_day(existing_order, date_obj, settings_obj)
            if locked and choice_int != 0:
                continue
            greg_date = to_gregorian_date(date_obj)
            Order.objects.update_or_create(
                member=selected_member,
                date=greg_date,
                defaults={
                    'choice_food': choice_int,
                    'menu_locked': locked or choice_int in (1, 2),
                }
            )
        # حفظ پارامترهای selector برای ادمین
        params = []
        if is_order_admin(user) and selected_dept_id:
            params.append(f'department_id={selected_dept_id}')
        params.append(f'member_id={selected_member.id}')
        return redirect(reverse('weekly_order') + '?' + '&'.join(params))

    context = {
        'is_admin': is_order_admin(user),
        'departments': departments,
        'all_members': all_members,
        'selected_dept_id': selected_dept_id,
        'selected_member': selected_member,
        'campaign': campaign,
        'campaign_start_display': jdatetime.date.fromgregorian(date=campaign.start_date).strftime("%Y/%m/%d"),
        'campaign_end_display': jdatetime.date.fromgregorian(date=campaign.end_date).strftime("%Y/%m/%d"),
        'today_display': jdatetime.date.today().strftime("%Y/%m/%d"),
        'survey_data': survey_data,
        'current_week_data': current_week_data,
        'orders_saved': request.GET.get('orders_saved') == '1',
        'next_week_data': next_week_data,
        'next_campaign': next_campaign,
        'survey_completed': survey_completed,
        'RATING_CHOICES': Survey.RATING_CHOICES,
        'show_selection_warning': settings_obj.show_selection_warning,
    }
    month_history, month_stats = get_member_month_history(selected_member)
    context.update({
        'month_history': month_history,
        'month_stats': month_stats,
        'month_title': jdatetime.date.today().strftime('%Y/%m'),
    })
    return render(request, 'weekly_order.html', context)


def admin_week_menu(request):
    """مشاهده و اصلاح همه غذاهای هفته جاری توسط ادمین یا سوپریوزر."""
    if not request.user.is_authenticated:
        return redirect("login")
    if not is_order_admin(request.user):
        return redirect("weekly_order")

    campaign = get_current_week_campaign()
    if campaign is None:
        return render(request, 'admin_week_menu.html', {
            'campaign': None,
            'error': 'برای هفته جاری کمپینی ثبت نشده است.',
        })

    menu_path = get_campaign_menu_path(campaign)
    menu_dict = get_campaign_menu_dict(campaign)
    week_dates = sorted(menu_dict)
    menu_days = []
    for d in week_dates:
        menu1, menu2 = menu_dict[d]
        menu_days.append({
            'date': d,
            'date_str': d.strftime("%Y%m%d"),
            'date_display': d.strftime("%Y/%m/%d"),
            'weekday': get_weekday_name(d.weekday()),
            'is_thursday': d.weekday() == 5,
            'menu1': menu1,
            'menu2': menu2,
        })

    if request.method == 'POST':
        if not os.path.exists(menu_path):
            return render(request, 'admin_week_menu.html', {
                'campaign': campaign,
                'menu_days': menu_days,
                'error': 'فایل اکسل منوی این هفته پیدا نشد.',
            })
        rows = []
        for day in menu_days:
            rows.append({
                'date_str': day['date_str'],
                'menu1': request.POST.get(f"menu1_{day['date_str']}", '').strip(),
                'menu2': request.POST.get(f"menu2_{day['date_str']}", '').strip(),
            })
        if save_campaign_menu(campaign, rows):
            # بعد از ذخیره، نام‌های جدید در سفارش، نظرسنجی و گزارش‌ها استفاده می‌شوند.
            return redirect(reverse('admin_week_menu') + '?saved=1')
        return render(request, 'admin_week_menu.html', {
            'campaign': campaign,
            'menu_days': menu_days,
            'error': 'ذخیره فایل منو انجام نشد؛ ساختار ستون‌های اکسل را بررسی کنید.',
        })

    return render(request, 'admin_week_menu.html', {
        'campaign': campaign,
        'campaign_start_display': jdatetime.date.fromgregorian(date=campaign.start_date).strftime("%Y/%m/%d"),
        'campaign_end_display': jdatetime.date.fromgregorian(date=campaign.end_date).strftime("%Y/%m/%d"),
        'menu_days': menu_days,
        'menu_file': os.path.basename(menu_path),
        'saved': request.GET.get('saved') == '1',
    })


def aggregate_report(request):
    if not request.user.is_authenticated:
        return redirect("login")

    # تاریخ امروز شمسی
    today = jdatetime.date.today()
    # تبدیل به میلادی برای جستجو در جدول Order
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # دریافت سفارش‌های امروز
    orders = Order.objects.filter(date=today_greg).select_related('member__department__location')

    # ساختار داده: location -> {department: {'menu1': count, 'menu2': count}}
    locations_data = {}
    location_totals = {}  # جمع هر ساختمان

    for order in orders:
        member = order.member
        dept = member.department
        if not dept:
            continue
        location = dept.location
        location_name = location.name if location else "بدون ساختمان"
        dept_name = dept.name
        choice = order.choice_food  # 0,1,2

        # مقداردهی اولیه
        if location_name not in locations_data:
            locations_data[location_name] = {}
            location_totals[location_name] = {'menu1': 0, 'menu2': 0}
        if dept_name not in locations_data[location_name]:
            locations_data[location_name][dept_name] = {'menu1': 0, 'menu2': 0}

        # اضافه کردن آمار
        if choice == 1:
            locations_data[location_name][dept_name]['menu1'] += 1
            location_totals[location_name]['menu1'] += 1
        elif choice == 2:
            locations_data[location_name][dept_name]['menu2'] += 1
            location_totals[location_name]['menu2'] += 1
        # غایب (0) را نادیده می‌گیریم

    # تبدیل به لیست برای نمایش در قالب
    report_data = []
    grand_total_menu1 = 0
    grand_total_menu2 = 0

    for location_name, depts_dict in locations_data.items():
        loc_row = {
            'location': location_name,
            'departments': [],
            'loc_menu1': location_totals[location_name]['menu1'],
            'loc_menu2': location_totals[location_name]['menu2'],
        }
        for dept_name, counts in depts_dict.items():
            loc_row['departments'].append({
                'name': dept_name,
                'menu1': counts['menu1'],
                'menu2': counts['menu2'],
                'total': counts['menu1'] + counts['menu2'],
            })
            grand_total_menu1 += counts['menu1']
            grand_total_menu2 += counts['menu2']
        report_data.append(loc_row)

    # دریافت منوی امروز برای نمایش (از منوی هفته کمپین)
    menu1_name, menu2_name = get_daily_menu_names(today)
    if not menu1_name:
        menu1_name = "ثبت نشده"
    if not menu2_name:
        menu2_name = "ثبت نشده"

    context = {
        'report_data': report_data,
        'grand_total_menu1': grand_total_menu1,
        'grand_total_menu2': grand_total_menu2,
        'grand_total_all': grand_total_menu1 + grand_total_menu2,
        'today_date': today.strftime("%Y/%m/%d"),
        'weekday_str': get_weekday_name(today.weekday()),
        'menu1_name': menu1_name,
        'menu2_name': menu2_name,
    }
    return render(request, 'aggregate_report.html', context)



def weekly_aggregate_report(request):
    if not request.user.is_authenticated:
        return redirect("login")

    today = jdatetime.date.today()
    # شروع هفته: شنبه
    start_of_week = today - timedelta(days=today.weekday())
    # پایان هفته: پنج‌شنبه = شنبه + 5 روز
    end_of_week = start_of_week + timedelta(days=5)

    def to_gregorian(d):
        if hasattr(d, 'togregorian'):
            return d.togregorian()
        elif hasattr(d, 'to_gregorian'):
            return d.to_gregorian()
        else:
            return d.todate()

    start_greg = to_gregorian(start_of_week)
    end_greg = to_gregorian(end_of_week)

    # دریافت سفارش‌های هفته (شنبه تا پنج‌شنبه)
    orders = Order.objects.filter(date__range=[start_greg, end_greg]).select_related('member__department__location')

    # ساختار داده برای گزارش
    locations_data = {}
    location_totals = {}

    for order in orders:
        member = order.member
        dept = member.department
        if not dept:
            continue
        location = dept.location
        location_name = location.name if location else "بدون ساختمان"
        dept_name = dept.name
        choice = order.choice_food

        if location_name not in locations_data:
            locations_data[location_name] = {}
            location_totals[location_name] = {'menu1': 0, 'menu2': 0}
        if dept_name not in locations_data[location_name]:
            locations_data[location_name][dept_name] = {'menu1': 0, 'menu2': 0}

        if choice == 1:
            locations_data[location_name][dept_name]['menu1'] += 1
            location_totals[location_name]['menu1'] += 1
        elif choice == 2:
            locations_data[location_name][dept_name]['menu2'] += 1
            location_totals[location_name]['menu2'] += 1

    report_data = []
    grand_total_menu1 = 0
    grand_total_menu2 = 0

    for location_name, depts_dict in locations_data.items():
        loc_row = {
            'location': location_name,
            'departments': [],
            'loc_menu1': location_totals[location_name]['menu1'],
            'loc_menu2': location_totals[location_name]['menu2'],
        }
        for dept_name, counts in depts_dict.items():
            loc_row['departments'].append({
                'name': dept_name,
                'menu1': counts['menu1'],
                'menu2': counts['menu2'],
                'total': counts['menu1'] + counts['menu2'],
            })
            grand_total_menu1 += counts['menu1']
            grand_total_menu2 += counts['menu2']
        report_data.append(loc_row)

        # دریافت نام واقعی غذاهای هفته از منوی کمپین (با fallback به منوی عمومی)
    week_start_greg = start_of_week.togregorian()
    week_end_greg = end_of_week.togregorian()
    campaign = Campaign.objects.filter(
        is_active=True,
        start_date__lte=week_end_greg,
        end_date__gte=week_start_greg,
    ).select_related('catering').first()
    campaign_menu_dict = get_campaign_menu_dict(campaign) if campaign else {}
    menu_dict, _ = load_menu_dict()
    week_days = []
    for i in range(6):   # 0 تا 5 = شنبه تا پنج‌شنبه
        d = start_of_week + timedelta(days=i)
        menu1, menu2 = campaign_menu_dict.get(d) or menu_dict.get(d) or ("ثبت نشده", "ثبت نشده")
        week_days.append({
            'date': d.strftime("%Y/%m/%d"),
            'weekday': get_weekday_name(d.weekday()),
            'menu1': menu1,
            'menu2': menu2,
        })

    context = {
        'report_data': report_data,
        'grand_total_menu1': grand_total_menu1,
        'grand_total_menu2': grand_total_menu2,
        'grand_total_all': grand_total_menu1 + grand_total_menu2,
        'week_start': start_of_week.strftime("%Y/%m/%d"),
        'week_end': end_of_week.strftime("%Y/%m/%d"),
        'week_days': week_days,
    }
    return render(request, 'weekly_aggregate_report.html', context)



@staff_member_required  # فقط ادمین‌ها (superuser) می‌توانند ببینند
def date_range_aggregate_report(request):
    # دریافت تاریخ‌های شمسی از GET
    start_date_str = request.GET.get('start_date', '')
    end_date_str = request.GET.get('end_date', '')
    
    # اگر تاریخ وارد نشده، پیش‌فرض هفته جاری
    today = jdatetime.date.today()
    if not start_date_str:
        start_date = today - timedelta(days=today.weekday())  # شنبه این هفته
    else:
        start_date = parse_persian_date(start_date_str)
        if start_date is None:
            return HttpResponseBadRequest("تاریخ شروع نامعتبر است")
    
    if not end_date_str:
        end_date = start_date + timedelta(days=5)  # پنجشنبه همان هفته
    else:
        end_date = parse_persian_date(end_date_str)
        if end_date is None:
            return HttpResponseBadRequest("تاریخ پایان نامعتبر است")
    
    if start_date > end_date:
        return HttpResponseBadRequest("تاریخ شروع نباید از تاریخ پایان بزرگتر باشد")
    
    # تبدیل به میلادی برای جستجو در Order
    def to_gregorian(d):
        if hasattr(d, 'togregorian'):
            return d.togregorian()
        elif hasattr(d, 'to_gregorian'):
            return d.to_gregorian()
        else:
            return d.todate()
    
    start_greg = to_gregorian(start_date)
    end_greg = to_gregorian(end_date)
    
    # دریافت سفارش‌های بازه
    orders = Order.objects.filter(date__range=[start_greg, end_greg]).select_related('member__department__location')
    
    # ساختار داده: location -> {department: {'menu1': count, 'menu2': count}}
    locations_data = {}
    location_totals = {}
    
    for order in orders:
        member = order.member
        dept = member.department
        if not dept:
            continue
        location = dept.location
        location_name = location.name if location else "بدون ساختمان"
        dept_name = dept.name
        choice = order.choice_food
        
        if location_name not in locations_data:
            locations_data[location_name] = {}
            location_totals[location_name] = {'menu1': 0, 'menu2': 0}
        if dept_name not in locations_data[location_name]:
            locations_data[location_name][dept_name] = {'menu1': 0, 'menu2': 0}
        
        if choice == 1:
            locations_data[location_name][dept_name]['menu1'] += 1
            location_totals[location_name]['menu1'] += 1
        elif choice == 2:
            locations_data[location_name][dept_name]['menu2'] += 1
            location_totals[location_name]['menu2'] += 1
    
    report_data = []
    grand_total_menu1 = 0
    grand_total_menu2 = 0
    
    for location_name, depts_dict in locations_data.items():
        loc_row = {
            'location': location_name,
            'departments': [],
            'loc_menu1': location_totals[location_name]['menu1'],
            'loc_menu2': location_totals[location_name]['menu2'],
        }
        for dept_name, counts in depts_dict.items():
            loc_row['departments'].append({
                'name': dept_name,
                'menu1': counts['menu1'],
                'menu2': counts['menu2'],
                'total': counts['menu1'] + counts['menu2'],
            })
            grand_total_menu1 += counts['menu1']
            grand_total_menu2 += counts['menu2']
        report_data.append(loc_row)
    
    # دریافت نام واقعی غذاهای بازه از منوی کمپین (با fallback به منوی عمومی)
    menu_dict, _ = load_menu_dict()
    campaign_menu_dict = {}
    campaign = Campaign.objects.filter(
        is_active=True,
        start_date__lte=end_greg,
        end_date__gte=start_greg,
    ).select_related('catering').first()
    if campaign:
        campaign_menu_dict = get_campaign_menu_dict(campaign)
    days_in_range = []
    current = start_date
    while current <= end_date:
        if current.weekday() != 6:  # جمعه را حذف کن
            menu1, menu2 = campaign_menu_dict.get(current) or menu_dict.get(current) or ("ثبت نشده", "ثبت نشده")
            days_in_range.append({
                'date': current.strftime("%Y/%m/%d"),
                'weekday': get_weekday_name(current.weekday()),
                'menu1': menu1,
                'menu2': menu2,
            })
        current += timedelta(days=1)
    
    context = {
        'report_data': report_data,
        'grand_total_menu1': grand_total_menu1,
        'grand_total_menu2': grand_total_menu2,
        'grand_total_all': grand_total_menu1 + grand_total_menu2,
        'start_date': start_date.strftime("%Y/%m/%d"),
        'end_date': end_date.strftime("%Y/%m/%d"),
        'days_in_range': days_in_range,
        'start_date_str': start_date_str,
        'end_date_str': end_date_str,
    }
    return render(request, 'date_range_aggregate_report.html', context)



@staff_member_required
def daily_menu_report(request):
    """گزارش تعداد سفارش‌های منوی ۱ و ۲ به تفکیک روز در بازه انتخابی"""
    
    start_date_str = request.GET.get('start_date', '')
    end_date_str = request.GET.get('end_date', '')
    
    today = jdatetime.date.today()
    if not start_date_str:
        start_date = today - timedelta(days=today.weekday())  # شنبه این هفته
        start_date_str = start_date.strftime("%Y/%m/%d")
    else:
        start_date = parse_persian_date(start_date_str)
        if start_date is None:
            return HttpResponseBadRequest("تاریخ شروع نامعتبر است")
    
    if not end_date_str:
        end_date = start_date + timedelta(days=5)  # پنجشنبه همان هفته
        end_date_str = end_date.strftime("%Y/%m/%d")
    else:
        end_date = parse_persian_date(end_date_str)
        if end_date is None:
            return HttpResponseBadRequest("تاریخ پایان نامعتبر است")
    
    if start_date > end_date:
        return HttpResponseBadRequest("تاریخ شروع نباید از تاریخ پایان بزرگتر باشد")
    
    # تبدیل به میلادی
    def to_gregorian(d):
        if hasattr(d, 'togregorian'):
            return d.togregorian()
        elif hasattr(d, 'to_gregorian'):
            return d.to_gregorian()
        else:
            return d.todate()
    
    start_greg = to_gregorian(start_date)
    end_greg = to_gregorian(end_date)
    
    # دریافت تمام سفارش‌های بازه
    orders = Order.objects.filter(date__range=[start_greg, end_greg]).select_related('member')
    
    # دریافت منوها برای نمایش نام غذاها (منوی هفته کمپین + منوی عمومی)
    menu_dict, _ = load_menu_dict()
    campaign_menu_dict = {}
    campaign = Campaign.objects.filter(
        is_active=True,
        start_date__lte=end_greg,
        end_date__gte=start_greg,
    ).select_related('catering').first()
    if campaign:
        campaign_menu_dict = get_campaign_menu_dict(campaign)
    
    # ساختار داده: روز شمسی -> {'menu1': int, 'menu2': int}
    daily_data = {}
    
    # لیست روزهای کاری (بدون جمعه) در بازه
    date_list = []
    current = start_date
    while current <= end_date:
        if current.weekday() != 6:   # حذف جمعه
            date_list.append(current)
            daily_data[current] = {'menu1': 0, 'menu2': 0}
        current += timedelta(days=1)
    
    # پردازش سفارش‌ها
    for order in orders:
        member = order.member
        if not member:
            continue
        # تبدیل تاریخ میلادی به شمسی
        jd = jdatetime.date.fromgregorian(date=order.date)
        if jd.weekday() == 6:   # احتیاطاً جمعه‌ها نادیده گرفته می‌شوند
            continue
        choice = order.choice_food
        if choice == 1:
            daily_data[jd]['menu1'] += 1
        elif choice == 2:
            daily_data[jd]['menu2'] += 1
        # انتخاب 0 (حضور ندارد) نادیده گرفته می‌شود
    
    # آماده‌سازی ردیف‌های جدول
    rows = []
    grand_total_menu1 = 0
    grand_total_menu2 = 0
    for d in date_list:
        menu1, menu2 = campaign_menu_dict.get(d) or menu_dict.get(d) or ("ثبت نشده", "ثبت نشده")
        count1 = daily_data[d]['menu1']
        count2 = daily_data[d]['menu2']
        rows.append({
            'date': d.strftime("%Y/%m/%d"),
            'weekday': get_weekday_name(d.weekday()),
            'menu1_name': menu1,
            'menu2_name': menu2,
            'count1': count1,
            'count2': count2,
            'total': count1 + count2,
        })
        grand_total_menu1 += count1
        grand_total_menu2 += count2
    
    context = {
        'rows': rows,
        'start_date': start_date.strftime("%Y/%m/%d"),
        'end_date': end_date.strftime("%Y/%m/%d"),
        'start_date_str': start_date_str,
        'end_date_str': end_date_str,
        'grand_total_menu1': grand_total_menu1,
        'grand_total_menu2': grand_total_menu2,
        'grand_total_all': grand_total_menu1 + grand_total_menu2,
        'total_days': len(rows),
    }
    return render(request, 'daily_menu_report.html', context)




