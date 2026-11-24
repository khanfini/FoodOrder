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
        # فallback به فایل پیش‌فرض
        return load_menu_dict()

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


# ✅ تابع جدید: بررسی تکمیل نظرسنجی هفته قبل
def check_survey_completed(member, campaign):
    """بررسی اینکه آیا کاربر نظرسنجی هفته قبل را تکمیل کرده"""
    if not campaign or campaign.week_number == 1:
        return True  # هفته اول نیاز به نظرسنجی قبلی ندارد

    prev_week = campaign.week_number - 1
    prev_campaign = Campaign.objects.filter(
        catering=campaign.catering,
        week_number=prev_week,
        is_active=True
    ).first()

    if not prev_campaign:
        return True

    # بررسی وجود نظرسنجی برای تمام روزهای هفته قبل که کاربر غذا خورده
    prev_week_dates = get_week_dates_for_campaign(prev_campaign)
    member_orders = Order.objects.filter(
        member=member,
        date__in=[d.togregorian() for d in prev_week_dates]
    ).exclude(choice_food=0)

    if not member_orders.exists():
        return True  # کاربر هفته قبل غذا نخورده

    # بررسی تکمیل نظرسنجی برای تمام سفارشات
    survey_count = Survey.objects.filter(
        member=member,
        campaign=prev_campaign,
        date__in=[o.date for o in member_orders]
    ).count()

    return survey_count >= member_orders.count()


# ✅ تابع جدید: دریافت تاریخ‌های هفته برای کمپین
def get_week_dates_for_campaign(campaign):
    """دریافت تاریخ‌های شنبه تا پنجشنبه برای هفته کمپین"""
    start_date = jdatetime.date.fromgregorian(date=campaign.start_date)
    week_dates = []
    current = start_date
    for i in range(6):  # شنبه تا پنجشنبه
        if current.weekday() != 6:  # نه جمعه
            week_dates.append(current)
        current += timedelta(days=1)
    return week_dates


# ✅ تابع جدید: دریافت منو برای هفته کمپین
def get_campaign_menu_dict(campaign):
    """دریافت منو از فایل هفته کمپین"""
    # استخراج سال و ماه از تاریخ شروع
    start_jalali = jdatetime.date.fromgregorian(date=campaign.start_date)
    year_month = start_jalali.strftime("%Y%m")
    menu_dict, _, _, _ = load_weekly_menu_dict(campaign.week_number, year_month)
    return menu_dict


@ensure_csrf_cookie
def login(request):
    menu_today, debug_info = get_menu_for_today()
    if request.method == "POST":
        username = request.POST.get("username")
        password = request.POST.get("password")
        user = authenticate(request, username=username, password=password)
        if user is not None:
            auth_login(request, user)

            # بررسی نظرسنجی برای کاربران عادی (نه سوپریوزر)
            if not user.is_superuser:
                try:
                    member = Member.objects.get(name=user.username)
                    campaign = get_active_campaign()
                    if campaign and not check_survey_completed(member, campaign):
                        # هدایت به صفحه نظرسنجی
                        return redirect('survey')
                except Member.DoesNotExist:
                    pass

            return redirect("weekly_order")
        else:
            return render(request, "login.html", {"error": "نام کاربری یا رمز عبور اشتباه است", "menu_today": menu_today})
    return render(request, "login.html", {"menu_today": menu_today})


def logout_view(request):
    auth_logout(request)
    return redirect("login")


# ✅ ویو نظرسنجی
def survey(request):
    if not request.user.is_authenticated:
        return redirect("login")

    user = request.user
    if user.is_superuser:
        return redirect("weekly_order")

    try:
        member = Member.objects.get(name=user.username)
    except Member.DoesNotExist:
        return redirect("weekly_order")

    campaign = get_active_campaign()
    if not campaign:
        return redirect("weekly_order")

    # بررسی آیا نظرسنجی هفته قبل تکمیل شده
    if check_survey_completed(member, campaign):
        return redirect("weekly_order")

    # دریافت منو هفته جاری (برای نظرسنجی)
    current_menu_dict = get_campaign_menu_dict(campaign)
    current_week_dates = get_week_dates_for_campaign(campaign)

    # دریافت منو هفته بعد (برای سفارش)
    next_week_number = campaign.week_number + 1 if campaign.week_number < 4 else 1
    next_campaign = Campaign.objects.filter(
        catering=campaign.catering,
        week_number=next_week_number,
        is_active=True
    ).first()

    next_menu_dict = {}
    next_week_dates = []
    if next_campaign:
        next_menu_dict = get_campaign_menu_dict(next_campaign)
        next_week_dates = get_week_dates_for_campaign(next_campaign)

    # دریافت سفارشات کاربر برای هفته جاری
    gregorian_dates = [d.togregorian() for d in current_week_dates]
    member_orders = Order.objects.filter(member=member, date__in=gregorian_dates).exclude(choice_food=0)

    # بررسی نظرسنجی‌های قبلی
    existing_surveys = Survey.objects.filter(
        member=member,
        campaign=campaign,
        date__in=[o.date for o in member_orders]
    )
    survey_dict = {s.date: s for s in existing_surveys}

    # پردازش POST ثبت نظرسنجی
    if request.method == 'POST':
        all_rated = True
        for order in member_orders:
            rating_key = f'rating_{order.date.strftime("%Y%m%d")}'
            rating_value = request.POST.get(rating_key)
            if not rating_value:
                all_rated = False
                continue

            # ذخیره یا آپدیت نظرسنجی
            Survey.objects.update_or_create(
                member=member,
                campaign=campaign,
                date=order.date,
                defaults={
                    'menu_choice': order.choice_food,
                    'rating': rating_value,
                }
            )

        if all_rated and member_orders.exists():
            # هدایت به صفحه سفارش
            return redirect("weekly_order")

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
                'date_display': d.strftime("%Y/%m/%d"),
                'date_str': d.strftime("%Y%m%d"),
                'menu1': menu1 if menu1 else "—",
                'menu2': menu2 if menu2 else "—",
                'existing_choice': existing_order.choice_food if existing_order else None,
            })

    context = {
        'campaign': campaign,
        'survey_data': survey_data,
        'next_week_data': next_week_data,
        'next_campaign': next_campaign,
        'RATING_CHOICES': Survey.RATING_CHOICES,
    }
    return render(request, "survey.html", context)


# ✅ تابع جدید: دریافت منو بر اساس تاریخ دقیق
def get_menu_for_today():
    today = jdatetime.date.today()
    menu_dict, debug_info = load_menu_dict()
    # می‌توانید debug_info را به view برگردانید
    return menu_dict.get(today, (None, None)), debug_info


# ✅ تابع جدید: نام روز هفته
def get_weekday_name(weekday):
    return ["شنبه", "یکشنبه", "دوشنبه", "سه شنبه", "چهارشنبه", "پنجشنبه", "جمعه"][weekday]


# ✅ فرم ثبت سفارش (تغییر نکرده)
class OrderForm(forms.Form):
    def __init__(self, *args, **kwargs):
        user = kwargs.pop('user', None)
        super().__init__(*args, **kwargs)
        if user.is_superuser:
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

    # منوی امروز
    menu1_name, menu2_name = menu_dict.get(today, (None, None))
    if menu1_name is None:
        menu1_name = "منوی امروز ثبت نشده است"
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
    if user.is_superuser:
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

            if not user.is_superuser and member.department != user.department:
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
    if not request.user.is_superuser:
        return redirect("order")

    # تاریخ امروز شمسی و میلادی
    today = jdatetime.date.today()
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # دریافت منوی امروز از اکسل
    menu_dict, _ = load_menu_dict()
    menu1_name, menu2_name = menu_dict.get(today, ("", ""))

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
    if not request.user.is_superuser:
        return redirect("order")

    # ========== تاریخ امروز ==========
    today = jdatetime.date.today()
    if hasattr(today, 'togregorian'):
        today_greg = today.togregorian()
    elif hasattr(today, 'to_gregorian'):
        today_greg = today.to_gregorian()
    else:
        today_greg = today.todate()

    # ========== دریافت منوی امروز ==========
    menu_dict, _ = load_menu_dict()
    menu1_name, menu2_name = menu_dict.get(today, ("", ""))

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


def weekly_order(request):
    if not request.user.is_authenticated:
        return redirect("login")

    user = request.user
    settings_obj = get_order_settings()

    # دریافت کمپین فعال
    campaign = get_active_campaign()

    # برای سوپریوزر: سیستم قدیمی با منوی ماهانه
    if user.is_superuser:
        menu_dict, debug_info = load_menu_dict()
        today = jdatetime.date.today()
        date_list = get_dates_until_month_end(today)
        date_list.extend(get_next_month_menu_dates(menu_dict, today))

        week_data = []
        for d in date_list:
            menu1, menu2 = menu_dict.get(d, ("ثبت نشده", "ثبت نشده"))
            week_data.append({
                'date': d,
                'weekday': get_weekday_name(d.weekday()),
                'weekday_num': d.weekday(),
                'menu1': menu1 if menu1 else "—",
                'menu2': menu2 if menu2 else "—",
            })

        # مدیریت انتخاب دپارتمان
        selected_dept_id = None
        if request.method == 'POST':
            selected_dept_id = request.POST.get('department_id')
        elif request.method == 'GET':
            selected_dept_id = request.GET.get('department_id')
        if selected_dept_id and selected_dept_id.isdigit():
            selected_dept_id = int(selected_dept_id)

        # انتخاب عضو
        selected_member = None
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

        existing_orders = {}
        if selected_member:
            gregorian_dates = [to_gregorian_date(d) for d in date_list]
            orders = Order.objects.filter(member=selected_member, date__in=gregorian_dates)
            for o in orders:
                jd = jdatetime.date.fromgregorian(date=o.date)
                existing_orders[jd] = o

        if request.method == 'POST' and selected_member:
            for day_info in week_data:
                date_obj = day_info['date']
                choice_key = f'choice_{date_obj.strftime("%Y%m%d")}'
                choice_value = request.POST.get(choice_key)
                if choice_value is None or not choice_value.isdigit():
                    continue
                choice_int = int(choice_value)
                existing_order = existing_orders.get(date_obj)
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
            params = []
            if selected_dept_id:
                params.append(f'department_id={selected_dept_id}')
            params.append(f'member_id={selected_member.id}')
            return redirect(reverse('weekly_order') + '?' + '&'.join(params))

        for day in week_data:
            existing_order = existing_orders.get(day['date'])
            existing_choice = existing_order.choice_food if existing_order else None
            day['existing_choice'] = existing_choice
            day['date_str'] = day['date'].strftime("%Y%m%d")
            day['date_display'] = day['date'].strftime("%Y/%m/%d")
            day['is_locked'] = is_locked_day(existing_order, day['date'], settings_obj)

        if selected_dept_id:
            all_members = Member.objects.filter(department_id=selected_dept_id).select_related('department')
        else:
            all_members = Member.objects.select_related('department').all()
        departments = Department.objects.all()

        context = {
            'week_data': week_data,
            'selected_member': selected_member,
            'all_members': all_members,
            'today': jdatetime.date.today(),
            'debug_info': debug_info,
            'is_admin': True,
            'departments': departments,
            'selected_dept_id': selected_dept_id,
            'show_selection_warning': settings_obj.show_selection_warning,
            'campaign': None,
        }
        return render(request, 'weekly_order.html', context)

    # برای کاربران عادی: سیستم کمپین
    try:
        member = Member.objects.get(name=user.username)
    except Member.DoesNotExist:
        return redirect("login")

    if not campaign:
        # هیچ کمپین فعالی نیست - نمایش پیام
        context = {
            'campaign': None,
            'survey_data': [],
            'next_week_data': [],
            'error': 'هیچ کمپین فعالی وجود ندارد',
        }
        return render(request, 'weekly_order.html', context)

    # بررسی نظرسنجی
    survey_completed = check_survey_completed(member, campaign)

    # دریافت منو هفته جاری
    current_menu_dict = get_campaign_menu_dict(campaign)
    current_week_dates = get_week_dates_for_campaign(campaign)

    # دریافت منو هفته بعد
    next_week_number = campaign.week_number + 1 if campaign.week_number < 4 else 1
    next_campaign = Campaign.objects.filter(
        catering=campaign.catering,
        week_number=next_week_number,
        is_active=True
    ).first()

    next_menu_dict = {}
    next_week_dates = []
    if next_campaign:
        next_menu_dict = get_campaign_menu_dict(next_campaign)
        next_week_dates = get_week_dates_for_campaign(next_campaign)

    # هفته جاری - برای نظرسنجی (تنها اگر نظرسنجی تکمیل نشده)
    survey_data = []
    if not survey_completed:
        gregorian_dates = [d.togregorian() for d in current_week_dates]
        member_orders = Order.objects.filter(member=member, date__in=gregorian_dates).exclude(choice_food=0)

        existing_surveys = Survey.objects.filter(
            member=member,
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
            })

    # هفته بعد - برای سفارش
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
                'date_display': d.strftime("%Y/%m/%d"),
                'date_str': d.strftime("%Y%m%d"),
                'menu1': menu1 if menu1 else "—",
                'menu2': menu2 if menu2 else "—",
                'existing_choice': existing_order.choice_food if existing_order else None,
                'is_locked': is_locked_day(existing_order, d, settings_obj),
            })

    # پردازش POST ثبت سفارش برای هفته بعد
    if request.method == 'POST' and next_campaign:
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
                member=member,
                date=greg_date,
                defaults={
                    'choice_food': choice_int,
                    'menu_locked': locked or choice_int in (1, 2),
                }
            )
        return redirect('weekly_order')

    context = {
        'campaign': campaign,
        'survey_data': survey_data,
        'next_week_data': next_week_data,
        'next_campaign': next_campaign,
        'survey_completed': survey_completed,
        'RATING_CHOICES': Survey.RATING_CHOICES,
        'show_selection_warning': settings_obj.show_selection_warning,
    }
    return render(request, 'weekly_order.html', context)


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

    # دریافت منوی امروز برای نمایش (اختیاری)
    menu_dict, _ = load_menu_dict()
    menu1_name, menu2_name = menu_dict.get(today, ("ثبت نشده", "ثبت نشده"))

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

    # دریافت منوهای هفته (۶ روز: شنبه تا پنج‌شنبه)
    menu_dict, _ = load_menu_dict()
    week_days = []
    for i in range(6):   # 0 تا 5 = شنبه تا پنج‌شنبه
        d = start_of_week + timedelta(days=i)
        menu1, menu2 = menu_dict.get(d, ("ثبت نشده", "ثبت نشده"))
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
    
    # دریافت منوهای روزهای بازه (از فایل اکسل)
    menu_dict, _ = load_menu_dict()
    days_in_range = []
    current = start_date
    while current <= end_date:
        if current.weekday() != 6:  # جمعه را حذف کن
            menu1, menu2 = menu_dict.get(current, ("ثبت نشده", "ثبت نشده"))
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
    
    # دریافت منوها از فایل اکسل (برای نمایش نام منوها)
    menu_dict, _ = load_menu_dict()
    
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
        menu1, menu2 = menu_dict.get(d, ("ثبت نشده", "ثبت نشده"))
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




