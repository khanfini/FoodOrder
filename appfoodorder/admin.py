from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib import messages
from django.utils.translation import gettext_lazy as _

from .models import Department, Member, CustomUser, Location, Floor, OrderSettings

# ------------------------------------------------------------
# CustomUser Admin
# ------------------------------------------------------------
@admin.register(CustomUser)
class CustomUserAdmin(BaseUserAdmin):
    list_display = ('username', 'department', 'is_staff', 'is_active')
    list_filter = ('is_staff', 'is_active', 'department')
    search_fields = ('username', 'department__name')
    filter_horizontal = ('groups', 'user_permissions')
    
    actions = ['activate_users', 'deactivate_users']

    def activate_users(self, request, queryset):
        updated = queryset.update(is_active=True)
        self.message_user(request, _('%(count)d کاربر با موفقیت فعال شدند.') % {'count': updated}, messages.SUCCESS)
    activate_users.short_description = _("فعال کردن کاربران انتخاب شده")

    def deactivate_users(self, request, queryset):
        updated = queryset.update(is_active=False)
        self.message_user(request, _('%(count)d کاربر با موفقیت غیرفعال شدند.') % {'count': updated}, messages.WARNING)
    deactivate_users.short_description = _("غیرفعال کردن کاربران انتخاب شده")

    fieldsets = (
        (None, {'fields': ('username', 'password')}),
        ('Personal info', {'fields': ('first_name', 'last_name', 'email')}),
        ('Permissions', {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions')}),
        ('Important dates', {'fields': ('last_login', 'date_joined')}),
        ('Department info', {'fields': ('department',)}),
    )

    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('username', 'password1', 'password2', 'department', 'is_staff', 'is_active'),
        }),
    )


# ------------------------------------------------------------
# Member Admin (نمایش کد پرسنلی)
# ------------------------------------------------------------
@admin.register(Member)
class MemberAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'department', 'choice_food')
    list_filter = ('department', 'choice_food')
    search_fields = ('name', 'code')
    ordering = ('name',)

    # برای نمایش بهتر انتخاب غذا در لیست
    def get_choice_food_display(self, obj):
        return dict(obj.choice_food).get(obj.choice_food, 'نامشخص')
    get_choice_food_display.short_description = 'انتخاب غذا'
    get_choice_food_display.admin_order_field = 'choice_food'

    # در صورت تمایل می‌توان فیلد code را در فرم ادمین اجباری یا اختیاری کرد
    # (همین حالا چون blank=True در مدل دارد، اختیاری است)


# ------------------------------------------------------------
# سایر مدل‌ها با ثبت پیش‌فرض
# ------------------------------------------------------------
@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'location', 'floor')
    search_fields = ('name',)

@admin.register(Location)
class LocationAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(Floor)
class FloorAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)


@admin.register(OrderSettings)
class OrderSettingsAdmin(admin.ModelAdmin):
    list_display = ('label', 'lock_menu_after_order', 'show_selection_warning')
    list_display_links = ('label',)

    def label(self, obj):
        return "تنظیمات سفارش غذا (برای ویرایش کلیک کنید)"

    label.short_description = "تنظیمات"

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


# from django.contrib import admin
# from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
# from .models import Department, Member, CustomUser ,Location,Floor
# from django.contrib import messages
# from django.utils.translation import gettext_lazy as _

# # Register your models here.

# @admin.register(CustomUser)
# class CustomUserAdmin(BaseUserAdmin):
#     list_display = ('username', 'department', 'is_staff', 'is_active')
#     list_filter = ('is_staff', 'is_active', 'department')
#     search_fields = ('username', 'department__name')
#     filter_horizontal = ('groups', 'user_permissions')
    
#     # تعریف اکشن‌های سفارشی
#     actions = ['activate_users', 'deactivate_users']

#     def activate_users(self, request, queryset):
#         """فعال کردن کاربران انتخاب شده"""
#         updated = queryset.update(is_active=True)
#         self.message_user(request, _('%(count)d کاربر با موفقیت فعال شدند.') % {'count': updated}, messages.SUCCESS)
#     activate_users.short_description = _("فعال کردن کاربران انتخاب شده")

#     def deactivate_users(self, request, queryset):
#         """غیرفعال کردن کاربران انتخاب شده"""
#         updated = queryset.update(is_active=False)
#         self.message_user(request, _('%(count)d کاربر با موفقیت غیرفعال شدند.') % {'count': updated}, messages.WARNING)
#     deactivate_users.short_description = _("غیرفعال کردن کاربران انتخاب شده")

#     # بقیه متدهای fieldsets و add_fieldsets مانند قبل
#     fieldsets = (
#         (None, {'fields': ('username', 'password')}),
#         ('Personal info', {'fields': ('first_name', 'last_name', 'email')}),
#         ('Permissions', {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions')}),
#         ('Important dates', {'fields': ('last_login', 'date_joined')}),
#         ('Department info', {'fields': ('department',)}),
#     )

#     add_fieldsets = (
#         (None, {
#             'classes': ('wide',),
#             'fields': ('username', 'password1', 'password2', 'department', 'is_staff', 'is_active'),
#         }),
#     )
# admin.site.register(Department)
# admin.site.register(Member)
# admin.site.register(Location)
# admin.site.register(Floor)


