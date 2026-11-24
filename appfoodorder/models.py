from django.db import models
from django.contrib.auth.models import AbstractUser

class Location(models.Model):
    name = models.CharField(max_length=100)

    def __str__(self):
        return self.name
    
class Floor(models.Model):
    name = models.CharField(max_length=100)

    def __str__(self):
        return self.name

class Department(models.Model):
    name = models.CharField(max_length=100)
    location = models.ForeignKey(Location,on_delete=models.CASCADE,null=True,blank=True)
    floor = models.ForeignKey(Floor,on_delete=models.CASCADE,null=True,blank=True)

    def __str__(self):
        return self.name
    
class CustomUser(AbstractUser):
    department = models.ForeignKey(
        Department,
        on_delete=models.SET_NULL,
        null=True,
        blank=True
    )

    REQUIRED_FIELDS = []  


class Member(models.Model):
    name = models.CharField(max_length=100)
    code = models.IntegerField(null = True,blank=True)
    # location = models.ForeignKey(Location,on_delete=models.CASCADE)
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    choice_food = models.IntegerField(
        choices=[(1, "منوی 1"), (2, "منوی 2"), (3, "حضور ندارد")],
        default=1
    )

    def __str__(self):
        return self.name


class Order(models.Model):
    CHOICES = [(0, 'حضور ندارد'), (1, 'منوی ۱'), (2, 'منوی ۲')]

    member = models.ForeignKey(Member, on_delete=models.CASCADE)
    date = models.DateField()      # تاریخ سفارش (مثلاً 1405/02/07)
    choice_food = models.IntegerField(choices=CHOICES, default=0)
    menu_locked = models.BooleanField(default=False)

    class Meta:
        unique_together = ('member', 'date')   # هر عضو فقط یک سفارش برای هر روز

    def __str__(self):
        return f"{self.member.name} - {self.date} - {self.get_choice_food_display()}"


class OrderSettings(models.Model):
    lock_menu_after_order = models.BooleanField(default=True)
    show_selection_warning = models.BooleanField(default=True)

    def __str__(self):
        return "تنظیمات سفارش غذا"


def get_order_settings():
    settings_obj, _ = OrderSettings.objects.get_or_create(pk=1)
    return settings_obj


class Catering(models.Model):
    name = models.CharField(max_length=100)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = "کترینگ"
        verbose_name_plural = "کترینگ‌ها"


class Campaign(models.Model):
    catering = models.ForeignKey(Catering, on_delete=models.CASCADE, related_name='campaigns')
    week_number = models.IntegerField(choices=[(1, 'هفته ۱'), (2, 'هفته ۲'), (3, 'هفته ۳'), (4, 'هفته ۴')])
    start_date = models.DateField()
    end_date = models.DateField()
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return f"{self.catering.name} - هفته {self.week_number}"

    class Meta:
        verbose_name = "کمپین"
        verbose_name_plural = "کمپین‌ها"
        unique_together = ('catering', 'week_number')


class Survey(models.Model):
    RATING_CHOICES = [
        ('good', 'خوب'),
        ('average', 'متوسط'),
        ('bad', 'ضعیف'),
    ]

    member = models.ForeignKey(Member, on_delete=models.CASCADE, related_name='surveys')
    campaign = models.ForeignKey(Campaign, on_delete=models.CASCADE, related_name='surveys')
    date = models.DateField()
    menu_choice = models.IntegerField(choices=Order.CHOICES)
    rating = models.CharField(max_length=10, choices=RATING_CHOICES)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.member.name} - {self.campaign} - {self.date} - {self.get_rating_display()}"

    class Meta:
        verbose_name = "نظرسنجی"
        verbose_name_plural = "نظرسنجی‌ها"
        unique_together = ('member', 'campaign', 'date')
