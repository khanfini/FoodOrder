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

    class Meta:
        unique_together = ('member', 'date')   # هر عضو فقط یک سفارش برای هر روز

    def __str__(self):
        return f"{self.member.name} - {self.date} - {self.get_choice_food_display()}"
