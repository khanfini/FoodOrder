"""
URL configuration for foodorder project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from appfoodorder.views import login,logout_view,order,printable_list,weekly_order,aggregate_report,weekly_aggregate_report,date_range_aggregate_report,daily_menu_report,export_printable_list_excel


urlpatterns = [
    path('admin/', admin.site.urls),
    path('',login,name="login"),
    path('logout/', logout_view, name="logout"),
    path('order/',order,name="order"),
    path('printable/',printable_list,name = 'printable_list'),
    path('weekly-order/', weekly_order, name='weekly_order'), 
    path('aggregate-report/', aggregate_report, name='aggregate_report'),
    path('weekly-aggregate-report/',weekly_aggregate_report, name='weekly_aggregate_report'),
    path('date-range-aggregate-report/',date_range_aggregate_report, name='date_range_aggregate_report'),
    path('daily-menu-report/',daily_menu_report, name='daily_menu_report'),
    path('printable-list/export-excel/',export_printable_list_excel, name='export_printable_list_excel'),
]
