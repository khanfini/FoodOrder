import os
from waitress import serve
from django.core.wsgi import get_wsgi_application

# تنظیم متغیر محیطی DJANGO_SETTINGS_MODULE (نام پروژه خود را جایگزین کنید)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'foodorder.settings')  # نام پروژه شما

application = get_wsgi_application()

if __name__ == '__main__':
    print("Starting Waitress on 0.0.0.0:8082 ...")
    serve(application, host='0.0.0.0', port=8082)