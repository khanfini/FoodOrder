from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model

class Command(BaseCommand):
    help = 'Reactivate admin user'

    def handle(self, *args, **options):
        User = get_user_model()
        user = User.objects.filter(username='admin', is_active=False).first()
        if user:
            user.is_active = True
            user.save()
            self.stdout.write(self.style.SUCCESS('Admin activated'))
        else:
            self.stdout.write(self.style.WARNING('Admin not found or already active'))