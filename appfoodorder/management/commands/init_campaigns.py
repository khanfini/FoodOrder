import jdatetime
import os
import re
import pandas as pd
from django.conf import settings
from django.core.management.base import BaseCommand
from appfoodorder.models import Catering, Campaign


class Command(BaseCommand):
    help = 'Initialize food survey campaigns for 4 weeks'

    def handle(self, *args, **options):
        today = jdatetime.date.today()
        self.stdout.write(f'Today: {today}')

        # Create 4 caterings
        catering_names = ['Catering 1', 'Catering 2', 'Catering 3', 'Catering 4']
        caterings = []
        for i, name in enumerate(catering_names, 1):
            catering, created = Catering.objects.get_or_create(name=name, defaults={'is_active': True})
            caterings.append(catering)
            status = 'Created' if created else 'Exists'
            self.stdout.write(f'{status}: Catering {i}')

        for week_num in range(1, 5):
            filename = f'Menu_{today.strftime("%Y%m")}-w{week_num}.xlsx'
            path = os.path.join(settings.BASE_DIR, 'media', filename)
            if not os.path.exists(path):
                self.stdout.write(f'Skipped: {filename} not found')
                continue

            dates = []
            for value in pd.read_excel(path)['Date']:
                parts = re.split(r'[/\-]', str(value).strip())
                if len(parts) == 3:
                    dates.append(jdatetime.date(*map(int, parts)))
            if not dates:
                self.stdout.write(f'Skipped: {filename} has no valid dates')
                continue

            catering = caterings[week_num - 1]
            campaign, created = Campaign.objects.update_or_create(
                catering=catering,
                week_number=week_num,
                defaults={
                    'start_date': min(dates).togregorian(),
                    'end_date': max(dates).togregorian(),
                    'is_active': True,
                }
            )

            status = 'Created' if created else 'Exists'
            self.stdout.write(f'{status}: Week {week_num} from {filename}')

        self.stdout.write(self.style.SUCCESS('Campaign initialization complete!'))
