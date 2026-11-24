import jdatetime
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

        # Calculate week start dates for current month
        year = today.year
        month = today.month

        # Week 1: Days 1-7 (Saturday to Thursday)
        # Week 2: Days 8-14
        # Week 3: Days 15-21
        # Week 4: Days 22-end

        for week_num in range(1, 5):
            # Find first Saturday of the week
            start_day = (week_num - 1) * 7 + 1
            try:
                week_start = jdatetime.date(year, month, start_day)
            except:
                continue

            # Adjust to Saturday if needed
            days_until_saturday = (0 - week_start.weekday()) % 7
            if days_until_saturday == 0 and week_start.weekday() != 0:
                days_until_saturday = 7
            week_start = week_start + jdatetime.timedelta(days=days_until_saturday)

            # End date is Thursday (5 days after Saturday)
            week_end = week_start + jdatetime.timedelta(days=5)

            catering = caterings[week_num - 1]

            # Check if campaign exists
            campaign, created = Campaign.objects.get_or_create(
                catering=catering,
                week_number=week_num,
                defaults={
                    'start_date': week_start.togregorian(),
                    'end_date': week_end.togregorian(),
                    'is_active': True,
                }
            )

            status = 'Created' if created else 'Exists'
            self.stdout.write(f'{status}: Week {week_num} | {week_start} to {week_end}')

        self.stdout.write(self.style.SUCCESS('Campaign initialization complete!'))