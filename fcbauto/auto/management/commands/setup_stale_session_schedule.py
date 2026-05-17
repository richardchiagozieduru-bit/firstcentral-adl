from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = (
        "Register a Django Q schedule to run detect_stale_sessions() every 15 minutes. "
        "Safe to run multiple times — will not create a duplicate if the schedule exists."
    )

    def handle(self, *args, **options):
        from django_q.models import Schedule

        func = 'auto.tasks.detect_stale_sessions'

        if Schedule.objects.filter(func=func).exists():
            self.stdout.write(self.style.WARNING(
                f"Schedule for '{func}' already exists. Nothing to do."
            ))
            return

        Schedule.objects.create(
            name='Detect stale upload sessions',
            func=func,
            schedule_type=Schedule.MINUTES,
            minutes=15,
            repeats=-1,  # run indefinitely
        )

        self.stdout.write(self.style.SUCCESS(
            f"Schedule created: '{func}' will run every 15 minutes."
        ))
