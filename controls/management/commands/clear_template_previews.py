from django.core.management.base import BaseCommand
from controls.models import ReportTemplate


class Command(BaseCommand):
    help = "Deletes all cached template preview PDFs (both the files on disk and the DB references). Safe to re-run."

    def handle(self, *args, **options):
        count = 0
        for template in ReportTemplate.objects.exclude(preview_pdf=""):
            if template.preview_pdf:
                template.preview_pdf.delete(save=False)
                template.preview_pdf = None
                template.save(update_fields=["preview_pdf"])
                count += 1
        self.stdout.write(self.style.SUCCESS(f"Cleared {count} cached preview PDF(s)"))
