from django.core.management.base import BaseCommand
from controls.models import StandardPhrase

EN_PHRASES = [
    "discussion with the process owner to .*",
    "inspection of existing process documentation:",
    "review of .*with regard to the following aspects:",
]

DE_PHRASES = [
    "gespräch mit dem prozessverantwortlichen .*",
    "einsichtnahme in die bestehende prozessdokumentation:",
    "überprüfung .*im hinblick auf die folgenden aspekte:",
]

# Earlier seed versions used slightly different wording -- remove those exact
# strings so re-running this command replaces them cleanly instead of leaving
# stale duplicates alongside the new ones.
OLD_EN_PHRASES = [
    "Discussion with the process owner to .*",
    "Inspection of existing process documentation",
    "Review of .* with regard to the following aspects:",
]
OLD_DE_PHRASES = [
    "Besprechung mit dem Prozessverantwortlichen .*",
    "Einsichtnahme in die bestehende Prozessdokumentation",
    "Überprüfung .* im Hinblick auf die folgenden Aspekte:",
]


class Command(BaseCommand):
    help = "Seeds default standard phrases for EN and DE (safe to re-run; replaces older default wording)"

    def handle(self, *args, **options):
        removed = 0
        for language, old_phrases in [("en", OLD_EN_PHRASES), ("de", OLD_DE_PHRASES)]:
            deleted, _ = StandardPhrase.objects.filter(language=language, phrase__in=old_phrases).delete()
            removed += deleted

        created = 0
        for language, phrases in [("en", EN_PHRASES), ("de", DE_PHRASES)]:
            for phrase in phrases:
                _, was_created = StandardPhrase.objects.get_or_create(language=language, phrase=phrase)
                created += int(was_created)

        self.stdout.write(self.style.SUCCESS(f"Seeded standard phrases ({created} new, {removed} old default(s) replaced)"))
