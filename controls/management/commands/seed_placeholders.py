from django.core.management.base import BaseCommand
from controls.models import Placeholder

# (key, section, description, field_type, is_custom)
# field_type only matters for display (Date fields get a date picker on the
# project form; actual EN/DE formatting is always auto-decided from the
# project's report language, not stored per-placeholder).
PLACEHOLDERS = [
    # --- Project info, in the exact order shown on the project form ---
    # Descriptions swapped per explicit request: {{ report_kind }} now
    # documents the audit standard, {{ audit_type }} now documents Type 1/2.
    ("report_kind", "project_info", "Audit type display name (SOC 2 / PS 951 / PS 3000)", "text", False),
    ("audit_type", "project_info", "Type 1 or Type 2", "text", False),
    ("customer_long_name", "project_info", "Customer's full/formal legal name", "text", True),
    ("customer_short_name", "project_info", "Customer's short name/abbreviation", "text", True),
    ("customer_specification", "project_info", "Customer-specific detail or qualifier for the report", "text", True),
    ("customer_long_address", "project_info", "Customer's full formal address", "text", True),
    ("customer_federal_state", "project_info", "Customer's federal state/region", "text", True),
    ("report_date", "project_info", "Report date", "text", False),
    ("audit_conducted_from_to", "project_info", "Examination period conducted, e.g. \"1st January 2026 to 30th June 2026\"", "text", False),
    ("iks_date", "project_info", "Date the IKS-Beschreibung finalized", "date", True),
    # --- Audit period ---
    ("is_type2", "audit_period", "True/false -- use with {% if is_type2 %} to show period section", "text", False),
    ("audit_periods", "audit_period", "List of {label, start_date, end_date} -- loop with {% for p in audit_periods %}", "text", False),
    # --- Control table ---
    ("controls", "controls", "List of controls, each with control_id, kontrollziel, kontrollbeschreibung, test_activities, result_text", "text", False),
    ("c.control_id", "controls", "Control ID, e.g. S-CC-1.0 (inside the controls loop)", "text", False),
    ("c.kontrollziel", "controls", "Kontrollziel / criteria reference (inside the controls loop)", "text", False),
    ("c.kontrollbeschreibung", "controls", "Control description (inside the controls loop)", "text", False),
    ("c.test_activities", "controls", "Nested bullet structure: [{text, sub_items}] (inside the controls loop)", "text", False),
    ("c.result_text", "controls", "Result of test -- 'No deviations noted.' or the actual finding text", "text", False),
]

# These are no longer part of the visible project form (customer_name/
# customer_address are auto-derived from customer_long_name/long_address
# instead of being separate fields) -- removed from the catalog so they
# don't clutter the Placeholders page. The underlying context keys still
# get set on every report regardless, so this doesn't affect templates.
DEPRECATED_KEYS = ["customer_name", "customer_address", "customer_short_address", "examination_date"]


class Command(BaseCommand):
    help = "Seeds the placeholder catalog and default admin user (safe to re-run)"

    def handle(self, *args, **options):
        removed, _ = Placeholder.objects.filter(key__in=DEPRECATED_KEYS).delete()

        created = 0
        for key, section, description, field_type, is_custom in PLACEHOLDERS:
            _, was_created = Placeholder.objects.get_or_create(
                key=key, defaults={"section": section, "description": description, "field_type": field_type, "is_custom": is_custom}
            )
            created += int(was_created)
        self.stdout.write(self.style.SUCCESS(
            f"Seeded placeholder catalog ({created} new, {len(PLACEHOLDERS) - created} already existed, {removed} deprecated removed)"
        ))

        from django.contrib.auth.models import User
        if not User.objects.filter(username="taitran").exists():
            User.objects.create_superuser("taitran", "", "2812")
            self.stdout.write(self.style.SUCCESS('Created superuser "taitran"'))
        else:
            self.stdout.write("Superuser \"taitran\" already exists — skipped")
