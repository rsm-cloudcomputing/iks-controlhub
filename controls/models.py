import json
from django.db import models
from django.contrib.auth.models import User


class Project(models.Model):
    """
    A single audit engagement / project space -- e.g. "Siemens SOC 2 Type 1, 2026".
    Everything else (controls, working papers, generated reports) belongs to one Project.
    """

    AUDIT_TYPE_CHOICES = [
        ("soc2", "SOC 2"),
        ("ps951", "PS 951"),
        ("ps3000", "PS 3000"),
        ("isae3402", "ISAE 3402"),
    ]
    REPORT_KIND_CHOICES = [
        ("type1", "Type 1"),
        ("type2", "Type 2"),
    ]
    LANGUAGE_CHOICES = [
        ("en", "English"),
        ("de", "Deutsch"),
    ]

    customer_name = models.CharField(max_length=200)
    customer_address = models.TextField(blank=True)
    audit_type = models.CharField(max_length=20, choices=AUDIT_TYPE_CHOICES, default="soc2")
    report_kind = models.CharField(max_length=10, choices=REPORT_KIND_CHOICES, default="type1")
    report_date = models.DateField()
    language = models.CharField(max_length=2, choices=LANGUAGE_CHOICES, default="en")

    # Type 1 engagements report a single point-in-time date rather than a period.
    examination_date = models.DateField(null=True, blank=True, help_text="Type 1 only: date of examination")

    # Custom placeholder values (see Placeholder catalog) -- e.g. {"engagement_partner": "Jane Doe"}
    extra_fields = models.JSONField(default=dict, blank=True)

    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="projects_created")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="projects_updated")
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.customer_name} -- {self.get_audit_type_display()} {self.get_report_kind_display()} ({self.report_date})"

    @property
    def is_type2(self):
        return self.report_kind == "type2"


class AuditPeriod(models.Model):
    """
    One audit period for a Type 2 engagement. Most projects have exactly one,
    but a period can be split (e.g. system changed mid-year) -- add as many
    as needed.
    """
    project = models.ForeignKey(Project, on_delete=models.CASCADE, related_name="audit_periods")
    label = models.CharField(max_length=100, blank=True, help_text='Optional, e.g. "Period 1"')
    start_date = models.DateField()
    end_date = models.DateField()

    class Meta:
        ordering = ["start_date"]

    def __str__(self):
        return f"{self.start_date} to {self.end_date}"


class ReportTemplate(models.Model):
    """
    An uploaded Word report template (docxtpl placeholders inside), scoped to
    an audit type + language. These are what generate_report() renders
    against -- replacing the two hardcoded report_templates/*.docx files.
    """
    AUDIT_TYPE_CHOICES = Project.AUDIT_TYPE_CHOICES
    LANGUAGE_CHOICES = Project.LANGUAGE_CHOICES

    name = models.CharField(max_length=150)
    audit_type = models.CharField(max_length=20, choices=AUDIT_TYPE_CHOICES)
    language = models.CharField(max_length=2, choices=LANGUAGE_CHOICES)
    file = models.FileField(upload_to="report_templates/")
    preview_pdf = models.FileField(upload_to="report_templates/previews/", blank=True, null=True)
    is_active = models.BooleanField(default=True, help_text="Used when generating reports for this audit type/language")

    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self):
        return f"{self.name} ({self.get_audit_type_display()}, {self.get_language_display()})"


class Placeholder(models.Model):
    """
    Documentation/catalog entry for a docxtpl placeholder available in report
    templates -- e.g. "customer_name", "controls", "audit_periods". Built-in
    ones ship with the app; custom ones can be added through the UI and are
    filled in per-project via Project.extra_fields.
    """
    SECTION_CHOICES = [
        ("project_info", "Project info"),
        ("audit_period", "Audit period"),
        ("controls", "Control table"),
        ("custom", "Custom / other"),
    ]
    FIELD_TYPE_CHOICES = [
        ("text", "Text"),
        ("date", "Date"),
    ]
    DATE_FORMAT_CHOICES = [
        ("en_ordinal", "English — e.g. 8th October 2025"),
        ("de_ordinal", "German — e.g. 1. Januar 2025"),
    ]

    key = models.CharField(max_length=100, unique=True, help_text='Used in templates as {{ key }}')
    section = models.CharField(max_length=20, choices=SECTION_CHOICES, default="custom")
    description = models.CharField(max_length=255, blank=True)
    is_custom = models.BooleanField(default=False, help_text="True for user-added placeholders (stored in Project.extra_fields)")
    field_type = models.CharField(max_length=10, choices=FIELD_TYPE_CHOICES, default="text",
                                   help_text="Date fields get a date picker on the project form, and are formatted per date_format when a report is generated")
    date_format = models.CharField(max_length=10, choices=DATE_FORMAT_CHOICES, default="en_ordinal", blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["section", "key"]

    def __str__(self):
        return self.key


class Control(models.Model):
    """
    One control from the master IKS control list (imported from the IKS.xlsx
    columns: ID, Kontrollziel, Kontrollbeschreibung), scoped to one Project.
    """

    project = models.ForeignKey(Project, on_delete=models.CASCADE, related_name="controls")
    control_id = models.CharField(max_length=50)
    kontrollziel = models.CharField(max_length=255, blank=True)
    kontrollbeschreibung = models.TextField(blank=True)

    # Manual "looks good" sign-off toggle -- highlights the whole row green
    # in the Controls table when checked. Separate from geprueft_von (which
    # is auto-extracted from the working paper itself) -- this is whoever is
    # reviewing the imported data in the app confirming it's fine.
    reviewed_ok = models.BooleanField(default=False)
    reviewed_ok_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="controls_reviewed_ok")
    reviewed_ok_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        unique_together = ("project", "control_id")
        ordering = ["control_id"]

    def __str__(self):
        return f"{self.control_id} ({self.project.customer_name})"

    @property
    def latest_submission(self):
        return self.submissions.order_by("-uploaded_at").first()

    @property
    def validation_status(self):
        """'match' / 'mismatch' / 'no_submission' -- summarizes the latest submission's checks."""
        s = self.latest_submission
        if not s:
            return "no_submission"
        results = s.validation_results_list
        if not results:
            return "no_submission"
        return "match" if all(r["passed"] for r in results) else "mismatch"

    @property
    def has_edit_history(self):
        """True if this control has ever been manually edited in the app (not just imported)."""
        return self.change_log.filter(event_type="edit").exists()


class ArbeitspapierSubmission(models.Model):
    """
    The extracted content from one team member's working paper (Arbeitspapier)
    for a single control: test activities (section 5.1) and control result (5.3).
    Each import creates a NEW row (never overwrites) -- this is what gives us
    the edit history / "who submitted what, when" trail per control.
    """

    control = models.ForeignKey(Control, on_delete=models.CASCADE, related_name="submissions")
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)

    test_activities = models.TextField(blank=True)  # JSON: [{"text":..., "level":0}, ...]
    kontrollergebnis_raw = models.TextField(blank=True)
    no_deviation = models.BooleanField(default=False)

    # Validation results from comparing this working paper against the master IKS list.
    # JSON: [{"check": "kontrollziel_match", "passed": true, "message": "..."}, ...]
    validation_results = models.TextField(blank=True)

    # Extracted from the working paper's own metadata table ("Geprüft: FNO", "Datum: ...")
    # -- this is who actually reviewed the control, per the working paper itself.
    geprueft_von = models.CharField(max_length=100, blank=True)
    geprueft_datum = models.DateField(null=True, blank=True)

    source_file = models.CharField(max_length=255, blank=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    # Set when a human manually edits the imported text in the app (vs. raw import)
    edited_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="submissions_edited")
    edited_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self):
        return f"Submission for {self.control.control_id} ({self.source_file})"

    @property
    def result_text(self):
        # Always the actual Kontrollergebnis (5.3) text as entered/imported --
        # previously this substituted a canned "No deviations noted." phrase
        # whenever no_deviation was True, silently discarding any additional
        # context in the real text.
        return self.kontrollergebnis_raw or "[Not yet assessed]"

    @property
    def test_activities_list(self):
        try:
            return json.loads(self.test_activities) if self.test_activities else []
        except (json.JSONDecodeError, TypeError):
            return [{"text": line.strip("- ").strip(), "level": 0}
                    for line in self.test_activities.splitlines() if line.strip()]

    @property
    def validation_results_list(self):
        try:
            return json.loads(self.validation_results) if self.validation_results else []
        except (json.JSONDecodeError, TypeError):
            return []

    def test_activities_as_editable_text(self):
        """
        Flattens the nested bullet structure into plain text for editing in a
        <textarea>: main bullets as-is, sub-bullets prefixed with "- ".
        """
        lines = []
        for item in self.test_activities_list:
            prefix = "- " if item.get("level") else ""
            lines.append(f"{prefix}{item.get('text', '')}")
        return "\n".join(lines)

    @staticmethod
    def parse_editable_text(raw_text):
        """Reverses test_activities_as_editable_text() back into the JSON structure."""
        items = []
        for line in raw_text.splitlines():
            if not line.strip():
                continue
            stripped = line.strip()
            if stripped.startswith("- "):
                items.append({"text": stripped[2:].strip(), "level": 1})
            else:
                items.append({"text": stripped, "level": 0})
        return json.dumps(items)


class ChangeLogEntry(models.Model):
    """
    One entry per import or manual edit of a control's working paper data.
    Unlike ArbeitspapierSubmission (which holds current state, one row per
    import), this is a pure append-only log capturing WHAT changed at each
    step, so the history page can show a real diff -- not just "submitted
    at time X" but "test_activities changed from A to B by user Y at time Z".
    """
    EVENT_CHOICES = [
        ("import", "Working paper imported"),
        ("edit", "Manually edited in app"),
        ("delete", "Submission deleted"),
    ]

    control = models.ForeignKey(Control, on_delete=models.CASCADE, related_name="change_log")
    event_type = models.CharField(max_length=10, choices=EVENT_CHOICES)
    actor = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    timestamp = models.DateTimeField(auto_now_add=True)
    source_file = models.CharField(max_length=255, blank=True)

    # {"field_name": {"old": "...", "new": "..."}, ...} -- only fields that actually changed
    diff = models.JSONField(default=dict, blank=True)

    class Meta:
        ordering = ["-timestamp"]
        verbose_name_plural = "Change log entries"

    def __str__(self):
        return f"{self.get_event_type_display()} on {self.control.control_id} at {self.timestamp}"


class StandardPhrase(models.Model):
    """
    A regex pattern (case-insensitive) that section 5.1 test activities are
    expected to contain, scoped to a language -- since EN and DE working
    papers use different standard wording. Configurable from the
    "Standard Phrases" sidebar page instead of being hardcoded.

    Write the phrase exactly as you want it matched, wildcards included --
    e.g. "Discussion with the process owner to .*" or
    "Review of .* with regard to the following aspects:".
    """
    LANGUAGE_CHOICES = Project.LANGUAGE_CHOICES

    language = models.CharField(max_length=2, choices=LANGUAGE_CHOICES)
    phrase = models.CharField(max_length=300, help_text="Matched as a case-insensitive regex, e.g. 'Review of .* with regard to the following aspects:'")
    active = models.BooleanField(default=True)

    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["language", "id"]

    def __str__(self):
        return f"[{self.language}] {self.phrase}"


class ReferenceReport(models.Model):
    """
    An old/historical finished report, imported purely as reference material
    for reviewers -- not tied to any single project. Its table rows (control
    description, tests performed, result) are parsed out so they're
    searchable later when reviewing a new working paper.
    """
    name = models.CharField(max_length=200)
    file = models.FileField(upload_to="reference_reports/", blank=True, null=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self):
        return self.name


class ReferenceReportEntry(models.Model):
    """One parsed row from a ReferenceReport's results table."""
    reference_report = models.ForeignKey(ReferenceReport, on_delete=models.CASCADE, related_name="entries")
    order = models.PositiveIntegerField(default=0)

    criteria = models.CharField(max_length=255, blank=True)
    control_id = models.CharField(max_length=50, blank=True)
    kontrollbeschreibung = models.TextField(blank=True)
    test_performed = models.TextField(blank=True)
    result_text = models.TextField(blank=True)

    class Meta:
        ordering = ["reference_report", "order"]

    def __str__(self):
        return f"{self.control_id or '(no id)'} — {self.reference_report.name}"
