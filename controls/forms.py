from django import forms
from .models import Project, ReportTemplate, Placeholder, AuditPeriod, ArbeitspapierSubmission, StandardPhrase


class ProjectForm(forms.ModelForm):
    class Meta:
        model = Project
        fields = ["audit_type", "report_kind", "report_date", "examination_date", "language"]
        widgets = {
            "report_date": forms.DateInput(attrs={"type": "date"}),
            "examination_date": forms.DateInput(attrs={"type": "date"}),
        }


class ExcelUploadForm(forms.Form):
    file = forms.FileField(label="IKS control list (.xlsx)")


class MultiFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class ArbeitspapierUploadForm(forms.Form):
    files = forms.FileField(
        label="Arbeitspapier working papers (.docx)",
        widget=MultiFileInput(attrs={"multiple": True}),
    )


class ReportTemplateForm(forms.ModelForm):
    class Meta:
        model = ReportTemplate
        fields = ["name", "audit_type", "language", "file", "is_active"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # File is required when uploading a new template, but optional when
        # editing metadata only (keep the existing file if none is chosen).
        if self.instance and self.instance.pk:
            self.fields["file"].required = False


class PlaceholderForm(forms.ModelForm):
    class Meta:
        model = Placeholder
        fields = ["key", "section", "description", "field_type"]
        help_texts = {
            "key": "How it appears in the template, e.g. engagement_partner -> {{ engagement_partner }}",
            "field_type": "Date fields get a date picker on the project form. Format (English/German) is chosen automatically from the project's report language.",
        }


class AuditPeriodForm(forms.ModelForm):
    class Meta:
        model = AuditPeriod
        fields = ["label", "start_date", "end_date"]
        widgets = {
            "start_date": forms.DateInput(attrs={"type": "date"}),
            "end_date": forms.DateInput(attrs={"type": "date"}),
        }


class SubmissionEditForm(forms.Form):
    """Edits a submission's imported text directly in the app."""
    test_activities_text = forms.CharField(
        label="Test activities (section 5.1) — one per line, indent sub-bullets with a Tab",
        widget=forms.Textarea(attrs={"rows": 8}),
        required=False,
    )
    kontrollergebnis_raw = forms.CharField(
        label="Kontrollergebnis (section 5.3)",
        widget=forms.Textarea(attrs={"rows": 3}),
        required=False,
    )
    no_deviation = forms.BooleanField(label="No deviation (renders as 'No deviations noted.')", required=False)


class StandardPhraseForm(forms.ModelForm):
    class Meta:
        model = StandardPhrase
        fields = ["language", "phrase"]
        help_texts = {
            "phrase": "Matched as a case-insensitive regex, exactly as written — e.g. \"Discussion with the process owner to .*\" or \"Review of .* with regard to the following aspects:\"",
        }


class ReferenceReportUploadForm(forms.Form):
    file = forms.FileField(label="Old report (.docx)")


class ReferenceReportCreateManualForm(forms.Form):
    name = forms.CharField(label="Reference report name", help_text='e.g. "CANCOM ISAE 3402 - additional entries"')


class ReferenceReportEntryForm(forms.Form):
    control_id = forms.CharField(label="Control ID", required=False)
    criteria = forms.CharField(label="Criteria", required=False)
    kontrollbeschreibung = forms.CharField(label="Kontrollbeschreibung", widget=forms.Textarea(attrs={"rows": 3}))
    test_performed = forms.CharField(label="Test performed", widget=forms.Textarea(attrs={"rows": 4}),
                                      help_text='Use "- " at the start of a line for a sub-bullet, same as elsewhere in the app.')
    result_text = forms.CharField(label="Result", required=False, widget=forms.Textarea(attrs={"rows": 2}))
