from django.contrib import admin
from .models import Project, Control, ArbeitspapierSubmission, AuditPeriod, ReportTemplate, Placeholder, ChangeLogEntry, StandardPhrase, ReferenceReport, ReferenceReportEntry


class AuditPeriodInline(admin.TabularInline):
    model = AuditPeriod
    extra = 0


@admin.register(Project)
class ProjectAdmin(admin.ModelAdmin):
    list_display = ("customer_name", "audit_type", "report_kind", "language", "report_date", "updated_by", "updated_at")
    list_filter = ("audit_type", "report_kind", "language")
    search_fields = ("customer_name",)
    inlines = [AuditPeriodInline]


class SubmissionInline(admin.TabularInline):
    model = ArbeitspapierSubmission
    extra = 0
    readonly_fields = ("uploaded_at",)


@admin.register(Control)
class ControlAdmin(admin.ModelAdmin):
    list_display = ("control_id", "project", "kontrollziel", "validation_status")
    list_filter = ("project",)
    search_fields = ("control_id", "kontrollbeschreibung")
    inlines = [SubmissionInline]


@admin.register(ArbeitspapierSubmission)
class ArbeitspapierSubmissionAdmin(admin.ModelAdmin):
    list_display = ("control", "no_deviation", "source_file", "uploaded_by", "uploaded_at")
    list_filter = ("no_deviation",)


@admin.register(ReportTemplate)
class ReportTemplateAdmin(admin.ModelAdmin):
    list_display = ("name", "audit_type", "language", "is_active", "uploaded_by", "uploaded_at")
    list_filter = ("audit_type", "language", "is_active")


@admin.register(Placeholder)
class PlaceholderAdmin(admin.ModelAdmin):
    list_display = ("key", "section", "is_custom", "description")
    list_filter = ("section", "is_custom")


@admin.register(AuditPeriod)
class AuditPeriodAdmin(admin.ModelAdmin):
    list_display = ("project", "label", "start_date", "end_date")


@admin.register(ChangeLogEntry)
class ChangeLogEntryAdmin(admin.ModelAdmin):
    list_display = ("control", "event_type", "actor", "timestamp")
    list_filter = ("event_type",)


@admin.register(StandardPhrase)
class StandardPhraseAdmin(admin.ModelAdmin):
    list_display = ("language", "phrase", "active")
    list_filter = ("language", "active")


class ReferenceReportEntryInline(admin.TabularInline):
    model = ReferenceReportEntry
    extra = 0


@admin.register(ReferenceReport)
class ReferenceReportAdmin(admin.ModelAdmin):
    list_display = ("name", "uploaded_by", "uploaded_at")
    inlines = [ReferenceReportEntryInline]
