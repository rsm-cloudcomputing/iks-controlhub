import os
import json
from pathlib import Path

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import HttpResponse, FileResponse, Http404, JsonResponse
from django.conf import settings
from django.utils import timezone

from .models import Project, Control, ArbeitspapierSubmission, ReportTemplate, Placeholder, AuditPeriod, StandardPhrase, ReferenceReport, ReferenceReportEntry
from .forms import (
    ProjectForm, ExcelUploadForm, ArbeitspapierUploadForm,
    ReportTemplateForm, PlaceholderForm, AuditPeriodForm, SubmissionEditForm, StandardPhraseForm,
    ReferenceReportUploadForm, ReferenceReportCreateManualForm, ReferenceReportEntryForm,
)
from .services import (
    import_iks_excel, import_arbeitspapier, generate_report,
    resolve_template, convert_docx_to_pdf, scan_template_placeholders,
    resolve_documented_ancestor, import_reference_report,
)


# ---------------------------------------------------------------------------
# Dashboard / Projects
# ---------------------------------------------------------------------------

# Custom placeholders in this order appear first (in this exact sequence) on
# the project form; any other custom placeholders follow, alphabetically.
CUSTOM_PLACEHOLDER_PRIORITY = [
    "customer_long_name", "customer_short_name", "customer_specification",
    "customer_long_address", "customer_federal_state",
]


def _ordered_custom_placeholders():
    placeholders = list(Placeholder.objects.filter(is_custom=True).order_by("created_at", "id"))

    def sort_key(ph):
        try:
            return (0, CUSTOM_PLACEHOLDER_PRIORITY.index(ph.key))
        except ValueError:
            return (1, ph.created_at)  # any new placeholder lands after the priority ones, in creation order

    return sorted(placeholders, key=sort_key)


@login_required
def project_list(request):
    projects = Project.objects.all().order_by("-updated_at")
    all_controls = Control.objects.filter(project__in=projects)
    total_controls = all_controls.count()
    submitted = sum(1 for c in all_controls if c.latest_submission)
    missing = total_controls - submitted
    stats = {
        "total_projects": projects.count(),
        "total_controls": total_controls,
        "submitted": submitted,
        "missing": missing,
    }
    return render(request, "controls/project_list.html", {"projects": projects, "stats": stats})


BUILTIN_LABEL_KEYS = ["audit_type", "report_kind", "report_date"]


def _builtin_placeholder_labels():
    return {p.key: p for p in Placeholder.objects.filter(key__in=BUILTIN_LABEL_KEYS, is_custom=False)}


def _split_custom_placeholders_for_form():
    """
    Splits custom placeholders into two groups for the project form: the
    priority customer_* fields render between Report kind and Report date,
    while anything else (like iks_date) renders after Examination date --
    matching their position in the Placeholder catalog table.
    """
    all_custom = _ordered_custom_placeholders()
    before = [p for p in all_custom if p.key in CUSTOM_PLACEHOLDER_PRIORITY]
    after = [p for p in all_custom if p.key not in CUSTOM_PLACEHOLDER_PRIORITY]
    return before, after


@login_required
def project_create(request):
    custom_placeholders, custom_placeholders_after = _split_custom_placeholders_for_form()
    builtin_placeholders = _builtin_placeholder_labels()
    if request.method == "POST":
        form = ProjectForm(request.POST)
        if form.is_valid():
            project = form.save(commit=False)
            project.created_by = request.user
            project.updated_by = request.user

            extra = {}
            for ph in custom_placeholders + custom_placeholders_after:
                extra[ph.key] = request.POST.get(f"extra__{ph.key}", "")
            # audit_conducted_from/to aren't standalone catalog placeholders
            # (they combine into audit_conducted_from_to at render time), so
            # they need to be captured explicitly here.
            extra["audit_conducted_from"] = request.POST.get("extra__audit_conducted_from", "")
            extra["audit_conducted_to"] = request.POST.get("extra__audit_conducted_to", "")
            project.extra_fields = extra

            # customer_name/customer_address are no longer shown as separate
            # form fields -- derived from the new customer_* placeholders so
            # the rest of the app (project list, headers, etc.) still works.
            project.customer_name = extra.get("customer_long_name") or extra.get("customer_short_name") or ""
            project.customer_address = extra.get("customer_long_address") or extra.get("customer_short_address") or ""
            project.save()

            # Optional audit period rows submitted alongside the form (Type 2)
            starts = request.POST.getlist("period_start[]")
            ends = request.POST.getlist("period_end[]")
            periods_added = 0
            for start, end in zip(starts, ends):
                if start and end:
                    AuditPeriod.objects.create(project=project, start_date=start, end_date=end)
                    periods_added += 1

            messages.success(request, f'Project "{project.customer_name}" created.' + (f" Added {periods_added} audit period(s)." if periods_added else ""))
            return redirect("project_detail", pk=project.pk)
    else:
        form = ProjectForm()
    return render(request, "controls/project_form.html", {
        "form": form, "custom_placeholders": custom_placeholders, "custom_placeholders_after": custom_placeholders_after,
        "builtin_placeholders": builtin_placeholders,
    })


def _revalidate_project_submissions(project):
    """
    Re-runs validation for every control's latest submission in a project --
    needed whenever the report language changes, since the standard-phrasing
    check (and its EN/DE phrase set) depends on it. Without this, the
    Validation column would keep showing results computed against the OLD
    language until the next re-import.
    """
    from .services import run_validations
    for control in project.controls.all():
        submission = control.latest_submission
        if not submission:
            continue
        control_values = {"kontrollziel": control.kontrollziel, "kontrollbeschreibung": control.kontrollbeschreibung}
        parsed_like = {
            "kontrollziel": control.kontrollziel,
            "kontrollbeschreibung": control.kontrollbeschreibung,
            "test_activities": submission.test_activities,
        }
        validation_results = run_validations(control_values, parsed_like, project.language)
        submission.validation_results = json.dumps(validation_results)
        submission.save(update_fields=["validation_results"])


@login_required
def project_update_field(request, pk):
    """AJAX: save a single project field immediately, no full page reload -- used
    for Report type on the edit page, so subsequent full-page actions (like
    Add audit period) see the correct just-changed value after reloading."""
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST" and request.headers.get("X-Requested-With") == "XMLHttpRequest":
        field = request.POST.get("field")
        value = request.POST.get("value")
        allowed_fields = {"report_kind", "audit_type", "language"}
        if field in allowed_fields and value:
            old_language = project.language
            setattr(project, field, value)
            project.updated_by = request.user
            project.save(update_fields=[field, "updated_by", "updated_at"])
            if field == "language" and value != old_language:
                _revalidate_project_submissions(project)
            return JsonResponse({"status": "ok"})
        return JsonResponse({"status": "error", "message": "Field not allowed"}, status=400)
    return JsonResponse({"status": "error"}, status=400)


@login_required
def project_update(request, pk):
    project = get_object_or_404(Project, pk=pk)
    custom_placeholders, custom_placeholders_after = _split_custom_placeholders_for_form()
    if request.method == "POST":
        old_language = project.language
        form = ProjectForm(request.POST, instance=project)
        if form.is_valid():
            project = form.save(commit=False)
            project.updated_by = request.user

            extra = dict(project.extra_fields or {})
            for ph in custom_placeholders + custom_placeholders_after:
                extra[ph.key] = request.POST.get(f"extra__{ph.key}", "")
            extra["audit_conducted_from"] = request.POST.get("extra__audit_conducted_from", "")
            extra["audit_conducted_to"] = request.POST.get("extra__audit_conducted_to", "")
            project.extra_fields = extra

            project.customer_name = extra.get("customer_long_name") or extra.get("customer_short_name") or ""
            project.customer_address = extra.get("customer_long_address") or extra.get("customer_short_address") or ""
            project.save()

            if project.language != old_language:
                _revalidate_project_submissions(project)
                messages.success(request, f'Project "{project.customer_name}" updated. Validation results were re-checked against the new report language.')
            else:
                messages.success(request, f'Project "{project.customer_name}" updated.')
            return redirect("project_detail", pk=project.pk)
    else:
        form = ProjectForm(instance=project)
    return render(request, "controls/project_form.html", {
        "form": form, "project": project, "is_edit": True, "custom_placeholders": custom_placeholders,
        "custom_placeholders_after": custom_placeholders_after,
        "audit_periods": project.audit_periods.all(), "builtin_placeholders": _builtin_placeholder_labels(),
    })


@login_required
def project_delete(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        name = project.customer_name
        project.delete()
        messages.success(request, f'Project "{name}" and all its controls/submissions were deleted.')
        return redirect("project_list")
    return render(request, "controls/project_confirm_delete.html", {"project": project})


@login_required
def controls_bulk_delete_submissions(request, pk):
    """
    Clears the working paper data (Test activities, Kontrollergebnis, Status,
    Validation, Geprüft von) for the SELECTED controls only -- the controls
    themselves (ID, Kontrollziel, Kontrollbeschreibung from the Excel import)
    are kept, ready for a corrected working paper to be re-imported later.
    """
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        control_ids = request.POST.getlist("control_ids")
        if not control_ids:
            messages.warning(request, "No controls were selected.")
            return redirect("project_detail", pk=pk)

        from .models import ChangeLogEntry
        controls = Control.objects.filter(project=project, pk__in=control_ids)
        cleared = 0
        for control in controls:
            latest = control.latest_submission
            if not latest:
                continue
            ChangeLogEntry.objects.create(
                control=control, event_type="delete", actor=request.user,
                source_file=latest.source_file,
                diff={
                    "test_activities": {"old": latest.test_activities_as_editable_text(), "new": ""},
                    "kontrollergebnis_raw": {"old": latest.kontrollergebnis_raw, "new": ""},
                },
            )
            control.submissions.all().delete()
            cleared += 1

        project.updated_by = request.user
        project.save(update_fields=["updated_by", "updated_at"])

        if cleared:
            messages.success(request, f"Cleared working paper data for {cleared} control(s). The controls themselves were kept — import corrected working papers whenever ready.")
        else:
            messages.info(request, "None of the selected controls had any working paper data to clear.")
    return redirect("project_detail", pk=pk)


@login_required
def controls_bulk_delete(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        control_ids = request.POST.getlist("control_ids")
        if control_ids:
            deleted_count, _ = Control.objects.filter(project=project, pk__in=control_ids).delete()
            project.updated_by = request.user
            project.save(update_fields=["updated_by", "updated_at"])
            messages.success(request, f"Deleted {len(control_ids)} control(s).")
        else:
            messages.warning(request, "No controls were selected.")
    return redirect("project_detail", pk=pk)


@login_required
def project_detail(request, pk):
    project = get_object_or_404(Project, pk=pk)
    controls = project.controls.all().order_by("control_id")

    # Filtering (controls section)
    validation_filter = request.GET.get("validation", "")
    if validation_filter == "finding":
        controls = [c for c in controls if c.latest_submission and not c.latest_submission.no_deviation]
    elif validation_filter:
        controls = [c for c in controls if c.validation_status == validation_filter]
    else:
        controls = list(controls)

    submitted = sum(1 for c in project.controls.all() if c.latest_submission)
    missing = project.controls.count() - submitted
    mismatch = sum(1 for c in project.controls.all() if c.validation_status == "mismatch")
    stats = {"total": project.controls.count(), "submitted": submitted, "missing": missing, "mismatch": mismatch}

    excel_form = ExcelUploadForm()
    arbeitspapier_form = ArbeitspapierUploadForm()
    audit_period_form = AuditPeriodForm()

    # Custom placeholder fields (from the catalog) for this project's extra_fields
    custom_placeholders = _ordered_custom_placeholders()

    return render(request, "controls/project_detail.html", {
        "project": project,
        "controls": controls,
        "stats": stats,
        "excel_form": excel_form,
        "arbeitspapier_form": arbeitspapier_form,
        "audit_period_form": audit_period_form,
        "audit_periods": project.audit_periods.all(),
        "custom_placeholders": custom_placeholders,
        "current_validation_filter": validation_filter,
    })


# ---------------------------------------------------------------------------
# Uploads
# ---------------------------------------------------------------------------

@login_required
def upload_excel(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                count = import_iks_excel(project, request.FILES["file"])
                project.updated_by = request.user
                project.save(update_fields=["updated_by", "updated_at"])
                if count:
                    messages.success(request, f"Imported {count} controls from the IKS control list.")
                else:
                    messages.warning(request, "The file was read but 0 rows were imported — check the Excel has data below the header row.")
            except Exception as e:
                messages.error(request, f"Import failed: {e}")
        else:
            messages.error(request, f"Upload rejected: {form.errors.as_text()}")
    return redirect("project_detail", pk=pk)


@login_required
def upload_arbeitspapier(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        files = request.FILES.getlist("files")
        if not files:
            messages.error(request, "No files were received by the server — check the upload field name/selection.")
            return redirect("project_detail", pk=pk)
        imported, failed, mismatches = 0, [], []
        for f in files:
            try:
                submission = import_arbeitspapier(project, f, f.name, request.user)
                imported += 1
                if submission.control.validation_status == "mismatch":
                    mismatches.append(submission.control.control_id)
            except Exception as e:
                failed.append(f"{f.name}: {e}")
        if imported:
            project.updated_by = request.user
            project.save(update_fields=["updated_by", "updated_at"])
            messages.success(request, f"Imported {imported} working paper(s).")
        if mismatches:
            messages.warning(request, f"Validation mismatch found for: {', '.join(mismatches)} — check the Controls table below.")
        for err in failed:
            messages.error(request, err)
    return redirect("project_detail", pk=pk)


# ---------------------------------------------------------------------------
# Audit periods (Type 2)
# ---------------------------------------------------------------------------

@login_required
def audit_period_add(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        form = AuditPeriodForm(request.POST)
        if form.is_valid():
            period = form.save(commit=False)
            period.project = project
            period.save()
            project.updated_by = request.user
            project.save(update_fields=["updated_by", "updated_at"])
            messages.success(request, "Audit period added.")
        else:
            messages.error(request, f"Could not add audit period: {form.errors.as_text()}")
    next_url = request.POST.get("next")
    if next_url:
        return redirect(next_url)
    return redirect("project_detail", pk=pk)


@login_required
def audit_period_delete(request, pk, period_pk):
    period = get_object_or_404(AuditPeriod, pk=period_pk, project_id=pk)
    period.delete()
    project = Project.objects.get(pk=pk)
    project.updated_by = request.user
    project.save(update_fields=["updated_by", "updated_at"])
    messages.success(request, "Audit period removed.")
    next_url = request.POST.get("next")
    if next_url:
        return redirect(next_url)
    return redirect("project_detail", pk=pk)


# ---------------------------------------------------------------------------
# Custom placeholder values (per-project)
# ---------------------------------------------------------------------------

@login_required
def update_extra_fields(request, pk):
    project = get_object_or_404(Project, pk=pk)
    if request.method == "POST":
        custom_keys = Placeholder.objects.filter(is_custom=True).values_list("key", flat=True)
        extra = dict(project.extra_fields or {})
        for key in custom_keys:
            extra[key] = request.POST.get(f"extra__{key}", "")
        project.extra_fields = extra
        project.updated_by = request.user
        project.save()
        messages.success(request, "Custom field values updated.")
    return redirect("project_detail", pk=pk)


# ---------------------------------------------------------------------------
# Controls: review sign-off, edit imported text
# ---------------------------------------------------------------------------

@login_required
def control_toggle_review(request, pk, control_pk):
    """
    Toggles the "reviewed OK" flag for a control -- highlights the whole row
    green in the Controls table. AJAX-friendly (no page reload/scroll jump).
    """
    control = get_object_or_404(Control, pk=control_pk, project_id=pk)
    if request.method == "POST":
        control.reviewed_ok = not control.reviewed_ok
        if control.reviewed_ok:
            control.reviewed_ok_by = request.user
            control.reviewed_ok_at = timezone.now()
        else:
            control.reviewed_ok_by = None
            control.reviewed_ok_at = None
        control.save()

        project = control.project
        project.updated_by = request.user
        project.save(update_fields=["updated_by", "updated_at"])

        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return JsonResponse({"status": "ok", "reviewed_ok": control.reviewed_ok})
    return redirect("project_detail", pk=pk)


@login_required
def control_history(request, pk, control_pk):
    control = get_object_or_404(Control, pk=control_pk, project_id=pk)
    entries = control.change_log.all().order_by("-timestamp")
    return render(request, "controls/control_history.html", {"control": control, "entries": entries})


@login_required
def submission_delete(request, pk, submission_pk):
    """
    Deletes ONE bad working paper import (ArbeitspapierSubmission) without
    touching the Control itself -- the control (and its Excel-imported
    Kontrollziel/Kontrollbeschreibung) stays, ready for a corrected working
    paper to be re-imported later. If this was the control's only
    submission, it simply reverts to "Not yet submitted".
    """
    submission = get_object_or_404(ArbeitspapierSubmission, pk=submission_pk, control__project_id=pk)
    if request.method == "POST":
        control = submission.control
        source_file = submission.source_file

        from .models import ChangeLogEntry
        ChangeLogEntry.objects.create(
            control=control, event_type="delete", actor=request.user,
            source_file=source_file,
            diff={
                "test_activities": {"old": submission.test_activities_as_editable_text(), "new": ""},
                "kontrollergebnis_raw": {"old": submission.kontrollergebnis_raw, "new": ""},
            },
        )

        submission.delete()

        project = control.project
        project.updated_by = request.user
        project.save(update_fields=["updated_by", "updated_at"])

        messages.success(request, f"Deleted the working paper import for {control.control_id} ({source_file}). The control itself was kept — import a corrected working paper whenever it's ready.")
    return redirect("project_detail", pk=pk)


@login_required
def submission_edit(request, pk, submission_pk):
    """
    Saves inline edits made directly in the Controls table (test activities,
    Kontrollergebnis, no-deviation flag). Also touches the parent Project's
    updated_at/updated_by so the dashboard's "last updated" reflects this.
    """
    submission = get_object_or_404(ArbeitspapierSubmission, pk=submission_pk, control__project_id=pk)
    if request.method == "POST":
        from .services import NO_DEVIATION_MARKERS, compute_diff, run_validations

        old_values = {
            "test_activities": submission.test_activities_as_editable_text(),
            "kontrollergebnis_raw": submission.kontrollergebnis_raw,
        }

        submission.test_activities = ArbeitspapierSubmission.parse_editable_text(request.POST.get("test_activities_text", ""))
        kontrollergebnis = request.POST.get("kontrollergebnis_raw", "")
        submission.kontrollergebnis_raw = kontrollergebnis
        submission.no_deviation = any(marker in kontrollergebnis.lower() for marker in NO_DEVIATION_MARKERS)
        submission.edited_by = request.user
        submission.edited_at = timezone.now()

        # Recheck validation against the current master control values, using
        # the freshly edited text -- an edit can fix (or introduce) a mismatch,
        # so the Validation column needs to reflect the edited content, not
        # whatever was true at the original import.
        control = submission.control
        control_values = {"kontrollziel": control.kontrollziel, "kontrollbeschreibung": control.kontrollbeschreibung}
        parsed_like = {
            "kontrollziel": control.kontrollziel,
            "kontrollbeschreibung": control.kontrollbeschreibung,
            "test_activities": submission.test_activities,
        }
        validation_results = run_validations(control_values, parsed_like, control.project.language)
        submission.validation_results = json.dumps(validation_results)

        submission.save()

        new_values = {
            "test_activities": submission.test_activities_as_editable_text(),
            "kontrollergebnis_raw": submission.kontrollergebnis_raw,
        }
        diff = compute_diff(old_values, new_values, ["test_activities", "kontrollergebnis_raw"])
        if diff:
            from .models import ChangeLogEntry
            ChangeLogEntry.objects.create(
                control=submission.control, event_type="edit", actor=request.user, diff=diff,
            )

        project = submission.control.project
        project.updated_by = request.user
        project.save(update_fields=["updated_by", "updated_at"])

        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return JsonResponse({
                "status": "ok",
                "control_id": submission.control.pk,
                "no_deviation": submission.no_deviation,
                "has_edit_history": submission.control.has_edit_history,
                "validation_results": validation_results,
            })

        messages.success(request, f"Updated working paper text for {submission.control.control_id}.")
    return redirect("project_detail", pk=pk)


# ---------------------------------------------------------------------------
# Report templates
# ---------------------------------------------------------------------------

@login_required
def template_toggle_active(request, pk):
    """AJAX: toggle a template's Active flag directly from the list table checkbox."""
    template = get_object_or_404(ReportTemplate, pk=pk)
    if request.method == "POST":
        template.is_active = not template.is_active
        template.save(update_fields=["is_active"])
        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return JsonResponse({"status": "ok", "is_active": template.is_active})
    return redirect("template_list")


@login_required
def template_list(request):
    templates = ReportTemplate.objects.all()
    form = ReportTemplateForm()
    return render(request, "controls/template_list.html", {"templates": templates, "form": form})


@login_required
def template_upload(request):
    if request.method == "POST":
        form = ReportTemplateForm(request.POST, request.FILES)
        if form.is_valid():
            template = form.save(commit=False)
            template.uploaded_by = request.user
            template.is_active = True  # new uploads are active by default
            template.save()
            # Try to generate a PDF preview (best-effort; fine if LibreOffice isn't available)
            pdf_path = convert_docx_to_pdf(template.file.path, Path(settings.MEDIA_ROOT) / "report_templates" / "previews")
            if pdf_path:
                with open(pdf_path, "rb") as f:
                    from django.core.files.base import ContentFile
                    template.preview_pdf.save(pdf_path.name, ContentFile(f.read()), save=True)
            messages.success(request, f'Template "{template.name}" uploaded.')
        else:
            messages.error(request, f"Upload rejected: {form.errors.as_text()}")
    return redirect("template_list")


@login_required
def template_update(request, pk):
    template = get_object_or_404(ReportTemplate, pk=pk)
    if request.method == "POST":
        form = ReportTemplateForm(request.POST, request.FILES, instance=template)
        if form.is_valid():
            new_file_uploaded = bool(request.FILES.get("file"))
            template = form.save()
            if new_file_uploaded:
                # Re-uploaded file invalidates any cached preview -- clear it so
                # the next Preview click regenerates from the new file.
                template.preview_pdf.delete(save=False)
                template.preview_pdf = None
                template.save(update_fields=["preview_pdf"])
            messages.success(request, f'Template "{template.name}" updated.')
            return redirect("template_list")
        else:
            messages.error(request, f"Could not update template: {form.errors.as_text()}")
    else:
        form = ReportTemplateForm(instance=template)
    return render(request, "controls/template_form.html", {"form": form, "template": template})


@login_required
def template_delete(request, pk):
    template = get_object_or_404(ReportTemplate, pk=pk)
    if request.method == "POST":
        name = template.name
        template.delete()
        messages.success(request, f'Template "{name}" deleted.')
        return redirect("template_list")
    return render(request, "controls/template_confirm_delete.html", {"template": template})


@login_required
def template_preview(request, pk):
    template = get_object_or_404(ReportTemplate, pk=pk)
    if not template.preview_pdf:
        # Lazily convert now (e.g. template was added via /admin/, or LibreOffice
        # wasn't available at upload time but is now).
        pdf_path = convert_docx_to_pdf(template.file.path, Path(settings.MEDIA_ROOT) / "report_templates" / "previews")
        if pdf_path:
            with open(pdf_path, "rb") as f:
                from django.core.files.base import ContentFile
                template.preview_pdf.save(pdf_path.name, ContentFile(f.read()), save=True)
        else:
            messages.error(request, "Couldn't generate a preview for this template — LibreOffice may not be available on this server.")
            return redirect("template_list")
    return FileResponse(template.preview_pdf.open("rb"), content_type="application/pdf")


@login_required
def template_placeholders(request, pk):
    template = get_object_or_404(ReportTemplate, pk=pk)
    scan = scan_template_placeholders(template.file.path)
    scanned = scan["placeholders"]
    loop_bindings = scan["loop_bindings"]
    if_conditions = scan["if_conditions"]
    placeholder_objs = list(Placeholder.objects.all())
    catalog = {p.key: p for p in placeholder_objs}
    catalog_keys = set(catalog.keys())

    used_top_level_keys = set()
    seen_keys = set()  # de-dup safety net -- scan already dedupes, but belt-and-suspenders

    rows = []
    for item in scanned:
        if item["key"] in seen_keys:
            continue
        seen_keys.add(item["key"])

        if item["is_loop_var"]:
            ancestor_key = resolve_documented_ancestor(item["key"], loop_bindings, catalog_keys)
            if ancestor_key:
                used_top_level_keys.add(ancestor_key)
            rows.append({
                "key": item["key"],
                "is_loop_var": True,
                "loop_source": item["loop_source"],
                "documented": ancestor_key is not None,
                "placeholder": catalog.get(ancestor_key),
            })
        else:
            used_top_level_keys.add(item["key"])
            rows.append({
                "key": item["key"],
                "is_loop_var": False,
                "documented": item["key"] in catalog,
                "placeholder": catalog.get(item["key"]),
            })

    # Also credit placeholders only referenced as a loop SOURCE, e.g. "controls"
    # in "{%tr for c in controls %}" -- that's real usage even with no {{ controls }} tag.
    for source_expr in loop_bindings.values():
        if source_expr in catalog:
            used_top_level_keys.add(source_expr)

    # Also credit placeholders only referenced in a {% if x %} conditional,
    # e.g. "is_type2" in "{% if is_type2 %}" -- never appears as a {{ }} tag.
    used_top_level_keys |= if_conditions

    # Reverse check: catalog entries that this template doesn't actually reference.
    # All placeholders (custom or built-in) render as bare {{ key }}.
    missing_from_template = []
    for key, placeholder in catalog.items():
        if key not in used_top_level_keys:
            missing_from_template.append(placeholder)

    return render(request, "controls/template_placeholders.html", {
        "template": template, "rows": rows, "missing_from_template": missing_from_template,
    })


# ---------------------------------------------------------------------------
# Placeholder catalog
# ---------------------------------------------------------------------------

@login_required
def placeholder_list(request):
    placeholders = list(Placeholder.objects.all())
    # Preserve the intentional section order (project info -> audit period -> controls -> custom)
    # rather than alphabetical, since that's the order these actually appear in a report.
    section_order = [key for key, _ in Placeholder.SECTION_CHOICES]
    section_labels = dict(Placeholder.SECTION_CHOICES)

    # Within "project info", match the exact field order on the project form:
    # audit_type, report_kind, [priority custom fields], report_date, examination_date.
    # Anything else (new custom placeholders, or the now-form-hidden
    # customer_name/customer_address) sorts after, in creation order.
    project_info_order = ["audit_type", "report_kind"] + CUSTOM_PLACEHOLDER_PRIORITY + ["report_date", "audit_conducted_from_to"]

    def project_info_sort_key(ph):
        if ph.key in project_info_order:
            return (0, project_info_order.index(ph.key))
        return (1, ph.created_at)

    by_section = {}
    for section_key in section_order:
        matching = [p for p in placeholders if p.section == section_key]
        if not matching:
            continue
        if section_key == "project_info":
            matching = sorted(matching, key=project_info_sort_key)
        by_section[section_labels[section_key]] = matching

    prefill_key = request.GET.get("prefill_key", "")
    form = PlaceholderForm(initial={"key": prefill_key} if prefill_key else None)
    return render(request, "controls/placeholder_list.html", {"by_section": by_section, "form": form, "prefill_key": prefill_key})


@login_required
def placeholder_create(request):
    if request.method == "POST":
        form = PlaceholderForm(request.POST)
        if form.is_valid():
            placeholder = form.save(commit=False)
            placeholder.is_custom = True
            placeholder.created_by = request.user
            placeholder.save()
            messages.success(request, f'Placeholder "{placeholder.key}" added — use it as {{{{ {placeholder.key} }}}} in templates, and fill in a value per project on the project page.')
        else:
            messages.error(request, f"Could not add placeholder: {form.errors.as_text()}")
    return redirect("placeholder_list")


@login_required
def placeholder_update(request, pk):
    placeholder = get_object_or_404(Placeholder, pk=pk)
    if request.method == "POST":
        form = PlaceholderForm(request.POST, instance=placeholder)
        if form.is_valid():
            form.save()
            messages.success(request, f'Placeholder "{placeholder.key}" updated.')
        else:
            messages.error(request, f"Could not update placeholder: {form.errors.as_text()}")
    return redirect("placeholder_list")


@login_required
def placeholder_delete(request, pk):
    placeholder = get_object_or_404(Placeholder, pk=pk)
    if request.method == "POST":
        key = placeholder.key
        was_custom = placeholder.is_custom
        placeholder.delete()
        note = "" if was_custom else " Note: this only removes it from the catalog/documentation — built-in fields are still supplied by the app's code regardless."
        messages.success(request, f'Placeholder "{key}" deleted.{note}')
    return redirect("placeholder_list")


# ---------------------------------------------------------------------------
# Report download
# ---------------------------------------------------------------------------

@login_required
def download_report(request, pk):
    project = get_object_or_404(Project, pk=pk)

    template = resolve_template(project)
    if template:
        template_path = template.file.path
    else:
        # Fall back to the bundled default templates if nothing's been uploaded yet.
        template_filename = f"report_template_{project.language}.docx"
        template_path = os.path.join(settings.BASE_DIR, "report_templates", template_filename)
        if not os.path.exists(template_path):
            messages.error(request, f"No report template found for {project.get_audit_type_display()} / {project.get_language_display()}. Upload one on the Templates page.")
            return redirect("project_detail", pk=pk)

    buffer = generate_report(project, template_path)
    filename = f"{project.customer_name.replace(' ', '_')}_{project.audit_type}_{project.report_kind}_report.docx"
    response = HttpResponse(
        buffer.read(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


# ---------------------------------------------------------------------------
# Standard phrases (per-language validation configuration)
# ---------------------------------------------------------------------------

@login_required
def standard_phrase_list(request):
    phrases = list(StandardPhrase.objects.all())
    language_order = [key for key, _ in StandardPhrase.LANGUAGE_CHOICES]
    language_labels = dict(StandardPhrase.LANGUAGE_CHOICES)
    by_language = {}
    for lang_key in language_order:
        matching = [p for p in phrases if p.language == lang_key]
        by_language[language_labels[lang_key]] = matching  # show even if empty, so it's obvious nothing's configured
    form = StandardPhraseForm()
    return render(request, "controls/standard_phrase_list.html", {"by_language": by_language, "form": form})


@login_required
def standard_phrase_create(request):
    if request.method == "POST":
        form = StandardPhraseForm(request.POST)
        if form.is_valid():
            phrase = form.save(commit=False)
            phrase.created_by = request.user
            phrase.save()
            messages.success(request, f'Standard phrase added for {phrase.get_language_display()}.')
        else:
            messages.error(request, f"Could not add phrase: {form.errors.as_text()}")
    return redirect("standard_phrase_list")


@login_required
def standard_phrase_update(request, pk):
    phrase = get_object_or_404(StandardPhrase, pk=pk)
    if request.method == "POST":
        form = StandardPhraseForm(request.POST, instance=phrase)
        if form.is_valid():
            form.save()
            messages.success(request, "Standard phrase updated.")
        else:
            messages.error(request, f"Could not update phrase: {form.errors.as_text()}")
    return redirect("standard_phrase_list")


@login_required
def standard_phrase_delete(request, pk):
    phrase = get_object_or_404(StandardPhrase, pk=pk)
    if request.method == "POST":
        phrase.delete()
        messages.success(request, "Standard phrase deleted.")
    return redirect("standard_phrase_list")


# ---------------------------------------------------------------------------
# Reference reports (old finished reports, imported for reviewer reference)
# ---------------------------------------------------------------------------

@login_required
def reference_report_create_manual(request):
    if request.method == "POST":
        form = ReferenceReportCreateManualForm(request.POST)
        if form.is_valid():
            report = ReferenceReport.objects.create(name=form.cleaned_data["name"], uploaded_by=request.user)
            messages.success(request, f'Created "{report.name}" — add entries to it below.')
            return redirect("reference_report_detail", pk=report.pk)
        else:
            messages.error(request, f"Could not create report: {form.errors.as_text()}")
    return redirect("reference_report_list")


@login_required
def reference_report_entry_add(request, pk):
    report = get_object_or_404(ReferenceReport, pk=pk)
    if request.method == "POST":
        form = ReferenceReportEntryForm(request.POST)
        if form.is_valid():
            order = report.entries.count()
            ReferenceReportEntry.objects.create(
                reference_report=report,
                order=order,
                control_id=form.cleaned_data["control_id"],
                criteria=form.cleaned_data["criteria"],
                kontrollbeschreibung=form.cleaned_data["kontrollbeschreibung"],
                test_performed=form.cleaned_data["test_performed"],
                result_text=form.cleaned_data["result_text"],
            )
            messages.success(request, "Entry added.")
        else:
            messages.error(request, f"Could not add entry: {form.errors.as_text()}")
    return redirect("reference_report_detail", pk=pk)


@login_required
def reference_report_entry_delete(request, pk, entry_pk):
    entry = get_object_or_404(ReferenceReportEntry, pk=entry_pk, reference_report_id=pk)
    if request.method == "POST":
        entry.delete()
        messages.success(request, "Entry deleted.")
    return redirect("reference_report_detail", pk=pk)


@login_required
def reference_report_list(request):
    reports = ReferenceReport.objects.all()
    all_entries = ReferenceReportEntry.objects.all().select_related("reference_report").order_by("reference_report__name", "order")
    form = ReferenceReportUploadForm()
    manual_form = ReferenceReportCreateManualForm()
    return render(request, "controls/reference_report_list.html", {
        "reports": reports, "form": form, "manual_form": manual_form, "all_entries": all_entries,
    })


@login_required
def reference_report_upload(request):
    if request.method == "POST":
        form = ReferenceReportUploadForm(request.POST, request.FILES)
        if form.is_valid():
            f = request.FILES["file"]
            try:
                report = import_reference_report(f, f.name, request.user)
                messages.success(request, f'Imported "{report.name}" — {report.entries.count()} control(s) found.')
            except ValueError as e:
                messages.error(request, str(e))
            except Exception as e:
                messages.error(request, f"Import failed: {e}")
        else:
            messages.error(request, f"Upload rejected: {form.errors.as_text()}")
    return redirect("reference_report_list")


@login_required
def reference_report_detail(request, pk):
    report = get_object_or_404(ReferenceReport, pk=pk)
    entries = report.entries.all()
    entry_form = ReferenceReportEntryForm()
    return render(request, "controls/reference_report_detail.html", {"report": report, "entries": entries, "entry_form": entry_form})


@login_required
def reference_report_delete(request, pk):
    report = get_object_or_404(ReferenceReport, pk=pk)
    if request.method == "POST":
        name = report.name
        report.delete()
        from .services import rebuild_reference_fts_index
        rebuild_reference_fts_index()
        messages.success(request, f'Deleted reference report "{name}".')
        return redirect("reference_report_list")
    return render(request, "controls/reference_report_confirm_delete.html", {"report": report})


@login_required
def submission_suggest(request, pk, submission_pk):
    """
    AJAX endpoint: given the current (possibly unsaved) test activities text
    from the textarea PLUS the control's Kontrollbeschreibung, returns the
    top 5 most similar past entries from the Reference Report library.

    Kontrollbeschreibung is always included, not just test_activities --
    it's the primary signal of what the control is even about, and it's
    especially important when test_activities is still empty or sparse
    (which is exactly when a reviewer most wants suggestions, to help them
    start writing rather than only after they've already written something).
    """
    submission = get_object_or_404(ArbeitspapierSubmission, pk=submission_pk, control__project_id=pk)
    if request.method == "POST":
        from .services import suggest_similar_test_activities
        test_activities_text = request.POST.get("test_activities_text", "")
        kontrollbeschreibung = submission.control.kontrollbeschreibung or ""
        suggestions = suggest_similar_test_activities(test_activities_text, kontrollbeschreibung, top_n=5)
        return JsonResponse({"suggestions": suggestions})
    return JsonResponse({"suggestions": []})
