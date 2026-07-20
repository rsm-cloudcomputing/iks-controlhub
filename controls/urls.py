from django.urls import path
from . import views

urlpatterns = [
    path("", views.project_list, name="project_list"),
    path("projects/new/", views.project_create, name="project_create"),
    path("projects/<int:pk>/", views.project_detail, name="project_detail"),
    path("projects/<int:pk>/edit/", views.project_update, name="project_update"),
    path("projects/<int:pk>/update-field/", views.project_update_field, name="project_update_field"),
    path("projects/<int:pk>/delete/", views.project_delete, name="project_delete"),
    path("projects/<int:pk>/upload-excel/", views.upload_excel, name="upload_excel"),
    path("projects/<int:pk>/upload-arbeitspapier/", views.upload_arbeitspapier, name="upload_arbeitspapier"),
    path("projects/<int:pk>/controls/bulk-delete/", views.controls_bulk_delete, name="controls_bulk_delete"),
    path("projects/<int:pk>/controls/bulk-delete-submissions/", views.controls_bulk_delete_submissions, name="controls_bulk_delete_submissions"),
    path("projects/<int:pk>/download-report/", views.download_report, name="download_report"),
    path("projects/<int:pk>/extra-fields/", views.update_extra_fields, name="update_extra_fields"),

    path("projects/<int:pk>/audit-periods/add/", views.audit_period_add, name="audit_period_add"),
    path("projects/<int:pk>/audit-periods/<int:period_pk>/delete/", views.audit_period_delete, name="audit_period_delete"),

    path("projects/<int:pk>/controls/<int:control_pk>/history/", views.control_history, name="control_history"),
    path("projects/<int:pk>/controls/<int:control_pk>/toggle-review/", views.control_toggle_review, name="control_toggle_review"),
    path("projects/<int:pk>/submissions/<int:submission_pk>/edit/", views.submission_edit, name="submission_edit"),
    path("projects/<int:pk>/submissions/<int:submission_pk>/delete/", views.submission_delete, name="submission_delete"),
    path("projects/<int:pk>/submissions/<int:submission_pk>/suggest/", views.submission_suggest, name="submission_suggest"),

    path("templates/", views.template_list, name="template_list"),
    path("templates/upload/", views.template_upload, name="template_upload"),
    path("templates/<int:pk>/toggle-active/", views.template_toggle_active, name="template_toggle_active"),
    path("templates/<int:pk>/edit/", views.template_update, name="template_update"),
    path("templates/<int:pk>/delete/", views.template_delete, name="template_delete"),
    path("templates/<int:pk>/preview/", views.template_preview, name="template_preview"),
    path("templates/<int:pk>/placeholders/", views.template_placeholders, name="template_placeholders"),

    path("placeholders/", views.placeholder_list, name="placeholder_list"),
    path("placeholders/add/", views.placeholder_create, name="placeholder_create"),
    path("placeholders/<int:pk>/edit/", views.placeholder_update, name="placeholder_update"),
    path("placeholders/<int:pk>/delete/", views.placeholder_delete, name="placeholder_delete"),

    path("standard-phrases/", views.standard_phrase_list, name="standard_phrase_list"),
    path("standard-phrases/add/", views.standard_phrase_create, name="standard_phrase_create"),
    path("standard-phrases/<int:pk>/edit/", views.standard_phrase_update, name="standard_phrase_update"),
    path("standard-phrases/<int:pk>/delete/", views.standard_phrase_delete, name="standard_phrase_delete"),

    path("reference-reports/", views.reference_report_list, name="reference_report_list"),
    path("reference-reports/upload/", views.reference_report_upload, name="reference_report_upload"),
    path("reference-reports/create-manual/", views.reference_report_create_manual, name="reference_report_create_manual"),
    path("reference-reports/<int:pk>/", views.reference_report_detail, name="reference_report_detail"),
    path("reference-reports/<int:pk>/delete/", views.reference_report_delete, name="reference_report_delete"),
    path("reference-reports/<int:pk>/entries/add/", views.reference_report_entry_add, name="reference_report_entry_add"),
    path("reference-reports/<int:pk>/entries/<int:entry_pk>/delete/", views.reference_report_entry_delete, name="reference_report_entry_delete"),
]
