# IKS ControlHub

A Django app for running a multi-user internal control (IKS) review: create a
project space per audit engagement, import the master control list from
Excel, collect team members' working papers (Arbeitspapier), and generate the
final Word report from a template — matching the structure of your real
`BerichtVorlage.docx`.

Tested end-to-end against your real files: `IKS.xlsx` and the two
`*_Arbeitspapier_*.docx` working papers (see `sample_files/`).

## Why Django's ORM instead of SQLAlchemy

You asked for SQLAlchemy for the data layer and Django for admin/login.
Those two don't combine cleanly — Django ships its own ORM and its own
`User`/auth tables, and wiring SQLAlchemy in alongside it means maintaining
two separate database toolkits fighting over the same tables. Since you
wanted Django's built-in admin panel and login system, this project uses
**Django's own ORM** instead — it gives you the exact behavior you wanted
from SQLAlchemy (write the models once, run locally on SQLite, flip one
setting to point at PostgreSQL later) through Django's own migration system.
See "Switching to PostgreSQL" below — it's a single environment variable.

## Project structure

```
iks_platform/
  iks_platform/          Django project settings, urls
  controls/               The main app
    models.py             Project, Control, ArbeitspapierSubmission
    services.py            Excel import, Arbeitspapier parsing, report generation
    views.py / forms.py / urls.py
    admin.py                Django admin registration
    templates/controls/     HTML templates (project list/detail, upload UI)
  report_templates/
    report_template_en.docx   English report template (docxtpl placeholders)
    report_template_de.docx   German report template
  sample_files/             Your real IKS.xlsx + 2 Arbeitspapier files, for testing
  requirements.txt
  Procfile                  For DigitalOcean App Platform / gunicorn
  app.yaml                  DigitalOcean App Platform spec
  .env.example               Copy to .env for local dev
```

## Running locally (SQLite)

```bash
cd iks_platform
pip install -r requirements.txt
cp .env.example .env               # defaults to SQLite — no DB_ENGINE needed
python manage.py migrate
python manage.py createsuperuser   # your first login
python manage.py runserver
```

Open `http://127.0.0.1:8000/` — log in, create a project space, then on the
project page drag-and-drop `sample_files/IKS.xlsx` and both sample
Arbeitspapier files to see the pipeline work, then click "Generate report."

**Adding more users:** as the superuser, go to `/admin/` → Users → Add user.
That's the "basic admin + account creation" you asked for — Django ships
this, nothing custom needed.

## How the parsing actually works (tuned to your real files)

**Excel import** (`services.import_iks_excel`) reads the `ID`,
`Kontrollziel`, `Kontrollbeschreibung` columns from `IKS.xlsx` directly —
this matched your real file with no changes needed.

**Arbeitspapier parsing** (`services.parse_arbeitspapier`) scans every
paragraph for the "N.M" numbered-section pattern your working papers use
(e.g. `1.1  Kontrollnummer...`, `5.1  Durchgeführte Testaktivitäten`) and
groups the following paragraphs under that section number. It then pulls:

| Report field | Comes from |
|---|---|
| Control ID | Section 1.1, first word (e.g. `S-CC-1.0`) |
| Kontrollziel | Section 1.2 |
| Kontrollbeschreibung | Section 1.3 |
| Tests performed (bulleted) | Section 5.1 |
| Result of test | Section 5.3 |

For the result: if section 5.3 contains "keine Feststellung(en)", "keine
Abweichung(en)", "no deviation", or "no exception" (case-insensitive), the
report shows **"No deviations noted."** (or the German equivalent) instead of
the raw text — exactly the rule you asked for. Both of your sample working
papers hit this case and were verified to render correctly.

If your real working papers ever use different wording for "no findings",
add it to `NO_DEVIATION_MARKERS` in `controls/services.py`.

## The report templates

Both `report_template_en.docx` and `report_template_de.docx` were built by
taking your actual `BerichtVorlage.docx` — same fonts, same table styling,
same header row — and replacing the two example rows (S-CC-1.0, S-CC-1.1)
with docxtpl placeholder tags that repeat once per control:

- `{{ c.kontrollziel }}`, `{{ c.control_id }}`, `{{ c.kontrollbeschreibung }}`,
  `{{ c.result_text }}` — simple field substitution
- Test activities use docxtpl's paragraph-loop tag (`{%p for a in
  c.test_activities %}` / `{%p endfor %}`) to produce one bullet per activity
- The whole row repeats via `{%tr for c in controls %}` / `{%tr endfor %}`

Open either file directly in Word if you want to adjust layout, fonts, or
wording — docxtpl preserves all Word formatting around the tags.

**Project language**: whichever language you pick when creating the project
space, the download view (`views.download_report`) automatically picks
`report_template_en.docx` or `report_template_de.docx` to render against.
Note: this switches the *report's own labels* (column headers etc.) between
English/German — the actual control descriptions and test activities are
copied verbatim from your working papers regardless of language, since
that's your real underlying data.

## New in this version: templates, placeholders, validation, review workflow

**Report templates (Templates page, sidebar)** — upload Word templates per audit
type (SOC 2 / PS 951 / PS 3000) + language, with metadata (uploaded by, when).
Click "Preview" to view the template as a PDF right in the browser (requires
LibreOffice on the server — see `services.convert_docx_to_pdf`; falls back
gracefully if unavailable). `download_report` now looks up the active
template for the project's audit type/language via `services.resolve_template`,
falling back to the bundled `report_templates/*.docx` if nothing's uploaded yet.

**Placeholder catalog (Placeholders page, sidebar)** — documents every
`{{ }}` tag available in templates, categorized by section (project info /
audit period / controls / custom). "Add placeholder" registers a **custom**
field (e.g. `engagement_partner`) — it then shows up as an editable input on
each project's page, and is available in templates as `{{ extra.your_key }}`.
The "Placeholders" link on each template scans that specific `.docx` for
`{{ }}` tags and flags any that aren't yet in the catalog.

**Audit periods (Type 2 only)** — shown on the project page when
`report_kind = type2`. Add one or more date ranges (e.g. if the audit was
split into two periods); all are listed in the generated report via the
`{{ audit_periods }}` placeholder.

**Validation on working paper import** — every Arbeitspapier upload is
checked against the master IKS list: does the working paper's Kontrollziel
(1.2) and Kontrollbeschreibung (1.3) match what's already in the Control
record from the Excel import? Results show as a Match/Mismatch badge per
control, with a filter to show only mismatches. The check logic lives in
`services.py` as a small validator framework (`VALIDATORS` list) — adding a
new check later (e.g. a writing-style consistency check) means writing one
function with the same signature and adding it to that list. A stub
(`_validate_writing_style`) is already there, commented out, noting the
planned embeddings-based approach for comparing 5.1 phrasing across reports.

**Review sign-off ("geprüft von")** — "Mark reviewed" on each control records
who reviewed it and when (`Control.reviewed_by` / `reviewed_at`), separate
from who uploaded the working paper.

**Edit history** — every working paper import creates a new
`ArbeitspapierSubmission` row rather than overwriting the old one, so
`control_history` shows the full trail: every upload, who did it, when, and
the validation result at that time. Editing a submission's text in-app
(rather than re-uploading) is tracked separately via `edited_by`/`edited_at`.

**Inline editing of section 5.1 / 5.3** — "Edit" next to a control's latest
submission lets you correct the imported test activities and result text
directly in the app (useful for typos or quick fixes) without re-uploading
the working paper. There's no automatic write-back to the original `.docx`
file — the intent is you review/fix here, then manually copy the corrected
text back into the working paper yourself if it needs to live there too.

**Dashboard "last updated by"** — every project list row shows when and by
whom it was last touched (`Project.updated_by`/`updated_at`, refreshed on
edits and on every Excel/working-paper upload).

### Seeding the placeholder catalog

The built-in placeholders (customer_name, controls, audit_periods, etc.)
need to be loaded once:
```bash
python manage.py seed_placeholders
```
Safe to re-run — it only creates entries that don't already exist.


No code changes — just environment variables. In your `.env` (or
DigitalOcean App Platform's environment variables):

```
DB_ENGINE=postgres
DB_NAME=your-db-name
DB_USER=your-db-user
DB_PASSWORD=your-db-password
DB_HOST=your-db-host.db.ondigitalocean.com
DB_PORT=25060
```

Then run `python manage.py migrate` once against that database to create the
tables, and everything else works identically to the SQLite setup.

## Deploying to Render (Docker-based, includes template preview)

Render's native Python runtime can't `apt-get install` system packages, so
this repo includes a `Dockerfile` + `render.yaml` Blueprint for a
Docker-based deploy that installs LibreOffice (needed only for the "Preview"
button on uploaded templates — see the note below if you'd rather skip it).

1. Push this repo to GitHub (see the DigitalOcean section below for the git commands).
2. In the Render dashboard: **New +** → **Blueprint** → connect your repo. Render
   reads `render.yaml` automatically and provisions both the web service and
   a managed Postgres database, wiring the `DB_*` environment variables for you.
3. First deploy takes a few minutes longer than usual — the LibreOffice
   install inside the Docker build adds ~600MB+ and noticeably increases
   build time. This is a one-time cost per deploy, not per request.
4. `DJANGO_SECRET_KEY` is auto-generated by Render (see `generateValue: true`
   in `render.yaml`) — nothing to set manually.

### Skipping LibreOffice (smaller/faster deploy)

If you don't need the "Preview template as PDF" feature, delete the
`RUN apt-get install libreoffice` line from `Dockerfile` (or use Render's
native Python runtime instead of Docker entirely, via `Procfile` the same
way as the DigitalOcean setup below). Everything else — report generation,
uploads, validation — works identically without it; only the Preview button
on the Templates page will show its graceful "couldn't generate a preview"
message instead of an actual PDF.

## Deploying: GitHub → DigitalOcean App Platform

Note: the buildpack-based deploy below (via `Procfile`) does **not** include
LibreOffice, same situation as Render's native runtime — the "Preview
template as PDF" button will show its graceful fallback message rather than
an actual PDF. If you want that feature working on DigitalOcean too, point
App Platform at this repo's `Dockerfile` instead of using the buildpack
(App Platform supports both — choose "Dockerfile" as the source type when
creating the app).

1. **Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/iks-controlhub.git
   git push -u origin main
   ```
   (`.gitignore` already excludes `.env`, `db.sqlite3`, and `media/` — your
   secrets and local data won't be pushed.)

2. **Create a DigitalOcean Managed Postgres database** (App Platform → your
   app → Create/Attach Database, or standalone under Databases).

3. **Create the App Platform app**, either:
   - Via the DigitalOcean UI: "Create App" → connect your GitHub repo → it
     will detect `Procfile`/`requirements.txt` automatically → attach the
     Postgres database you created → set the env vars listed above (App
     Platform can auto-fill `DB_*` values from an attached database using
     its `${db.HOSTNAME}` style bindings, already set up in `app.yaml`).
   - Or via CLI: `doctl apps create --spec app.yaml` (edit the `repo:` field
     in `app.yaml` to your actual GitHub path first).

4. **Set `DJANGO_SECRET_KEY`** as a secret env var in the App Platform
   dashboard (don't reuse the default one from `.env.example`).

5. Every push to `main` auto-deploys (see `deploy_on_push: true` in
   `app.yaml`).

## What's still manual / worth knowing

- **Report template editing**: if you want to change wording, column widths,
  or fonts in the report, edit the `.docx` templates directly in Word — no
  code changes needed as long as you don't touch the `{{ }}` / `{% %}` tags.
- **New project fields or tables later** (e.g. AI-polish tracking, embeddings
  for Q&A — as discussed): add a Django model, run
  `python manage.py makemigrations && python manage.py migrate`. Works
  identically whether you're on SQLite locally or Postgres in production.
- **Multi-user editing**: Postgres handles concurrent writes natively; the
  `ForeignKey`s to Django's `User` model (`created_by`, `uploaded_by`) are
  already in place if you want to restrict who can edit what later.
