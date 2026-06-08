# Online Excel Report App Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a deployable Python web app that accepts Excel uploads, lets the user select a report type, and returns XLSX, HTML, and PDF report outputs.

**Architecture:** Add a new Flask app at the repository root while preserving the existing Windows executable package. Report generation lives in `src/reports.py`, routes live in `app.py`, and templates live in `templates/`.

**Tech Stack:** Python 3.11, Flask, pandas, openpyxl, reportlab, pytest.

---

### Task 1: Project Scaffolding

**Files:**
- Create: `requirements.txt`
- Create: `render.yaml`
- Create: `src/__init__.py`
- Create: `tests/__init__.py`

- [x] Add runtime dependencies for Flask, Excel parsing, XLSX output, PDF output, and tests.
- [x] Add a Render web service definition that runs `gunicorn app:app`.

### Task 2: Report Engine

**Files:**
- Create: `src/reports.py`
- Test: `tests/test_reports.py`

- [x] Implement `generate_excel_summary_report()` for single-workbook row/column summary output.
- [x] Implement `generate_monthly_compare_report()` for old/new workbook comparison summaries.
- [x] Implement `generate_student_gantt_report()` for student task schedules using common English, Chinese, and Japanese column names.
- [x] Emit XLSX, HTML, and PDF files for each report job.

### Task 3: Flask UI

**Files:**
- Create: `app.py`
- Create: `templates/index.html`
- Create: `templates/result.html`

- [x] Add upload form with report type selector.
- [x] Validate required files by report type.
- [x] Store generated jobs under local `generated_reports/`.
- [x] Add download route for generated artifacts.

### Task 4: Verification

**Files:**
- Modify: `README.md`

- [x] Run unit tests with `pytest`.
- [x] Run Flask route smoke tests.
- [x] Add local and Render deployment instructions.
