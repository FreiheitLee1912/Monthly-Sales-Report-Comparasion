from __future__ import annotations

import os
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file
from werkzeug.utils import secure_filename

from src.reports import (
    ReportResult,
    generate_excel_summary_report,
    generate_monthly_compare_report,
    generate_student_gantt_report,
)


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
REPORT_ROOT = Path(os.environ.get("REPORT_ROOT", "generated_reports"))


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/report")
def create_report():
    report_type = request.form.get("report_type", "summary")
    REPORT_ROOT.mkdir(parents=True, exist_ok=True)
    upload_dir = REPORT_ROOT / "_uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    try:
        if report_type == "summary":
            source = _save_upload("excel_file", upload_dir)
            result = generate_excel_summary_report(source, REPORT_ROOT)
        elif report_type == "compare":
            old_file = _save_upload("old_file", upload_dir)
            new_file = _save_upload("new_file", upload_dir)
            result = generate_monthly_compare_report(old_file, new_file, REPORT_ROOT)
        elif report_type == "student_gantt":
            source = _save_upload("excel_file", upload_dir)
            result = generate_student_gantt_report(source, REPORT_ROOT)
        else:
            return render_template("index.html", error="Unknown report type."), 400
    except ValueError as exc:
        return render_template("index.html", error=str(exc)), 400

    return render_template("result.html", result=result)


@app.get("/download/<job_id>/<artifact>")
def download(job_id: str, artifact: str):
    allowed = {"report.xlsx", "report.html", "report.pdf"}
    if artifact not in allowed:
        abort(404)
    path = REPORT_ROOT / secure_filename(job_id) / artifact
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=artifact)


def _save_upload(field_name: str, upload_dir: Path) -> Path:
    uploaded = request.files.get(field_name)
    if uploaded is None or uploaded.filename == "":
        raise ValueError(f"Missing required file: {field_name}")
    filename = secure_filename(uploaded.filename)
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise ValueError("Only Excel files are supported.")
    path = upload_dir / filename
    uploaded.save(path)
    return path


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
