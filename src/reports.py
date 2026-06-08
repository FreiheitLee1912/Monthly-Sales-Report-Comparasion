from __future__ import annotations

from dataclasses import dataclass
from html import escape
from pathlib import Path
from uuid import uuid4

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib import colors


@dataclass(frozen=True)
class ReportResult:
    title: str
    job_id: str
    xlsx_path: Path
    html_path: Path
    pdf_path: Path


def generate_excel_summary_report(source: Path, output_root: Path) -> ReportResult:
    job_dir = _make_job_dir(output_root)
    sheets = pd.read_excel(source, sheet_name=None)
    rows = []
    for sheet_name, frame in sheets.items():
        rows.append(
            {
                "Sheet": sheet_name,
                "Rows": len(frame),
                "Columns": len(frame.columns),
                "Empty cells": int(frame.isna().sum().sum()),
                "Numeric columns": len(frame.select_dtypes(include="number").columns),
            }
        )
    summary = pd.DataFrame(rows)
    title = "Excel Summary Report"
    return _write_report(title, job_dir, [("Summary", summary)], _summary_html(title, source.name, summary))


def generate_monthly_compare_report(old_file: Path, new_file: Path, output_root: Path) -> ReportResult:
    job_dir = _make_job_dir(output_root)
    old_frame = pd.read_excel(old_file)
    new_frame = pd.read_excel(new_file)

    old_keyed = _key_rows(old_frame)
    new_keyed = _key_rows(new_frame)
    old_keys = set(old_keyed)
    new_keys = set(new_keyed)
    added = sorted(new_keys - old_keys)
    removed = sorted(old_keys - new_keys)
    common = sorted(old_keys & new_keys)
    changed = [key for key in common if old_keyed[key] != new_keyed[key]]

    summary = pd.DataFrame(
        [
            {"Metric": "Old rows", "Count": len(old_frame)},
            {"Metric": "New rows", "Count": len(new_frame)},
            {"Metric": "Added rows", "Count": len(added)},
            {"Metric": "Removed rows", "Count": len(removed)},
            {"Metric": "Changed rows", "Count": len(changed)},
        ]
    )
    details = pd.DataFrame(
        [{"Change type": "Added rows", "Row key": key} for key in added]
        + [{"Change type": "Removed rows", "Row key": key} for key in removed]
        + [{"Change type": "Changed rows", "Row key": key} for key in changed]
    )
    title = "Monthly Compare Report"
    html = _summary_html(title, f"{old_file.name} -> {new_file.name}", summary, details)
    return _write_report(title, job_dir, [("Summary", summary), ("Details", details)], html)


def generate_student_gantt_report(source: Path, output_root: Path) -> ReportResult:
    job_dir = _make_job_dir(output_root)
    frame = pd.read_excel(source)
    normalized = _normalize_gantt_columns(frame)
    normalized["start"] = pd.to_datetime(normalized["start"], errors="coerce")
    normalized["end"] = pd.to_datetime(normalized["end"], errors="coerce")
    normalized = normalized.dropna(subset=["student", "task", "start", "end"]).copy()
    normalized["days"] = (normalized["end"] - normalized["start"]).dt.days + 1
    normalized["progress"] = pd.to_numeric(normalized.get("progress", 0), errors="coerce").fillna(0).clip(0, 100)

    title = "Student Gantt Report"
    export = normalized[["student", "task", "start", "end", "days", "progress"]].rename(
        columns={
            "student": "Student",
            "task": "Task",
            "start": "Start",
            "end": "End",
            "days": "Days",
            "progress": "Progress",
        }
    )
    return _write_report(title, job_dir, [("Gantt", export)], _gantt_html(title, export))


def _make_job_dir(output_root: Path) -> Path:
    job_dir = output_root / uuid4().hex
    job_dir.mkdir(parents=True, exist_ok=True)
    return job_dir


def _write_report(title: str, job_dir: Path, sheets: list[tuple[str, pd.DataFrame]], html: str) -> ReportResult:
    xlsx_path = job_dir / "report.xlsx"
    html_path = job_dir / "report.html"
    pdf_path = job_dir / "report.pdf"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets:
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    html_path.write_text(html, encoding="utf-8")
    _write_pdf(pdf_path, title, sheets)
    return ReportResult(title=title, job_id=job_dir.name, xlsx_path=xlsx_path, html_path=html_path, pdf_path=pdf_path)


def _summary_html(title: str, source_label: str, summary: pd.DataFrame, details: pd.DataFrame | None = None) -> str:
    detail_html = "" if details is None or details.empty else f"<h2>Details</h2>{details.to_html(index=False, escape=True)}"
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{escape(title)}</title>
  <style>{_css()}</style>
</head>
<body>
  <main>
    <h1>{escape(title)}</h1>
    <p class="muted">{escape(source_label)}</p>
    {summary.to_html(index=False, escape=True)}
    {detail_html}
  </main>
</body>
</html>"""


def _gantt_html(title: str, frame: pd.DataFrame) -> str:
    if frame.empty:
        rows = "<p>No valid gantt rows found.</p>"
    else:
        min_start = frame["Start"].min()
        max_end = frame["End"].max()
        total_days = max((max_end - min_start).days + 1, 1)
        rendered = []
        for _, row in frame.sort_values(["Student", "Start", "End"]).iterrows():
            offset = ((row["Start"] - min_start).days / total_days) * 100
            width = max((row["Days"] / total_days) * 100, 2)
            progress = int(row["Progress"])
            rendered.append(
                f"""<div class="gantt-row">
  <div class="label"><strong>{escape(str(row["Student"]))}</strong><span>{escape(str(row["Task"]))}</span></div>
  <div class="track"><div class="bar" style="left:{offset:.2f}%;width:{width:.2f}%"><span>{progress}%</span></div></div>
</div>"""
            )
        rows = "\n".join(rendered)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{escape(title)}</title>
  <style>{_css()}</style>
</head>
<body>
  <main>
    <h1>{escape(title)}</h1>
    <section class="gantt">{rows}</section>
  </main>
</body>
</html>"""


def _write_pdf(path: Path, title: str, sheets: list[tuple[str, pd.DataFrame]]) -> None:
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    for sheet_name, frame in sheets:
        story.append(Paragraph(sheet_name, styles["Heading2"]))
        preview = frame.head(20).fillna("").astype(str)
        data = [list(preview.columns)] + preview.values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#172b4d")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#dcdfe4")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.extend([table, Spacer(1, 12)])
    doc.build(story)


def _key_rows(frame: pd.DataFrame) -> dict[str, tuple[str, ...]]:
    normalized = frame.fillna("").astype(str)
    return {"|".join(row): tuple(row) for row in normalized.to_numpy()}


def _normalize_gantt_columns(frame: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "student": ["student", "學生", "学生", "name", "student name", "氏名"],
        "task": ["task", "任務", "任务", "課題", "作業", "activity"],
        "start": ["start", "start date", "開始", "開始日", "開始日期"],
        "end": ["end", "end date", "結束", "終了", "結束日", "終了日"],
        "progress": ["progress", "進度", "进度", "完成率", "progress %"],
    }
    lowered = {str(column).strip().lower(): column for column in frame.columns}
    selected = {}
    for target, names in aliases.items():
        for name in names:
            if name.lower() in lowered:
                selected[target] = frame[lowered[name.lower()]]
                break
    missing = {"student", "task", "start", "end"} - set(selected)
    if missing:
        raise ValueError(f"Missing required gantt columns: {', '.join(sorted(missing))}")
    return pd.DataFrame(selected)


def _css() -> str:
    return """
body { margin: 0; background: #f7f8f9; color: #172b4d; font-family: Segoe UI, Arial, sans-serif; }
main { width: min(1120px, calc(100vw - 40px)); margin: 32px auto; }
h1 { font-size: 28px; margin: 0 0 8px; }
h2 { font-size: 18px; margin-top: 28px; }
.muted { color: #626f86; }
table { border-collapse: collapse; width: 100%; background: white; border: 1px solid #dcdfe4; }
th, td { border: 1px solid #dcdfe4; padding: 8px 10px; text-align: left; }
th { background: #172b4d; color: white; }
.gantt { display: grid; gap: 10px; }
.gantt-row { display: grid; grid-template-columns: 220px 1fr; gap: 12px; align-items: center; }
.label { background: white; border: 1px solid #dcdfe4; padding: 10px; border-radius: 6px; }
.label span { display: block; color: #626f86; margin-top: 3px; }
.track { position: relative; height: 34px; background: white; border: 1px solid #dcdfe4; border-radius: 6px; overflow: hidden; }
.bar { position: absolute; top: 5px; bottom: 5px; background: #0c66e4; border-radius: 4px; color: white; display: grid; place-items: center; font-size: 12px; }
"""
