from pathlib import Path

import pandas as pd

from src.reports import (
    generate_excel_summary_report,
    generate_monthly_compare_report,
    generate_student_gantt_report,
)


def test_excel_summary_report_creates_all_formats(tmp_path):
    source = tmp_path / "sales.xlsx"
    pd.DataFrame({"customer": ["A", "B"], "amount": [100, 250]}).to_excel(source, index=False)

    result = generate_excel_summary_report(source, tmp_path / "out")

    assert result.title == "Excel Summary Report"
    assert result.xlsx_path.exists()
    assert result.html_path.exists()
    assert result.pdf_path.exists()
    assert "sales.xlsx" in result.html_path.read_text(encoding="utf-8")


def test_monthly_compare_report_counts_added_removed_and_changed_rows(tmp_path):
    old_file = tmp_path / "old.xlsx"
    new_file = tmp_path / "new.xlsx"
    pd.DataFrame({"id": [1, 2], "amount": [100, 200]}).to_excel(old_file, index=False)
    pd.DataFrame({"id": [1, 3], "amount": [150, 300]}).to_excel(new_file, index=False)

    result = generate_monthly_compare_report(old_file, new_file, tmp_path / "compare")

    html = result.html_path.read_text(encoding="utf-8")
    assert result.xlsx_path.exists()
    assert "Added rows" in html
    assert "Removed rows" in html
    assert "Changed rows" in html


def test_student_gantt_report_accepts_chinese_columns(tmp_path):
    source = tmp_path / "students.xlsx"
    pd.DataFrame(
        {
            "學生": ["Lee", "Lee", "Chen"],
            "任務": ["Design", "Build", "Review"],
            "開始日": ["2026-06-01", "2026-06-03", "2026-06-02"],
            "結束日": ["2026-06-02", "2026-06-07", "2026-06-04"],
            "進度": [100, 60, 30],
        }
    ).to_excel(source, index=False)

    result = generate_student_gantt_report(source, tmp_path / "gantt")

    html = result.html_path.read_text(encoding="utf-8")
    assert result.xlsx_path.exists()
    assert result.pdf_path.exists()
    assert "Student Gantt Report" in html
    assert "Lee" in html
