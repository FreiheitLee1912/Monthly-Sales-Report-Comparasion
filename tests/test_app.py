import io

import pandas as pd

from app import app


def _excel_bytes(frame: pd.DataFrame) -> io.BytesIO:
    stream = io.BytesIO()
    frame.to_excel(stream, index=False)
    stream.seek(0)
    return stream


def test_home_page_loads_report_selector():
    client = app.test_client()

    response = client.get("/")

    assert response.status_code == 200
    assert b"Excel Summary" in response.data
    assert b"Student Gantt" in response.data


def test_summary_upload_returns_download_links(tmp_path, monkeypatch):
    monkeypatch.setattr("app.REPORT_ROOT", tmp_path)
    client = app.test_client()

    response = client.post(
        "/report",
        data={
            "report_type": "summary",
            "excel_file": (_excel_bytes(pd.DataFrame({"amount": [1, 2]})), "source.xlsx"),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    assert b"report.xlsx" in response.data
    assert b"report.html" in response.data
    assert b"report.pdf" in response.data
