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
    assert b"Old file" in response.data
    assert b"New file" in response.data
    assert b'name="report_type" value="compare"' in response.data


def test_compare_upload_returns_download_links(tmp_path, monkeypatch):
    monkeypatch.setattr("app.REPORT_ROOT", tmp_path)
    client = app.test_client()

    response = client.post(
        "/report",
        data={
            "report_type": "compare",
            "old_file": (_excel_bytes(pd.DataFrame({"A": ["old"], "B": ["x"], "C": ["c"], "KEY": ["K1"], "Amount": [1]})), "old.xlsx"),
            "new_file": (_excel_bytes(pd.DataFrame({"A": ["new"], "B": ["x"], "C": ["c"], "KEY": ["K1"], "Amount": [2]})), "new.xlsx"),
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    assert b"report.xlsx" in response.data
    assert b"report.html" in response.data
    assert b"report.pdf" in response.data
