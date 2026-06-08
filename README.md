# Excel Report Generator

This repository now contains two versions:

- `MonthlySalesPlanCompareExe.exe` with `_internal/`: the original Windows desktop package.
- `app.py` with `src/` and `templates/`: the online Python web app.

## Online Web App

The web app lets users upload Excel files, choose a report type, and download generated reports as:

- `report.xlsx`
- `report.html`
- `report.pdf`

## Report Types

### Excel Summary

Upload one workbook. The app reports sheet names, row counts, column counts, empty cell counts, and numeric column counts.

### Monthly Compare

Upload a previous-month workbook and a current-month workbook. The app reports added rows, removed rows, and changed rows by comparing whole row values.

### Student Gantt

Upload one workbook with schedule columns. Supported column names include:

- English: `Student`, `Task`, `Start`, `End`, `Progress`
- Chinese: `學生`, `任務`, `開始日`, `結束日`, `進度`
- Japanese: `学生`, `課題`, `開始日`, `終了日`, `完成率`

The app generates an XLSX schedule, an HTML Gantt view, and a PDF summary.

## Run Locally

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open:

```text
http://localhost:5000/
```

## Deploy To Render

1. Push this repository to GitHub.
2. Open Render and create a new Web Service from this repository.
3. Render can use `render.yaml` automatically.
4. If entering settings manually:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app`
   - Python version: `3.11.9`

GitHub Pages cannot run this app because it needs a Python backend for Excel processing.

## Test

```powershell
python -m pytest tests -q
```
