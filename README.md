# SalesPlan Compare

This repository contains two aligned versions:

- EXE package: `releases/exe/MonthlySalesPlanCompareExe.zip`
- Online web app: `app.py` with `src/`, `templates/`, `requirements.txt`, and `render.yaml`

See `VERSION.md` for the repository layout and consistency rule.

## Online Web App

The online app restores the original SalesPlan Compare flow:

1. Upload `Old file` for the previous month's monthly sales plan.
2. Upload `New file` for the current month's monthly sales plan.
3. Run comparison.
4. Review counts on the result screen.
5. Download the generated Excel, HTML, or PDF report.

## Compare Logic

- Header row: row 1
- Data starts from: row 2
- Matching key: Column D
- Rows without a Column D value are ignored.
- Added records are highlighted in yellow on `new data`.
- Removed records are highlighted in gray on `old data`.
- Changed values are highlighted in orange on `new data`.
- New columns in the current month file are highlighted in green.

The generated workbook contains:

- `Summary`
- `new data`
- `old data`

The following fields are excluded from value comparison and shaded gray:

- `BPCS_CUSTOMER_CODE`
- `BPCS_SHIPTO_CODE`
- `CUSTOMER_PN`
- `BUSINESS_CATEGORY_NAME`
- `TRANSACTION_CURRENCY`
- `Ex_Rate_JPY`
- `GENERAL_CODE_1`
- `GENERAL_CODE_2`
- `GENERAL_CODE_3`
- `GENERAL_CODE_4`
- `GENERAL_CODE_5`

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
   - Start command: `gunicorn --workers 1 --threads 1 --timeout 900 app:app`
   - Python version: `3.14.3`

The online app loads the original `compare_core` bytecode extracted from the packaged executable, so Python 3.14 is required.
The longer Gunicorn timeout is needed because large Excel files can take several minutes while openpyxl reads worksheet data on Render Free.

## EXE Download

Download and extract:

```text
releases/exe/MonthlySalesPlanCompareExe.zip
```

Then run:

```text
START_EXE.bat
```

GitHub Pages cannot run this app because it needs a Python backend for Excel processing.

## Test

```powershell
python -m pytest tests -q
```
