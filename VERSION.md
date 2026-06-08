# Version Layout

This repository keeps two deliverable versions of SalesPlan Compare.

## 1. EXE Package

Download package:

```text
releases/exe/MonthlySalesPlanCompareExe.zip
```

The zip contains the original Windows executable distribution:

- `MonthlySalesPlanCompareExe.exe`
- `_internal/`
- `START_EXE.bat`
- `README_EXE.txt`
- Business flow documents

Use this version when users need a local Windows desktop package.

## 2. Online Web App

Deployable program files:

```text
app.py
src/
templates/
requirements.txt
render.yaml
legacy/compare_core.pyc
```

Use this version when users need the app to run in the cloud, for example on Render.

## Consistency Rule

Both versions must stay functionally aligned.

- The EXE package includes the original packaged application.
- The online app directly loads `legacy/compare_core.pyc`, extracted from the same EXE package.
- Changes to comparison behavior must be reflected in both deliverables before release.
- If a new EXE is produced, update `releases/exe/MonthlySalesPlanCompareExe.zip` and refresh `legacy/compare_core.pyc` from that EXE.

The GitHub root is kept deployable for the online web app. The EXE version is stored as a single archive to keep downloads simple and avoid tracking expanded runtime files.
