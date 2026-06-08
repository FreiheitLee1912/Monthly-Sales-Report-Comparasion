MonthlySalesPlanCompareExe v1.6

How to use:
1. Double-click START_EXE.bat.
2. Open http://localhost:5002/ in a browser.
3. Upload the previous month file and current month file, then run comparison.

Notes:
- Python installation is not required.
- Do not delete MonthlySalesPlanCompareExe.exe or the _internal folder.
- Result files are saved under each user's local folder:
  %LOCALAPPDATA%\MonthlySalesPlanCompareExe\outputs
- v1.6 excludes the following columns from value comparison and shades them grey in old data, new data, and Comparison:
  BPCS_CUSTOMER_CODE, BPCS_SHIPTO_CODE, CUSTOMER_PN, BUSINESS_CATEGORY_NAME,
  TRANSACTION_CURRENCY, Ex_Rate_JPY, GENERAL_CODE_1, GENERAL_CODE_2,
  GENERAL_CODE_3, GENERAL_CODE_4, GENERAL_CODE_5.
