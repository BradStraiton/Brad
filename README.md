# Cash Flow Spreadsheet App

This app reads data directly from the attached Excel file (`CASH FLOW SOURCE-APRIL26.xlsx`) and shows it in a browser UI.

## Run

```bash
python app.py
```

Then open:

- http://127.0.0.1:8000

## How it works

- `/api/data` parses the Excel workbook using only Python standard library modules (`zipfile` + XML parser).
- The frontend calls `/api/data`, lets you select sheets, and filter rows.
- Every API request re-reads the spreadsheet so updates in Excel are reflected after refresh.
