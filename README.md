# Excel Hours Viewer & Payroll Formatter

A desktop app for loading driver hours from Excel, cleaning and joining data, and exporting payroll-ready reports. Built with Tkinter, Pandas, and tksheet.

## What It Does
- Loads two Excel files (hours + lunch/break data) and joins them.
- Normalizes key fields (employee ID, dates/times, PRMEAL).
- Calculates regular vs overtime hours (40-hour threshold, Monday–Sunday weeks).
- Provides search, sorting, and editable grid views.
- Exports payroll-friendly Excel files and optional daily/weekly breakdowns.
- Generates a print-friendly PDF of the current view.

## Requirements
- Python 3.10+ (recommended)
- Packages:
  - `pandas`
  - `tksheet`
  - `openpyxl`
  - `matplotlib`

Install:
```bash
pip install pandas tksheet openpyxl matplotlib
```

## Run
```bash
python excel_viewer.py
```

## Inputs
This tool expects two Excel files:
- **File A**: Lunch/break data (e.g., driver, date, break start/finish).
- **File B**: Driver hours data (e.g., date, shift, start/finish, total hours).

The app is fairly forgiving about column names and will attempt to normalize and locate the expected fields.

## Outputs
- **Payroll-ready Excel** with normalized/rounded hour columns.
- **Optional breakdowns**:
  - Daily breakdown
  - Weekly summary (Mon–Sun)
- **PDF snapshot** of the current view (for printing or review).

## Configuration
These are defined near the top of `excel_viewer.py`:
- `EXCLUDED_DRIVER_NAMES`: Full names to exclude from all outputs.
- `EXCLUDED_EMPLOYEE_IDS`: Employee IDs to exclude.
- `EMPLOYEE_NAME_OVERRIDES`: Per-employee first/last name corrections.
- Display limits and UI constants (column widths, fonts, etc.).

## Notes
- The app supports editing the joined table directly in the UI before exporting.
- Sorting and searching operate on the current view.

## Troubleshooting
- If the app opens but looks blank, try reducing `MAX_ROWS_DISPLAY` or `MAX_AUTOSIZE_ROWS`.
- If dates/times look wrong, check the original Excel formatting and column headers.

