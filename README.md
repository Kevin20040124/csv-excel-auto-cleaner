# CSV / Excel Auto Cleaner

A small Python tool that cleans messy CSV or Excel files in one command.

## What it fixes

- Removes empty rows and duplicate rows
- Trims extra spaces in headers and cells
- Standardizes empty values (`N/A`, `-`, blanks)
- Lowercases email addresses
- Converts mixed date formats (`2026/10/01`, `2026.10.05`) to `YYYY-MM-DD`
- Converts numbers like `1,200` into real numeric values
- Exports an Excel file with a cleaned sheet and a summary report sheet

## Usage

```
pip install pandas openpyxl

python clean_data.py                 # runs on built-in demo data
python clean_data.py your_file.csv   # cleans your own .csv or .xlsx file
```

Output: `your_file_cleaned.xlsx`

## Example

Before: 6 rows, with duplicates, an empty row, mixed date formats and `1,200`-style numbers.
After: 3 clean rows, plus a report showing what was removed.

## Need this for your own data?

I can adapt the tool to your files and rules (extra columns, custom formats, merging several files).
