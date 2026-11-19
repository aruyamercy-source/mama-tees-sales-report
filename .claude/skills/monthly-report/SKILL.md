---
name: monthly-report
description: Use when the user wants to clean a messy monthly sales CSV and build the Mama Tee's Bakery report.
---

# Monthly report for Mama Tee's Bakery

Run everything from the project folder. The tool is `report.py`. Branches are Wuse, Garki and Maitama. The 14 products and their prices are listed in CLAUDE.md.

## Steps
1. **Check the input.** The CSV must be in `data/` and have the columns date, branch, product, quantity, unit_price. Never edit or delete it. `data/` is read-only.
2. **Close Excel** if `output/sales-report.xlsx` is open. The script cannot overwrite an open file.
3. **Run the command:**
   - One month: `py report.py data/FILE.csv` (with no file name it uses the newest .csv in `data/`)
   - With a comparison: `py report.py data/FILE.csv --compare data/EARLIER.csv`
   The first file is the current month and the `--compare` file is the earlier month. Month names come from the dates inside the files, not the file names.
4. **What the script does, in order:** remove exact duplicate rows, fix branch names, fix product names, convert all dates to one format (slash dates are read day first), fill blank prices from the price list, send rows it cannot fix to `output/problems.csv` with a reason, remove duplicates again (rows that only matched after cleaning), then build the Excel report and summary.
5. **Confirm after every run** (each line must end in OK):
   - `rows read = clean + duplicates + problems`
   - `branch totals = grand total`
   - `branch x product, weekly, daily and product tables all add up to the grand total`
   - With `--compare` also: the Compare sheet totals, and "every section's totals equal its month's grand total".
   - A MISMATCH stops the run before any file is written. Do not use old outputs; find the cause.
   - Both input files are unchanged: compare size, modified time or `Get-FileHash` before and after.
6. **Outputs**, all in `output/` and overwritten on every run (copy last month's out first if it is needed):
   - `sales-report.xlsx`: 8 sheets, or 9 with `--compare` (a Compare sheet is added and the main sheets show both months)
   - `problems.csv`: the rows that could not be fixed, with a plain reason (with `--compare` it also has a `file` column)
   - `summary.md`: a one-page owner summary: 3 things that went well, 3 to watch, 1 suggestion (plus a comparison section with `--compare`)
7. **Tell the user** the counts (rows read, duplicates, problem rows, clean rows), total sales, and that the checks passed.

## Common problems
- **"Could not write sales-report.xlsx. If it is open in Excel, close it and run again."** Ask the user to close the workbook completely. A hidden `~$sales-report.xlsx` file in `output/` means Excel still has it open. Do not end Excel yourself: it may hold unsaved work.
- **"Input file not found" or "Comparison file not found".** Check the name and path against the files in `data/`.
- **"No .csv files found"** The month's file has not been put in `data/` yet.
- **"The file is missing these columns".** The header row must contain all five column names.
- **"The --compare file ... is not earlier than the current file".** The files are the wrong way round. Name the current month first and the earlier month after `--compare`. Naming the same file twice is also refused.
- **"No usable rows were found".** Every row had a problem. Check the file with the user.
