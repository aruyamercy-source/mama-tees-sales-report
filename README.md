# Mama Tee's Bakery – Monthly Sales Report

Mama Tee's Bakery is a made-up bakery with three branches in Abuja: Wuse, Garki and Maitama. This tool takes a messy monthly sales CSV and turns it into a clean report.

## What it does

- Cleans the CSV: removes duplicate rows, fixes branch and product names, makes dates consistent and fills blank prices from the price list.
- Builds an Excel report.
- Writes a list of rows it could not fix, with the reason for each.
- Writes a plain-English summary for the owner.

## How to run it

Put the month's CSV in `data/`, then run one of these:

```
py report.py data/FILE.csv
```

For a two-month comparison, give the current month first and the earlier month after `--compare`:

```
py report.py data/CURRENT.csv --compare data/EARLIER.csv
```

In Claude Code, you can run `/report` instead (for example `/report data/sales-october.csv data/sales-september.csv`).

## What is in `output/`

- `sales-report.xlsx` – the Excel report (with a Compare sheet when you use `--compare`)
- `problems.csv` – rows that could not be fixed, with a reason
- `summary.md` – a one-page summary for the owner

These files are overwritten each time you run the tool. Close the Excel report before running it again.

## The one rule

The original CSV files in `data/` are never changed or deleted. All results go to `output/`.
