---
description: Clean a monthly sales CSV and build the Mama Tee's Bakery report
argument-hint: [current-file.csv] [earlier-file.csv]
---

Build the Mama Tee's Bakery report. Follow the `monthly-report` skill. Arguments: $ARGUMENTS

1. **Pick the command from the arguments.** A bare file name means a file in `data/`.
   - No arguments: run `py report.py` (it uses the newest .csv in `data/`) for a single-month report.
   - One file: run `py report.py data/FILE.csv`.
   - Two files: the first is the current month and the second is the earlier month. Run `py report.py data/CURRENT.csv --compare data/EARLIER.csv`.
   - If a named file is not in `data/`, stop and say so. Do not guess.
2. **Make sure Excel is closed.** If an EXCEL process is running or `output/~$sales-report.xlsx` exists, ask me to close the report and wait. Do not end Excel yourself.
3. **Refresh PATH, then run the command** from the project folder. Before running, reload PATH from the Machine and User environment variables so plain `py` works in this shell:
   `$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')`
   Run it in the same PowerShell call as the report command. Never edit or delete anything in `data/`.
4. **Confirm in 5 lines or fewer:**
   - rows read
   - duplicates removed
   - rows sent to `problems.csv`
   - whether every check says OK (with `--compare`, say both months passed)
   - where the files are: `output/sales-report.xlsx`, `output/problems.csv`, `output/summary.md`

If a check says MISMATCH or the run fails, say what happened in one or two lines and do not claim success.
