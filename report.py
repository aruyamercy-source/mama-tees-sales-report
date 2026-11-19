"""Mama Tee's Bakery - monthly sales report.

Usage:
    py report.py                                   newest .csv in data/
    py report.py data/sales-october.csv            a named file
    py report.py data/sales-october.csv --compare data/sales-september.csv
                                                   also compare with an earlier month

The script only READS the input files and only WRITES to output/:
    output/problems.csv      rows that could not be fixed, with a reason
    output/sales-report.xlsx the Excel report (8 sheets; with --compare the main sheets
                             show both months and a Compare sheet is added)
    output/summary.md        plain-English summary for the owner
Each run overwrites these three files.
"""
import argparse
import calendar
import difflib
import math
import re
import sys
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

# --------------------------------------------------------------------------
# Settings. These match CLAUDE.md. Change a price here and it applies everywhere.
# --------------------------------------------------------------------------
BRANCHES = ["Wuse", "Garki", "Maitama"]
PRODUCTS = {  # product name: price in naira
    "Plain Bread": 1500,
    "Sliced Bread": 2000,
    "Chin Chin": 1200,
    "Puff Puff": 600,
    "Meat Pie": 1000,
    "Sausage Roll": 800,
    "Fish Roll": 700,
    "Buns": 300,
    "Doughnut": 500,
    "Cake Slice": 700,
    "Croissant": 1200,
    "Cinnamon Roll": 1000,
    "Banana Bread": 1800,
    "Chocolate Chunk Cookies": 600,
}
MATCH_CUTOFF = 0.75  # how close a spelling must be to count as the same name

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "output"
PROBLEMS_FILE = OUT_DIR / "problems.csv"
SUMMARY_FILE = OUT_DIR / "summary.md"
EXCEL_FILE = OUT_DIR / "sales-report.xlsx"

INPUT_COLUMNS = ["date", "branch", "product", "quantity", "unit_price"]
CLEAN_KEY = ["date", "branch", "product", "quantity", "unit_price"]
PROBLEM_COLUMNS = ["row_number", "date", "branch", "product", "quantity", "unit_price", "reason"]
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]
MONTH_LOOKUP = {name[:3].lower(): i + 1 for i, name in enumerate(MONTH_NAMES)}
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def naira(amount):
    return f"₦{amount:,.0f}"


def signed_naira(amount):
    sign = "+" if amount > 0 else "-" if amount < 0 else ""
    return f"{sign}{naira(abs(amount))}"


def signed_pct(fraction):
    return f"{fraction * 100:+.1f}%"


def fmt_date(d):
    return f"{d.day} {MONTH_NAMES[d.month - 1][:3]} {d.year}"


def short_date(d):
    return f"{d.day} {MONTH_NAMES[d.month - 1][:3]}"


def pct(part, whole):
    return part / whole * 100 if whole else 0.0


def output_path(path):
    """Only ever allow writing inside output/."""
    resolved = Path(path).resolve()
    if resolved.parent != OUT_DIR.resolve():
        raise SystemExit(f"Refusing to write outside output/: {resolved}")
    return resolved


def parse_args():
    parser = argparse.ArgumentParser(
        description="Mama Tee's Bakery monthly sales report. Reads from data/, writes to output/.")
    parser.add_argument("current", nargs="?",
                        help="this month's CSV file (default: the newest .csv in data/)")
    parser.add_argument("--compare", metavar="EARLIER_CSV",
                        help="an earlier month's CSV file to compare this month with")
    return parser.parse_args()


def find_input(arg=None):
    if arg:
        path = Path(arg)
        if not path.is_file():
            raise SystemExit(f"Input file not found: {path}")
        return path
    files = sorted(DATA_DIR.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise SystemExit(f"No .csv files found in {DATA_DIR}. Put the month's file there.")
    return files[0]


def read_input(path):
    """Read the CSV (read-only) as plain text so nothing is changed on the way in."""
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SystemExit(f"Could not read {path} as text.")
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in INPUT_COLUMNS if c not in df.columns]
    if missing:
        raise SystemExit(f"The file is missing these columns: {', '.join(missing)}")
    df = df[INPUT_COLUMNS].copy()
    df.insert(0, "row_number", range(1, len(df) + 1))  # 1 = first row under the header
    return df


# --------------------------------------------------------------------------
# Cleaning
# --------------------------------------------------------------------------
def norm(text):
    text = re.sub(r"[-_]", " ", text.strip().lower())
    return re.sub(r"\s+", " ", text).strip()


BRANCH_LOOKUP = {norm(b): b for b in BRANCHES}
PRODUCT_LOOKUP = {norm(p): p for p in PRODUCTS}


def match_name(raw, lookup):
    """Return the correct name for a messy one, or None if there is no close match."""
    key = norm(raw)
    if not key:
        return None
    if key in lookup:
        return lookup[key]
    close = difflib.get_close_matches(key, list(lookup), n=1, cutoff=MATCH_CUTOFF)
    return lookup[close[0]] if close else None


def parse_date(raw):
    """Return (date, None) or (None, reason). Slash dates are read day first."""
    text = raw.strip()
    if not text:
        return None, "Date is blank"
    if g := re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", text):
        y, m, d = int(g[1]), int(g[2]), int(g[3])
    elif g := re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text):
        d, m, y = int(g[1]), int(g[2]), int(g[3])
    elif g := re.fullmatch(r"(\d{1,2}) ([A-Za-z]+) (\d{4})", text):
        d, m, y = int(g[1]), MONTH_LOOKUP.get(g[2][:3].lower()), int(g[3])
    elif g := re.fullmatch(r"([A-Za-z]+) (\d{1,2}),? (\d{4})", text):
        m, d, y = MONTH_LOOKUP.get(g[1][:3].lower()), int(g[2]), int(g[3])
    else:
        return None, f"Date '{text}' is not in a known format"
    try:
        return date(y, m, d), None
    except (ValueError, TypeError):
        return None, f"Date '{text}' is not a real date"


def parse_number(raw):
    text = raw.strip().replace(",", "").replace("₦", "").replace(" ", "")
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def clean_rows(df):
    """Fix what can be fixed. Returns (clean rows, problem rows)."""
    good, problems = [], []
    for row in df.itertuples(index=False):
        reasons = []

        raw_branch = row.branch.strip()
        branch = match_name(row.branch, BRANCH_LOOKUP)
        if branch is None:
            reasons.append("Branch is blank" if not raw_branch
                           else f"Branch '{raw_branch}' is not recognised")

        raw_product = row.product.strip()
        product = match_name(row.product, PRODUCT_LOOKUP)
        if product is None:
            reasons.append("Product is blank" if not raw_product
                           else f"Product '{raw_product}' is not on the product list")

        day, date_reason = parse_date(row.date)
        if date_reason:
            reasons.append(date_reason)

        qty = parse_number(row.quantity)
        if qty is None or qty <= 0 or qty != int(qty):
            reasons.append("Quantity is missing or not a whole number above zero")

        price = parse_number(row.unit_price)
        price_filled = False
        if not row.unit_price.strip():
            if product is not None:
                price = PRODUCTS[product]  # blank price: use the price list
                price_filled = True
        elif price is None or price <= 0:
            reasons.append(f"Unit price '{row.unit_price.strip()}' is not a positive number")

        if reasons:
            problems.append({
                "row_number": row.row_number, "date": row.date, "branch": row.branch,
                "product": row.product, "quantity": row.quantity,
                "unit_price": row.unit_price, "reason": "; ".join(reasons),
            })
        else:
            good.append({
                "date": day, "branch": branch, "product": product,
                "quantity": int(qty), "unit_price": price,
                "branch_fixed": row.branch != branch,
                "product_fixed": row.product != product,
                "date_fixed": row.date.strip() != day.isoformat(),
                "price_filled": price_filled,
            })

    clean_cols = ["date", "branch", "product", "quantity", "unit_price",
                  "branch_fixed", "product_fixed", "date_fixed", "price_filled"]
    return pd.DataFrame(good, columns=clean_cols), pd.DataFrame(problems, columns=PROBLEM_COLUMNS)


# --------------------------------------------------------------------------
# Excel report: purple theme
# --------------------------------------------------------------------------
BG = "E9E3F7"        # light purple report background
HEAD = "5B3FA0"      # darker purple header rows (and the current month's title bars)
CHART_BG = "F3EFFB"  # light purple chart background
BAR = "7D5FC4"       # main bar / line colour
BAR_LIGHT = "B9A5E3"
GRID = "D9D6E3"      # light chart gridlines
EDGE = "C9BEE8"      # thin table borders
EARLIER_BAR = "9A83D6"        # lighter purple: title bars and charts for the earlier month
EARLIER_BAR_LIGHT = "D5CBEF"  # lighter still: the earlier month's "worst" bars
COMPARE_BAR = "8A3FA0"        # third purple: comparison blocks

MONEY = '"₦"#,##0'
UNITS = "#,##0"
PERCENT = "0.0%"
CHANGE_MONEY = '+"₦"#,##0;-"₦"#,##0;"₦"0'
CHANGE_UNITS = "+#,##0;-#,##0;0"
CHANGE_PERCENT = "+0.0%;-0.0%;0.0%"

BG_FILL = PatternFill("solid", fgColor=BG)
BORDER = Border(left=Side(style="thin", color=EDGE), right=Side(style="thin", color=EDGE),
                top=Side(style="thin", color=EDGE), bottom=Side(style="thin", color=EDGE))
ROW_CM = 0.53  # height of one default row in cm, used to place captions under charts
COL_CM = 1.7   # width of one default column in cm

# Columns of the Clean data sheet. One month alone uses A-F. With --compare the current
# month uses A-F and the earlier month H-M.
CLEAN_KEYS = ["date", "branch", "product", "qty", "price", "sales"]
SINGLE_COLS = dict(zip(CLEAN_KEYS, "ABCDEF"))
EARLIER_COLS = dict(zip(CLEAN_KEYS, "HIJKLM"))


def paint(ws, rows, cols, start=1):
    """Light purple background block, and no gridlines."""
    ws.sheet_view.showGridLines = False
    for r in range(start, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c).fill = BG_FILL


def put(ws, row, col, value, fmt=None, bold=False, italic=False, color=None, size=None,
        fill=None, border=False, wrap=False, align=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fmt:
        cell.number_format = fmt
    if bold or italic or color or size:
        cell.font = Font(bold=bold, italic=italic, color=color, size=size)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if wrap or align:
        cell.alignment = Alignment(wrap_text=wrap, vertical="top" if wrap else None,
                                   horizontal=align)
    if border:
        cell.border = BORDER
    return cell


def sheet_title(ws, text):
    put(ws, 1, 1, text, bold=True, color=HEAD, size=14)


def title_bar(ws, row, text, colour, ncols, first_col=1):
    """A full-width coloured title bar with white bold text."""
    ws.merge_cells(start_row=row, start_column=first_col, end_row=row,
                   end_column=first_col + ncols - 1)
    for c in range(first_col, first_col + ncols):
        ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=colour)
    cell = ws.cell(row=row, column=first_col, value=text)
    cell.font = Font(bold=True, color="FFFFFF", size=13)
    cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def heading(ws, d, top, text, ncols, plain=None):
    """Plain title when one month is shown; a coloured bar naming the month with --compare."""
    if d.get("bar"):
        title_bar(ws, top, f"{text} - {d['month_title']}", d["bar"], ncols)
    else:
        sheet_title(ws, plain or text)


def chart_title(d, text):
    return f"{text} - {d['month_title']}" if d.get("bar") else text


def clean_ref(d, key):
    """Where this month's data lives on the Clean data sheet: the whole column when one
    month is shown, only this month's own block with --compare (so months never mix)."""
    block = d.get("block")
    if block is None:
        col = SINGLE_COLS[key]
        return f"'Clean data'!${col}:${col}"
    col = block["cols"][key]
    return f"'Clean data'!${col}${block['first']}:${col}${block['last']}"


def table(ws, top, left, headers, rows, formats, total_last=False):
    """Purple header row, white body. Returns (first body row, last body row)."""
    for j, header in enumerate(headers):
        put(ws, top, left + j, header, bold=True, color="FFFFFF", fill=HEAD, border=True,
            align="center")
    for i, row in enumerate(rows, start=1):
        is_total = total_last and i == len(rows)
        for j, value in enumerate(row):
            put(ws, top + i, left + j, value, fmt=formats[j], fill="FFFFFF", border=True,
                bold=is_total)
    return top + 1, top + len(rows)


def set_widths(ws, widths):
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def note(ws, col, row, ncols, nrows, text):
    """A wrapped text block in a merged area."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row + nrows - 1,
                   end_column=col + ncols - 1)
    put(ws, row, col, text, italic=True, wrap=True, color="4A3A7A")


def style_chart(chart, title, x_title, y_title, plot, y_format=None, x_format=None):
    """Titles that do not overlap, light gridlines, light purple background."""
    chart.title = title
    chart.title.overlay = False
    chart.x_axis.title = x_title
    chart.x_axis.title.overlay = False
    chart.y_axis.title = y_title
    chart.y_axis.title.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend = None
    if y_format:
        chart.y_axis.number_format = y_format
    if x_format:
        chart.x_axis.number_format = x_format
    for axis in (chart.x_axis, chart.y_axis):
        if axis.majorGridlines is not None:
            axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID, w=6350))
    chart.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge",
        x=plot[0], y=plot[1], w=plot[2], h=plot[3]))
    chart.graphical_properties = GraphicalProperties(solidFill=CHART_BG,
                                                     ln=LineProperties(solidFill=EDGE))
    chart.plot_area.graphicalProperties = GraphicalProperties(solidFill=CHART_BG)


def place_chart(ws, chart, anchor, width, height, sentence):
    """Put the chart on the sheet with one plain sentence directly under it.
    Returns the last row used by the sentence."""
    chart.width, chart.height = width, height
    ws.add_chart(chart, anchor)
    col_letter, row = coordinate_from_string(anchor)
    col = column_index_from_string(col_letter)
    caption_row = row + math.ceil(height / ROW_CM) + 1
    note(ws, col, caption_row, math.ceil(width / COL_CM), 2, sentence)
    return caption_row + 1


def colour_series(series, colour):
    series.graphicalProperties.solidFill = colour
    series.graphicalProperties.line.solidFill = colour


# ---- sheet 1: Overview ----------------------------------------------------
def overview_section(ws, d, top):
    """One month's totals and highlights. Returns the last row used."""
    paint(ws, top + 21, 8, start=top)
    heading(ws, d, top, "Sales overview", 4, plain=f"Sales overview - {d['month_title']}")
    comp = d.get("compare")
    block = d.get("block")
    clean_rows_formula = ("=COUNTA('Clean data'!$A:$A)-1" if block is None
                          else f"=COUNTA({clean_ref(d, 'date')})")
    rows = [
        ["Total sales", f"=SUM({clean_ref(d, 'sales')})"],
        ["Items sold", f"=SUM({clean_ref(d, 'qty')})"],
        ["Clean rows", clean_rows_formula],
    ]
    if comp:  # one plain-text line: this month against the earlier month
        total = comp["total"]
        text = signed_naira(total["change"])
        if total["pct"] is not None:
            text += f" ({signed_pct(total['pct'])})"
        rows.append([f"Total sales vs {comp['earlier_title']}", text])
    table(ws, top + 2, 1, ["Measure", "Result"], rows, [None, None])
    ws[f"B{top + 3}"].number_format = MONEY
    ws[f"B{top + 4}"].number_format = UNITS
    ws[f"B{top + 5}"].number_format = UNITS
    table(ws, top + 8, 1, ["Highlight", "Name", "Sales (₦)"], d["highlights"],
          [None, None, MONEY])
    if comp:
        note_text = ("Best and worst day only count days that had sales. "
                     "The first three totals are Excel formulas: click a cell to see how it is "
                     "calculated. The comparison line is a plain value.")
    elif block is not None:
        note_text = ("Best and worst day only count days that had sales. The first three "
                     "totals are Excel formulas that read only this month's block on the "
                     "Clean data sheet.")
    else:
        note_text = ("Best and worst day only count days that had sales. "
                     "Totals above are Excel formulas: click a cell to see how it is calculated.")
    note(ws, 1, top + 16, 3, 3, note_text)
    return top + 18


def overview_comparison(ws, top, d_cur, d_earlier):
    """The Comparison block: both months side by side."""
    c = d_cur["compare"]
    paint(ws, top + 16, 8, start=top)
    title_bar(ws, top, f"Comparison - {d_cur['month_title']} vs {d_earlier['month_title']}",
              COMPARE_BAR, 4)
    rows = [
        ["Total sales", d_earlier["grand_sales"], d_cur["grand_sales"],
         d_cur["grand_sales"] - d_earlier["grand_sales"]],
        ["Items sold", d_earlier["grand_units"], d_cur["grand_units"],
         d_cur["grand_units"] - d_earlier["grand_units"]],
    ]
    for (label, e_name, e_amount), (_, c_name, c_amount) in zip(d_earlier["highlights"],
                                                                 d_cur["highlights"]):
        verdict = "same" if e_name == c_name else "name changed"
        rows.append([label, f"{e_name} - {naira(e_amount)}", f"{c_name} - {naira(c_amount)}",
                     f"{signed_naira(c_amount - e_amount)} ({verdict})"])
    first, last = table(ws, top + 2, 1, ["Metric", c["earlier_label"], c["cur_label"], "Change"],
                        rows, [None, MONEY, MONEY, CHANGE_MONEY])
    for col, fmt in ((2, UNITS), (3, UNITS), (4, CHANGE_UNITS)):  # the "Items sold" row
        ws.cell(row=first + 1, column=col).number_format = fmt
    note(ws, 1, last + 2, 4, 3,
         f"Change = {d_cur['month_title']} minus {d_earlier['month_title']}. For the best and "
         "worst rows each month shows the name and its sales, and Change is the difference "
         "between those two sales amounts, marked (same) or (name changed).")
    return last + 4


def sheet_overview(ws, d, d_earlier=None):
    ws.title = "Overview"
    end = overview_section(ws, d, 1)
    if d_earlier:
        end = overview_section(ws, d_earlier, end + 3)
        overview_comparison(ws, end + 3, d, d_earlier)
        set_widths(ws, [28, 30, 30, 26])
    else:
        set_widths(ws, [28, 30, 18])


# ---- sheet 2 (comparison mode only): Compare --------------------------------
def sheet_compare(ws, d):
    c = d["compare"]
    cur, earlier = c["cur_title"], c["earlier_title"]
    paint(ws, 40, 24)
    sheet_title(ws, f"{cur} vs {earlier}")
    put(ws, 2, 1, f"Change = {cur} minus {earlier}. Percent change is measured against "
                  f"{earlier}.", italic=True, color="4A3A7A")
    headers = ["", earlier, cur, "Change (₦)", "Change (%)"]
    formats = [None, MONEY, MONEY, CHANGE_MONEY, CHANGE_PERCENT]

    def rows_for(items, total_label):
        rows = [[i["name"], i["earlier"], i["current"], i["change"],
                 i["pct"] if i["pct"] is not None else "n/a"] for i in items]
        e = sum(i["earlier"] for i in items)
        n = sum(i["current"] for i in items)
        rows.append([total_label, e, n, n - e, (n - e) / e if e else "n/a"])
        return rows

    put(ws, 3, 1, "Sales per branch", bold=True, color=HEAD)
    first, last = table(ws, 4, 1, ["Branch"] + headers[1:], rows_for(c["branches"], "All branches"),
                        formats, total_last=True)
    branch_last = last - 1  # last branch row (the row after it is the total)

    put(ws, last + 2, 1, "Sales per product", bold=True, color=HEAD)
    table(ws, last + 3, 1, ["Product"] + headers[1:], rows_for(c["products"], "All products"),
          formats, total_last=True)

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=4, max_row=branch_last),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=branch_last))
    colour_series(chart.series[0], BAR_LIGHT)  # earlier month, light
    colour_series(chart.series[1], HEAD)       # current month, dark
    chart.gapWidth = 80
    style_chart(chart, f"Sales per branch: {c['earlier_label']} vs {c['cur_label']}",
                "Branch", "Sales (₦)", plot=(0.17, 0.16, 0.78, 0.50), y_format=MONEY)
    chart.legend = Legend()
    chart.legend.position = "b"
    chart.legend.overlay = False
    place_chart(ws, chart, "G3", 17, 9.5,
                f"This chart compares each branch's total sales in {c['earlier_label']} and in "
                f"{c['cur_label']}; taller bars mean more sales.")
    set_widths(ws, [26, 16, 16, 16, 13])


# ---- Branch totals -----------------------------------------------------------
def sheet_branches(ws, d, top=1):
    paint(ws, top + 29, 18, start=top)
    heading(ws, d, top, "Sales per branch", 15)
    first_row = top + 3
    total_row = first_row + len(BRANCHES)
    rows = []
    for i, branch in enumerate(BRANCHES):
        r = first_row + i
        rows.append([
            branch,
            f"=SUMIFS({clean_ref(d, 'qty')},{clean_ref(d, 'branch')},$A{r})",
            f"=SUMIFS({clean_ref(d, 'sales')},{clean_ref(d, 'branch')},$A{r})",
            f"=IF($C${total_row}=0,0,C{r}/$C${total_row})",
        ])
    rows.append(["Grand Total", f"=SUM({clean_ref(d, 'qty')})", f"=SUM({clean_ref(d, 'sales')})",
                 f"=SUM(D{first_row}:D{total_row - 1})"])
    table(ws, top + 2, 1, ["Branch", "Units sold", "Sales (₦)", "Share of sales"], rows,
          [None, UNITS, MONEY, PERCENT], total_last=True)
    note(ws, 1, total_row + 2, 4, 4,
         "The Grand Total is worked out directly from the clean rows, not by adding up the "
         "branch rows. Click any number to see its formula.")

    chart = BarChart()
    chart.type = "col"
    chart.add_data(Reference(ws, min_col=3, min_row=top + 2, max_row=total_row - 1),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first_row, max_row=total_row - 1))
    colour_series(chart.series[0], d["chart_colour"])
    chart.gapWidth = 80
    style_chart(chart, chart_title(d, "Sales per branch"), "Branch", "Sales (₦)",
                plot=(0.17, 0.16, 0.78, 0.60), y_format=MONEY)
    end = place_chart(ws, chart, f"F{top + 2}", 16, 8.5,
                      "This chart compares total sales for each branch (the Grand Total is "
                      "not shown).")
    set_widths(ws, [16, 14, 16, 16])
    return end


# ---- Products -----------------------------------------------------------------
def sheet_products(ws, d, top=1):
    paint(ws, top + 33, 22, start=top)
    heading(ws, d, top, "Best and worst selling products", 17)
    rows = []
    for rank, row in enumerate(d["best"].itertuples(index=False), start=1):
        rows.append(["Best", rank, row.product, int(row.units), float(row.sales)])
    for rank, row in enumerate(d["worst"].itertuples(index=False), start=1):
        rows.append(["Worst", rank, row.product, int(row.units), float(row.sales)])
    first, last = table(ws, top + 2, 1,
                        ["Group", "Rank", "Product", "Units sold", "Sales (₦)"],
                        rows, [None, None, None, UNITS, MONEY])

    chart = BarChart()
    chart.type = "bar"  # horizontal bars
    chart.add_data(Reference(ws, min_col=5, min_row=top + 2, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=3, min_row=first, max_row=last))
    colour_series(chart.series[0], d["chart_colour"])
    for i in range(len(d["best"]), len(rows)):  # lighter bars for the worst sellers
        point = DataPoint(idx=i)
        point.graphicalProperties = GraphicalProperties(solidFill=d["worst_colour"])
        chart.series[0].dPt.append(point)
    chart.gapWidth = 60
    chart.x_axis.scaling.orientation = "maxMin"  # best seller at the top
    chart.y_axis.crosses = "max"                 # keep the sales axis at the bottom
    style_chart(chart, chart_title(d, "Best and worst selling products"), "Product",
                "Sales (₦)", plot=(0.28, 0.14, 0.67, 0.64), y_format=MONEY)
    end = place_chart(ws, chart, f"G{top + 2}", 18, 9.5,
                      "This chart shows the 5 best sellers (dark bars) and the 5 worst sellers "
                      "(light bars) by total sales.")
    set_widths(ws, [9, 8, 26, 13, 16])
    return end


# ---- Branch x Product ----------------------------------------------------------
def sheet_branch_product(ws, d, top=1):
    paint(ws, top + 43, 10, start=top)
    heading(ws, d, top, "Sales per product per branch", 5)
    pivot, shares = d["pivot"], d["shares"]
    n = len(PRODUCTS)

    put(ws, top + 2, 1, "Sales (₦) per product per branch", bold=True, color=HEAD)
    rows = []
    for product in PRODUCTS:
        values = [float(pivot.loc[product, b]) for b in BRANCHES]
        rows.append([product, *values, sum(values)])
    totals = [float(pivot[b].sum()) for b in BRANCHES]
    rows.append(["All products", *totals, sum(totals)])
    first, last = table(ws, top + 3, 1, ["Product", *BRANCHES, "All branches"], rows,
                        [None, MONEY, MONEY, MONEY, MONEY], total_last=True)
    heat = d.get("heat_range")
    if heat:  # comparison mode: both months share one colour range
        rule = ColorScaleRule(start_type="num", start_value=heat[0], start_color="F3EEFB",
                              end_type="num", end_value=heat[1], end_color="8466CC")
    else:
        rule = ColorScaleRule(start_type="min", start_color="F3EEFB",
                              end_type="max", end_color="8466CC")
    ws.conditional_formatting.add(f"B{first}:D{first + n - 1}", rule)

    top2 = last + 3
    put(ws, top2 - 1, 1, "Each product as a share of its branch's sales", bold=True, color=HEAD)
    share_rows = []
    for product in PRODUCTS:
        share_rows.append([product, *[float(shares.loc[product, b]) for b in BRANCHES]])
    share_rows.append(["All products", *[float(shares[b].sum()) for b in BRANCHES]])
    _, last2 = table(ws, top2, 1, ["Product", *BRANCHES], share_rows,
                     [None, PERCENT, PERCENT, PERCENT], total_last=True)
    text = "Darker purple means higher sales. Each branch column in the second table adds up to 100%."
    if heat:
        text = ("Darker purple means higher sales. Both months use the same colour scale, so the "
                "same colour means the same amount. Each branch column in the second table adds "
                "up to 100%.")
    note(ws, 1, last2 + 2, 5, 2, text)
    set_widths(ws, [26, 15, 15, 15, 16])
    return last2 + 3


# ---- Weekly --------------------------------------------------------------------
def sheet_weekly(ws, d, top=1):
    paint(ws, top + 29, 18, start=top)
    heading(ws, d, top, "Sales by week", 17)
    rows = [[w["label"], w["start"], w["end"], w["units"], w["sales"], w["note"]]
            for w in d["weekly"]]
    first, last = table(ws, top + 2, 1, ["Week", "Week starting (Mon)", "Week ending (Sun)",
                                         "Units sold", "Sales (₦)", "Note"],
                        rows, [None, "yyyy-mm-dd", "yyyy-mm-dd", UNITS, MONEY, None])

    chart = BarChart()
    chart.type = "col"
    chart.add_data(Reference(ws, min_col=5, min_row=top + 2, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    colour_series(chart.series[0], d["chart_colour"])
    chart.gapWidth = 60
    style_chart(chart, chart_title(d, "Sales by week"), "Week (Monday to Sunday)",
                "Sales (₦)", plot=(0.17, 0.16, 0.78, 0.60), y_format=MONEY)
    end = place_chart(ws, chart, f"H{top + 2}", 16, 8.5,
                      "This chart shows total sales for each week of the month. "
                      "The first and last weeks may be part-weeks.")
    set_widths(ws, [18, 19, 18, 12, 16, 12])
    return end


# ---- Daily sales (current month only) -----------------------------------------
def sheet_daily(ws, d):
    paint(ws, 44, 26)
    sheet_title(ws, "Sales by day")
    rows = [[x["label"], x["date"], x["weekday"], x["units"], x["sales"]] for x in d["daily"]]
    first, last = table(ws, 3, 1, ["Day", "Date", "Weekday", "Units sold", "Sales (₦)"],
                        rows, [None, "yyyy-mm-dd", None, UNITS, MONEY])

    chart = LineChart()
    chart.add_data(Reference(ws, min_col=5, min_row=3, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    series = chart.series[0]
    series.graphicalProperties.line.solidFill = BAR
    series.graphicalProperties.line.width = 22000
    series.marker = Marker(symbol="circle", size=5)
    series.marker.graphicalProperties = GraphicalProperties(solidFill=BAR)
    series.marker.graphicalProperties.line.solidFill = BAR
    series.smooth = False
    chart.x_axis.tickLblSkip = 2
    style_chart(chart, "Sales by day", "Day", "Sales (₦)",
                plot=(0.10, 0.14, 0.86, 0.62), y_format=MONEY)
    place_chart(ws, chart, "G3", 24, 9.5,
                "This chart shows total sales for each day of the month, "
                "so busy and quiet days stand out.")
    set_widths(ws, [10, 13, 10, 12, 16])


# ---- Price vs Units (current month only) --------------------------------------
def sheet_price_units(ws, d):
    paint(ws, 40, 30)
    sheet_title(ws, "Price vs units sold")
    prod = d["product_tbl"].set_index("product")
    ranked = sorted(PRODUCTS, key=lambda p: (-int(prod.loc[p, "units"]), p))
    rows = [[rank, name, PRODUCTS[name], int(prod.loc[name, "units"]),
             float(prod.loc[name, "sales"])] for rank, name in enumerate(ranked, start=1)]
    first, last = table(ws, 3, 1, ["#", "Product", "Unit price (₦)", "Units sold",
                                   "Sales (₦)"],
                        rows, [None, None, MONEY, UNITS, MONEY])

    prices = [r[2] for r in rows]
    units = [r[3] for r in rows]
    spread = max(units) - min(units)
    pad = max(spread * 0.15, 5)
    y_min = max(0, math.floor((min(units) - pad) / 5) * 5)   # zoomed to where the dots sit
    y_max = math.ceil((max(units) + pad) / 5) * 5
    x_max = math.ceil(max(prices) * 1.15 / 500) * 500

    # Spread the labels out: sort by price, then give neighbours different positions in turn
    positions = ["t", "b", "r", "l"]
    label_pos = {}
    for turn, idx in enumerate(sorted(range(len(rows)), key=lambda i: (rows[i][2], rows[i][3]))):
        label_pos[idx] = positions[turn % len(positions)]

    chart = ScatterChart()
    chart.style = 13
    chart.scatterStyle = "lineMarker"
    for i, r in enumerate(range(first, last + 1)):  # one series (one dot) per product
        series = Series(Reference(ws, min_col=4, min_row=r), Reference(ws, min_col=3, min_row=r),
                        title=ws.cell(row=r, column=2).value)
        series.marker = Marker(symbol="circle", size=10)
        series.marker.graphicalProperties = GraphicalProperties(solidFill=BAR)
        series.marker.graphicalProperties.line.solidFill = HEAD
        series.graphicalProperties.line.noFill = True
        series.dLbls = DataLabelList()
        series.dLbls.showSerName = True
        series.dLbls.showVal = False
        series.dLbls.showCatName = False
        series.dLbls.showLegendKey = False
        series.dLbls.showPercent = False
        series.dLbls.position = label_pos[i]
        chart.series.append(series)
    style_chart(chart, "Unit price vs units sold", "Unit price (₦)", "Units sold (items)",
                plot=(0.07, 0.10, 0.89, 0.72), x_format=MONEY)
    chart.x_axis.scaling.min = 0
    chart.x_axis.scaling.max = x_max
    chart.y_axis.scaling.min = y_min
    chart.y_axis.scaling.max = y_max
    sentence = ("Each dot is one product: further right means a higher price, higher up means "
                "more items sold.")
    if y_min > 0:
        sentence += (f" The units axis starts at {y_min}, not at zero, so small differences "
                     "look bigger than they are.")
    place_chart(ws, chart, "G3", 26, 15, sentence)
    set_widths(ws, [5, 26, 16, 12, 16])


# ---- Clean data (kept plain) ----------------------------------------------------
def write_clean_block(ws, clean, first_col, header_row):
    headers = ["Date", "Branch", "Product", "Quantity", "Unit price (₦)", "Sales (₦)"]
    for j, header in enumerate(headers):
        put(ws, header_row, first_col + j, header, bold=True, color="FFFFFF", fill=HEAD)
    ordered = clean.sort_values(["date", "branch", "product"], kind="stable")
    for i, row in enumerate(ordered.itertuples(index=False), start=header_row + 1):
        put(ws, i, first_col, row.date, fmt="yyyy-mm-dd")
        put(ws, i, first_col + 1, row.branch)
        put(ws, i, first_col + 2, row.product)
        put(ws, i, first_col + 3, row.quantity, fmt=UNITS)
        put(ws, i, first_col + 4, row.unit_price, fmt=MONEY)
        put(ws, i, first_col + 5, row.quantity * row.unit_price, fmt=MONEY)


def sheet_clean(ws, d, d_earlier=None):
    ws.title = "Clean data"
    if d_earlier is None:
        write_clean_block(ws, d["clean"], 1, 1)
        ws.freeze_panes = "A2"
        set_widths(ws, [14, 12, 26, 10, 16, 14])
        return
    # Two blocks side by side: current month in A-F, a gap column G, earlier month in H-M.
    title_bar(ws, 1, f"Clean data - {d['month_title']}", d["bar"], 6, first_col=1)
    title_bar(ws, 1, f"Clean data - {d_earlier['month_title']}", d_earlier["bar"], 6, first_col=8)
    write_clean_block(ws, d["clean"], 1, 2)
    write_clean_block(ws, d_earlier["clean"], 8, 2)
    ws.freeze_panes = "A3"
    set_widths(ws, [14, 12, 26, 10, 16, 14, 4, 14, 12, 26, 10, 16, 14])


def build_excel(d, d_earlier=None):
    wb = Workbook()
    sheet_overview(wb.active, d, d_earlier)
    if d.get("compare"):
        sheet_compare(wb.create_sheet("Compare"), d)
    for name, builder in (("Branch totals", sheet_branches), ("Products", sheet_products),
                          ("Branch x Product", sheet_branch_product), ("Weekly", sheet_weekly)):
        ws = wb.create_sheet(name)
        end = builder(ws, d, top=1)
        if d_earlier:  # the earlier month goes in a second section below the first
            builder(ws, d_earlier, top=end + 3)
    sheet_daily(wb.create_sheet("Daily sales"), d)
    sheet_price_units(wb.create_sheet("Price vs Units"), d)
    sheet_clean(wb.create_sheet("Clean data"), d, d_earlier)
    wb.save(output_path(EXCEL_FILE))


# --------------------------------------------------------------------------
# Plain-English summary (every number below is calculated, none typed by hand)
# --------------------------------------------------------------------------
def describe_change(item, earlier_label):
    """One sentence about one branch or product that rose or fell."""
    name = item["name"] + (" (branch)" if item["kind"] == "branch" else "")
    if item["earlier"] == 0:
        return f"{name}: new this month, {naira(item['current'])} (no sales in {earlier_label})."
    direction = "up" if item["change"] > 0 else "down"
    return (f"{name}: {direction} {naira(abs(item['change']))} ({item['pct'] * 100:+.0f}%), "
            f"from {naira(item['earlier'])} to {naira(item['current'])}.")


def comparison_section(c):
    """The 'Compared with <month>' part of summary.md."""
    total = c["total"]
    if total["change"] > 0:
        move = f"up {naira(total['change'])}"
    elif total["change"] < 0:
        move = f"down {naira(abs(total['change']))}"
    else:
        move = "unchanged"
    if total["pct"] is not None and total["change"] != 0:
        move += f" ({signed_pct(total['pct'])})"
    items = c["branches"] + c["products"]
    rises = sorted([i for i in items if i["change"] > 0],
                   key=lambda i: (-i["change"], i["name"]))[:3]
    falls = sorted([i for i in items if i["change"] < 0],
                   key=lambda i: (i["change"], i["name"]))[:3]
    lines = [
        f"## Compared with {c['earlier_label']}",
        "",
        f"Total sales were **{naira(total['current'])}** in {c['cur_label']}, {move} from "
        f"{naira(total['earlier'])} in {c['earlier_label']}.",
        "",
        "**Biggest rises**",
        *([f"{n}. {describe_change(i, c['earlier_label'])}" for n, i in enumerate(rises, start=1)]
          or ["No branch or product rose."]),
        "",
        "**Biggest falls**",
        *([f"{n}. {describe_change(i, c['earlier_label'])}" for n, i in enumerate(falls, start=1)]
          or ["No branch or product fell."]),
        "",
    ]
    return lines


def build_summary(s):
    total = s["grand_sales"]
    top_b, low_b = s["top_branch"], s["low_branch"]
    top_sales = s["branch_sales"][top_b]
    low_sales = s["branch_sales"][low_b]

    part = " (a part week)" if s["best_week_note"] else ""
    went_well = [
        f"**{top_b}** was the strongest branch, bringing in {naira(top_sales)} "
        f"({pct(top_sales, total):.0f}% of all sales).",
        f"**{s['best_product']}** was the best seller: {naira(s['best_product_sales'])} "
        f"from {s['best_product_units']:,} units.",
        f"The strongest week started on Monday {fmt_date(s['best_week_start'])} with "
        f"{naira(s['best_week_sales'])} in sales{part}.",
    ]

    if s["worst_product_sales"] == 0:
        worst_line = f"**{s['worst_product']}** had no sales at all this month."
    else:
        worst_line = (f"**{s['worst_product']}** sold the least: {naira(s['worst_product_sales'])} "
                      f"from {s['worst_product_units']:,} units "
                      f"({pct(s['worst_product_sales'], total):.1f}% of all sales).")
    if s["unsold_count"] > 1:
        worst_line += f" In total {s['unsold_count']} products had no sales."

    if s["n_problem"] == 0:
        quality_line = "Every row in the file could be used."
    else:
        quality_line = (f"{s['n_problem']} of {s['rows_read']} rows "
                        f"({pct(s['n_problem'], s['rows_read']):.1f}%) could not be used and are left "
                        "out of these totals. The reasons are in `output/problems.csv`.")
    if s["price_filled"]:
        quality_line += (f" {s['price_filled']} more rows had no price, so the price list "
                         "was used.")
    watch = [
        f"**{low_b}** had the lowest sales at {naira(low_sales)} "
        f"({pct(low_sales, total):.0f}% of the total), {naira(top_sales - low_sales)} behind {top_b}.",
        worst_line,
        quality_line,
    ]

    if pct(s["spelling_fixed_rows"], s["clean_before_dedupe"]) >= 10:
        suggestion = (
            "Give each branch a fixed drop-down list of branch names and product names, and one "
            "date format, when they record sales. "
            f"{s['spelling_fixed_rows']} of {s['clean_before_dedupe']} usable rows "
            f"({pct(s['spelling_fixed_rows'], s['clean_before_dedupe']):.0f}%) needed a spelling fix "
            f"and {s['date_fixed']} had dates in a different format, which costs checking time "
            "and risks lost sales.")
    else:
        suggestion = (
            f"Review {s['worst_product']}: it brought in only "
            f"{pct(s['worst_product_sales'], total):.1f}% of sales. Consider promoting it, "
            "changing the price, or baking less of it.")

    lines = [
        f"# Mama Tee's Bakery - Sales summary for {s['month_title']}",
        "",
        f"Total sales: **{naira(total)}** from {s['grand_units']:,} items sold across "
        f"{len(BRANCHES)} branches ({', '.join(BRANCHES)} - Abuja).",
        "",
        "## Three things that went well",
        *[f"{i}. {text}" for i, text in enumerate(went_well, start=1)],
        "",
        "## Three things to watch",
        *[f"{i}. {text}" for i, text in enumerate(watch, start=1)],
        "",
        "## One suggestion",
        suggestion,
        "",
    ]
    if s.get("compare"):
        lines += comparison_section(s["compare"])
    lines += [
        "---",
        f"_Numbers behind this summary: {s['rows_read']} rows read; "
        f"{s['dup1']} exact duplicates removed before cleaning and {s['dup2']} after; "
        f"{s['fixed_rows']} rows needed fixing (branch names {s['branch_fixed']}, product names "
        f"{s['product_fixed']}, date format {s['date_fixed']}, blank price {s['price_filled']}); "
        f"{s['n_problem']} rows could not be fixed._",
        "",
    ]
    return "\n".join(lines)


# --------------------------------------------------------------------------
# One month: read, clean, count, tables (nothing is printed or written here)
# --------------------------------------------------------------------------
def analyse_month(path):
    raw = read_input(path)
    rows_read = len(raw)

    # 1. Remove exact duplicates on the raw rows
    deduped = raw.drop_duplicates(subset=INPUT_COLUMNS, keep="first")
    dup1 = rows_read - len(deduped)

    # 2. Clean; rows that cannot be fixed become problems
    clean_pre, problems = clean_rows(deduped)

    # 3. Remove duplicates again now that spellings and dates match
    clean = clean_pre.drop_duplicates(subset=CLEAN_KEY, keep="first").reset_index(drop=True)
    dup2 = len(clean_pre) - len(clean)
    if clean.empty:
        raise SystemExit(f"No usable rows were found in {Path(path).name}, "
                         "so no report was made.")
    clean["sales"] = clean["quantity"] * clean["unit_price"]

    # Fix counts (on rows that passed cleaning, before the second duplicate removal)
    any_fix = clean_pre[["branch_fixed", "product_fixed", "date_fixed", "price_filled"]].any(axis=1)
    spelling_fix = clean_pre[["branch_fixed", "product_fixed"]].any(axis=1)

    # Totals per branch and product, and the independent grand total
    branch_tbl = (clean.groupby("branch").agg(units=("quantity", "sum"), sales=("sales", "sum"))
                  .reindex(BRANCHES, fill_value=0))
    product_tbl = (clean.groupby("product").agg(units=("quantity", "sum"), sales=("sales", "sum"))
                   .reindex(list(PRODUCTS), fill_value=0).rename_axis("product").reset_index())
    grand_sales = float((clean["quantity"] * clean["unit_price"]).sum())

    # Month label: the month that most rows fall in
    year, month = Counter((d.year, d.month) for d in clean["date"]).most_common(1)[0][0]

    return {
        "path": path, "rows_read": rows_read, "dup1": dup1, "dup2": dup2,
        "n_problem": len(problems), "problems": problems,
        "clean_pre": clean_pre, "clean": clean,
        "branch_fixed": int(clean_pre["branch_fixed"].sum()),
        "product_fixed": int(clean_pre["product_fixed"].sum()),
        "date_fixed": int(clean_pre["date_fixed"].sum()),
        "price_filled": int(clean_pre["price_filled"].sum()),
        "fixed_rows": int(any_fix.sum()), "spelling_fixed_rows": int(spelling_fix.sum()),
        "branch_tbl": branch_tbl, "product_tbl": product_tbl,
        "grand_sales": grand_sales, "grand_units": int(clean["quantity"].sum()),
        "branch_sum": float(branch_tbl["sales"].sum()),
        "year": year, "month": month, "month_title": f"{MONTH_NAMES[month - 1]} {year}",
        "month_start": date(year, month, 1),
        "month_end": date(year, month, calendar.monthrange(year, month)[1]),
    }


def report_month(m, role=None, breakdowns=None):
    """Print one month's counts and checks. A failed check stops the run."""
    if role:
        print(f"=== {role}: {m['month_title']} ({Path(m['path']).name}) ===")
    print(f"Reading {m['path']}")
    print()
    print(f"Rows read:                         {m['rows_read']}")
    print(f"Duplicates removed before cleaning: {m['dup1']}")
    print(f"Rows sent to problems.csv:         {m['n_problem']}")
    print(f"Duplicates removed after cleaning:  {m['dup2']}")
    print(f"Total duplicates removed:          {m['dup1'] + m['dup2']}")
    print(f"Rows in the report:                {len(m['clean'])}")
    print(f"Fixes: branch names {m['branch_fixed']}, product names {m['product_fixed']}, "
          f"date formats {m['date_fixed']}, blank prices {m['price_filled']}")

    # Check 1: every row read is accounted for
    dups = m["dup1"] + m["dup2"]
    counted = len(m["clean"]) + dups + m["n_problem"]
    text = (f"rows read {m['rows_read']} = clean {len(m['clean'])} + duplicates {dups} "
            f"+ problems {m['n_problem']}")
    if counted == m["rows_read"]:
        print(f"Check: {text} -> OK")
    else:
        print(f"Check: {text} = {counted} -> MISMATCH")
        raise SystemExit("The row counts do not add up, so no report files were written.")

    # Check 2: branch totals add up to the independent grand total
    check_text = " + ".join(BRANCHES)
    if abs(m["branch_sum"] - m["grand_sales"]) < 0.005:
        print(f"Check: branch totals ({check_text}) = {m['branch_sum']:,.0f}, "
              f"grand total = {m['grand_sales']:,.0f} -> OK")
    else:
        print(f"Check: branch totals ({check_text}) = {m['branch_sum']:,.2f}, "
              f"grand total = {m['grand_sales']:,.2f} -> MISMATCH (difference "
              f"{m['branch_sum'] - m['grand_sales']:,.2f})")
        raise SystemExit("The totals do not agree, so no report files were written.")

    # Check 3 (current month only): every breakdown adds up to the grand total
    if breakdowns is not None:
        off = {k: v for k, v in breakdowns.items() if abs(v - m["grand_sales"]) >= 0.005}
        if off:
            raise SystemExit(f"These sheets do not add up to the grand total: {off}. "
                             "No report files were written.")
        print("Check: branch x product, weekly, daily and product tables all add up to the "
              "grand total -> OK")


# --------------------------------------------------------------------------
# A month's extra tables (weeks, days, branch x product, highlights)
# --------------------------------------------------------------------------
def analyse_current(m):
    clean, branch_tbl, product_tbl = m["clean"], m["branch_tbl"], m["product_tbl"]
    month_start, month_end = m["month_start"], m["month_end"]

    best = product_tbl.sort_values(["sales", "product"], ascending=[False, True]).head(5)
    worst = product_tbl.sort_values(["sales", "product"], ascending=[True, True]).head(5)

    # Branch x product grid, and each product's share of its branch's sales
    pivot = (clean.pivot_table(index="product", columns="branch", values="sales",
                               aggfunc="sum", fill_value=0)
             .reindex(index=list(PRODUCTS), columns=BRANCHES, fill_value=0))
    branch_totals = pivot.sum(axis=0).replace(0, float("nan"))
    shares = (pivot / branch_totals).fillna(0)

    # Weeks (Monday to Sunday) and days
    week_units, week_sales = Counter(), Counter()
    day_units, day_sales = Counter(), Counter()
    for d, qty, amount in zip(clean["date"], clean["quantity"], clean["sales"]):
        start = d - timedelta(days=d.weekday())
        week_units[start] += int(qty)
        week_sales[start] += float(amount)
        day_units[d] += int(qty)
        day_sales[d] += float(amount)

    weekly = []
    start = min(week_sales)
    while start <= max(week_sales):
        end = start + timedelta(days=6)
        weekly.append({
            "label": f"{short_date(start)} - {short_date(end)}",
            "start": start, "end": end, "units": week_units.get(start, 0),
            "sales": week_sales.get(start, 0.0),
            "note": "Part week" if start < month_start or end > month_end else "",
        })
        start += timedelta(days=7)
    best_week = max(weekly, key=lambda w: w["sales"])

    daily = []
    day = min(month_start, min(day_sales))  # every calendar day, so quiet days show as 0
    while day <= max(month_end, max(day_sales)):
        daily.append({
            "label": short_date(day), "date": day, "weekday": WEEKDAYS[day.weekday()],
            "units": day_units.get(day, 0), "sales": day_sales.get(day, 0.0),
        })
        day += timedelta(days=1)
    days_with_sales = [x for x in daily if x["sales"] > 0]  # best/worst day ignore empty days
    best_day = max(days_with_sales, key=lambda x: x["sales"])
    worst_day = min(days_with_sales, key=lambda x: x["sales"])

    breakdowns = {
        "branch x product": float(pivot.values.sum()),
        "weekly": sum(w["sales"] for w in weekly),
        "daily": sum(x["sales"] for x in daily),
        "products": float(product_tbl["sales"].sum()),
    }

    def day_name(x):
        return f"{x['weekday']} {fmt_date(x['date'])}"

    highlights = [
        ["Best branch", branch_tbl["sales"].idxmax(), float(branch_tbl["sales"].max())],
        ["Worst branch", branch_tbl["sales"].idxmin(), float(branch_tbl["sales"].min())],
        ["Best product", best.iloc[0]["product"], float(best.iloc[0]["sales"])],
        ["Worst product", worst.iloc[0]["product"], float(worst.iloc[0]["sales"])],
        ["Best day", day_name(best_day), best_day["sales"]],
        ["Worst day (days with sales)", day_name(worst_day), worst_day["sales"]],
    ]
    return {"best": best, "worst": worst, "pivot": pivot, "shares": shares, "weekly": weekly,
            "daily": daily, "best_week": best_week, "breakdowns": breakdowns,
            "highlights": highlights}


# --------------------------------------------------------------------------
# Comparison with an earlier month
# --------------------------------------------------------------------------
def make_item(name, kind, earlier, current):
    change = current - earlier
    return {"name": name, "kind": kind, "earlier": earlier, "current": current,
            "change": change, "pct": change / earlier if earlier else None}


def build_comparison(cur, earlier):
    """Per-branch and per-product sales for both months, with the change (current - earlier)."""
    branches = [make_item(b, "branch", float(earlier["branch_tbl"].loc[b, "sales"]),
                          float(cur["branch_tbl"].loc[b, "sales"])) for b in BRANCHES]
    earlier_prod = earlier["product_tbl"].set_index("product")["sales"]
    cur_prod = cur["product_tbl"].set_index("product")["sales"]
    products = [make_item(p, "product", float(earlier_prod[p]), float(cur_prod[p]))
                for p in PRODUCTS]
    same_year = cur["year"] == earlier["year"]
    return {
        "cur_title": cur["month_title"], "earlier_title": earlier["month_title"],
        "cur_label": MONTH_NAMES[cur["month"] - 1] if same_year else cur["month_title"],
        "earlier_label": MONTH_NAMES[earlier["month"] - 1] if same_year else earlier["month_title"],
        "branches": branches, "products": products,
        "total": make_item("All sales", "total", earlier["grand_sales"], cur["grand_sales"]),
    }


def check_comparison(comp, cur, earlier):
    """The Compare sheet's totals must equal each month's own grand total."""
    errors, parts = [], []
    for m, side in ((cur, "current"), (earlier, "earlier")):
        parts.append(f"{m['month_title']} {m['grand_sales']:,.0f}")
        for label, items in (("branch rows", comp["branches"]), ("product rows", comp["products"])):
            value = sum(i[side] for i in items)
            if abs(value - m["grand_sales"]) >= 0.005:
                errors.append(f"{m['month_title']}: {label} add up to {value:,.2f}, "
                              f"grand total is {m['grand_sales']:,.2f}")
    if errors:
        print("Check: Compare sheet totals equal each month's own grand total -> MISMATCH")
        for line in errors:
            print(f"  {line}")
        raise SystemExit("The comparison totals do not agree, so no report files were written.")
    print("Check: Compare sheet totals equal each month's own grand total "
          f"({'; '.join(parts)}) -> OK")


def check_sections(comp, cur, x_cur, earlier, x_earlier, d_cur, d_earlier):
    """Every section's totals must equal its own month's grand total, and the Overview
    comparison block must match the Compare sheet."""
    errors = []
    for m, x, d in ((cur, x_cur, d_cur), (earlier, x_earlier, d_earlier)):
        grand, label = m["grand_sales"], m["month_title"]
        sections = {
            "Branch totals": float(m["branch_tbl"]["sales"].sum()),
            "Branch x Product": float(x["pivot"].values.sum()),
            "Weekly": sum(w["sales"] for w in x["weekly"]),
            "Daily": sum(day["sales"] for day in x["daily"]),
            "Products (all 14)": float(m["product_tbl"]["sales"].sum()),
            "Clean data block": float((m["clean"]["quantity"] * m["clean"]["unit_price"]).sum()),
        }
        for name, value in sections.items():
            if abs(value - grand) >= 0.005:
                errors.append(f"{label}: {name} adds up to {value:,.2f}, grand total is {grand:,.2f}")
        if int(m["branch_tbl"]["units"].sum()) != m["grand_units"]:
            errors.append(f"{label}: branch units do not add up to the items sold")
        block = d["block"]
        if block["last"] - block["first"] + 1 != len(m["clean"]):
            errors.append(f"{label}: Clean data block does not hold all {len(m['clean'])} rows")

    # Overview comparison block against the Compare sheet
    total = comp["total"]
    pairs = (("earlier", d_earlier), ("current", d_cur))
    for side, d in pairs:
        hl = {row[0]: row[2] for row in d["highlights"]}
        want = {
            "Total sales": (d["grand_sales"], total[side]),
            "Best branch": (hl["Best branch"], max(i[side] for i in comp["branches"])),
            "Worst branch": (hl["Worst branch"], min(i[side] for i in comp["branches"])),
            "Best product": (hl["Best product"], max(i[side] for i in comp["products"])),
            "Worst product": (hl["Worst product"], min(i[side] for i in comp["products"])),
        }
        for name, (overview_value, compare_value) in want.items():
            if abs(overview_value - compare_value) >= 0.005:
                errors.append(f"Overview {name} ({side} month) is {overview_value:,.2f} but the "
                              f"Compare sheet gives {compare_value:,.2f}")
    if abs((d_cur["grand_sales"] - d_earlier["grand_sales"]) - total["change"]) >= 0.005:
        errors.append("Overview total sales change does not match the Compare sheet")

    if errors:
        print("Check: every section's totals equal its month's grand total, and the Overview "
              "comparison matches the Compare sheet -> MISMATCH")
        for line in errors:
            print(f"  {line}")
        raise SystemExit("A section does not add up, so no report files were written.")
    print("Check: every section's totals equal its month's grand total, and the Overview "
          "comparison matches the Compare sheet -> OK")


def make_pack(m, x, compare=None, bar=None, chart_colour=BAR, worst_colour=BAR_LIGHT,
              block=None, heat_range=None):
    """Everything the Excel sheets need for one month."""
    return {
        "clean": m["clean"], "month_title": m["month_title"], "highlights": x["highlights"],
        "best": x["best"], "worst": x["worst"], "product_tbl": m["product_tbl"],
        "pivot": x["pivot"], "shares": x["shares"], "weekly": x["weekly"], "daily": x["daily"],
        "grand_sales": m["grand_sales"], "grand_units": m["grand_units"],
        "compare": compare, "bar": bar, "chart_colour": chart_colour,
        "worst_colour": worst_colour, "block": block, "heat_range": heat_range,
    }


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    current_path = find_input(args.current)
    compare_path = None
    if args.compare:
        compare_path = Path(args.compare)
        if not compare_path.is_file():
            raise SystemExit(f"Comparison file not found: {compare_path}")
        if compare_path.resolve() == Path(current_path).resolve():
            raise SystemExit("The --compare file is the same file as the current month. Name the "
                             "current month's file first, then --compare with the earlier month.")

    cur = analyse_month(current_path)
    earlier = analyse_month(compare_path) if compare_path else None
    if earlier and (earlier["year"], earlier["month"]) >= (cur["year"], cur["month"]):
        raise SystemExit(f"The --compare file is {earlier['month_title']}, which is not earlier "
                         f"than the current file ({cur['month_title']}). Nothing was written.")
    extra = analyse_current(cur)
    extra_earlier = analyse_current(earlier) if earlier else None

    # Print each month's counts and checks (a failed check stops the run before any writing)
    report_month(cur, "Current month" if earlier else None, extra["breakdowns"])
    comparison = None
    d_cur = d_earlier = None
    if earlier:
        print()
        report_month(earlier, "Earlier month")
        comparison = build_comparison(cur, earlier)
        print()
        check_comparison(comparison, cur, earlier)

        # Where each month's rows sit on the Clean data sheet (title bar row 1, header row 2)
        block_cur = {"first": 3, "last": 2 + len(cur["clean"]), "cols": SINGLE_COLS}
        block_earlier = {"first": 3, "last": 2 + len(earlier["clean"]), "cols": EARLIER_COLS}
        # One shared heat-map range, so the same colour means the same amount in both months
        grids = (extra["pivot"].values, extra_earlier["pivot"].values)
        heat = (float(min(g.min() for g in grids)), float(max(g.max() for g in grids)))
        d_cur = make_pack(cur, extra, compare=comparison, bar=HEAD, block=block_cur,
                          heat_range=heat)
        d_earlier = make_pack(earlier, extra_earlier, bar=EARLIER_BAR, chart_colour=EARLIER_BAR,
                              worst_colour=EARLIER_BAR_LIGHT, block=block_earlier,
                              heat_range=heat)
        check_sections(comparison, cur, extra, earlier, extra_earlier, d_cur, d_earlier)
    else:
        d_cur = make_pack(cur, extra)

    best, worst, best_week = extra["best"], extra["worst"], extra["best_week"]
    branch_tbl, product_tbl = cur["branch_tbl"], cur["product_tbl"]
    stats = {
        "month_title": cur["month_title"], "grand_sales": cur["grand_sales"],
        "grand_units": cur["grand_units"],
        "branch_sales": {b: float(branch_tbl.loc[b, "sales"]) for b in BRANCHES},
        "top_branch": branch_tbl["sales"].idxmax(), "low_branch": branch_tbl["sales"].idxmin(),
        "best_product": best.iloc[0]["product"], "best_product_sales": float(best.iloc[0]["sales"]),
        "best_product_units": int(best.iloc[0]["units"]),
        "worst_product": worst.iloc[0]["product"], "worst_product_sales": float(worst.iloc[0]["sales"]),
        "worst_product_units": int(worst.iloc[0]["units"]),
        "unsold_count": int((product_tbl["sales"] == 0).sum()),
        "best_week_start": best_week["start"], "best_week_sales": best_week["sales"],
        "best_week_note": best_week["note"],
        "rows_read": cur["rows_read"], "dup1": cur["dup1"], "dup2": cur["dup2"],
        "n_problem": cur["n_problem"],
        "branch_fixed": cur["branch_fixed"], "product_fixed": cur["product_fixed"],
        "date_fixed": cur["date_fixed"], "price_filled": cur["price_filled"],
        "fixed_rows": cur["fixed_rows"], "spelling_fixed_rows": cur["spelling_fixed_rows"],
        "clean_before_dedupe": len(cur["clean_pre"]),
        "compare": comparison,
    }

    # problems.csv: the current month as before. In comparison mode both months' rows are
    # listed, with a "file" column saying which file each row came from.
    problems = cur["problems"]
    if earlier:
        problems = pd.concat([
            cur["problems"].assign(file=Path(cur["path"]).name),
            earlier["problems"].assign(file=Path(earlier["path"]).name),
        ], ignore_index=True)[["file"] + PROBLEM_COLUMNS]

    # Write the three output files (and nothing else)
    OUT_DIR.mkdir(exist_ok=True)
    try:
        problems.to_csv(output_path(PROBLEMS_FILE), index=False, encoding="utf-8-sig")
        build_excel(d_cur, d_earlier)
        output_path(SUMMARY_FILE).write_text(build_summary(stats), encoding="utf-8")
    except PermissionError as err:
        raise SystemExit(f"Could not write {Path(err.filename).name}. "
                         "If it is open in Excel, close it and run again.")

    print()
    print("Written:")
    for path in (PROBLEMS_FILE, EXCEL_FILE, SUMMARY_FILE):
        print(f"  {path.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()
