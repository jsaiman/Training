"""Column mappings, thresholds and rules for the Halvern workbook.

This file is the single source of truth. process.py imports it; change rules here and nowhere else.
"""
import datetime as dt

# ---- Sheets
SHEET_PORTFOLIO = "Portfolio"
SHEET_MARKS = "Quarterly Marks"
SHEET_FUNDS = "Fund Summary"

# ---- Column mappings: internal key -> exact header text in row 1
PORTFOLIO_COLUMNS = {
    "company":   "Company",                    # P1 text, unique -> company key
    "fund":      "Fund",                       # P2 must exist on Fund Summary
    "sector":    "Sector",                     # P3 text -> sector mix
    "geo":       "Geography",                  # P3 text -> geography mix
    "entry":     "Entry Date",                 # P4 true date value
    "invested":  "Invested Capital (£m)",      # P5 number > 0
    "valuation": "Current Valuation (£m)",     # P6 number >= 0, not blank
    "realised":  "Realised Proceeds (£m)",     # P7 number >= 0, not blank
    "moic":      "MOIC",                       # P8 must be formula (valuation+realised)/invested
    "status":    "Status",                     # P9 Active or Exited
}
FUND_COLUMNS = {
    "fund":           "Fund",
    "vintage":        "Vintage Year",          # integer
    "total_invested": "Total Invested (£m)",   # SUMIF formula, must span all Portfolio rows
    "total_value":    "Total Value (£m)",      # SUMIF formula, must span all Portfolio rows
    "gross_moic":     "Gross MOIC",
    "net_irr":        "Net IRR %",             # numeric, hard-coded, cannot be verified
    "dpi":            "DPI",
    "tvpi":           "TVPI",                  # must not be the same formula as Gross MOIC
}
MARKS_COMPANY_HEADER = "Company"               # every other header must be a quarter label
QUARTER_HEADER_REGEX = r"^Q([1-4]) (\d{4})$"   # e.g. "Q3 2025"

# ---- Values and thresholds
STATUS_ACTIVE, STATUS_EXITED = "Active", "Exited"
STATUSES = (STATUS_ACTIVE, STATUS_EXITED)
TOLERANCE = 0.05          # GBP m, used for Q6, Q7 and F3 comparisons
QOQ_LIMIT = 0.30          # Q5: absolute quarter-on-quarter move that gets flagged
TODAY = None              # None means the date the code is run; set a date to reproduce a past run
NAME_SUFFIXES = ["ltd", "limited", "plc", "inc", "group", "holdings", "llp"]   # P1 near-duplicate test
Q6_TREATMENT = "Portfolio tab confirmed as correct by Jeremie; mark difference shown for reference only"

# ---- Rules: (id, where, rule text, severity if it fails)
# Blocking = row excluded from every headline figure. Warning = shown and flagged. Info = shown for information.
RULES = [
    ("P1", "Portfolio Company", "Non-empty text, unique. Exact duplicates and near-duplicates (same name once suffixes in NAME_SUFFIXES are removed, same fund and sector) are blocking; all rows involved are quarantined.", "Blocking"),
    ("P2", "Portfolio Fund", "Must match a fund on Fund Summary.", "Blocking"),
    ("P3", "Portfolio Sector, Geography", "Non-empty text.", "Blocking"),
    ("P4", "Portfolio Entry Date", "True date value, not after today, not before fund vintage year. Text dates are not parsed. Drops the row from date views only.", "Warning"),
    ("P5", "Portfolio Invested Capital", "Number above 0.", "Blocking"),
    ("P6", "Portfolio Current Valuation", "Number of 0 or more. Blank is not treated as 0.", "Blocking"),
    ("P7", "Portfolio Realised Proceeds", "Number of 0 or more. Blank is not treated as 0.", "Blocking"),
    ("P8", "Portfolio MOIC", "Must be the formula (valuation+realised)/invested on the same row. A typed value, or a different formula, is blocking.", "Blocking"),
    ("P9", "Portfolio Status", "Exactly Active or Exited. Exited needs Valuation = 0 and Realised > 0. Active needs Valuation > 0.", "Blocking"),
    ("Q1", "Marks Company", "Names must match the Portfolio tab one-to-one.", "Warning"),
    ("Q2", "Marks quarter columns", "Headers match QUARTER_HEADER_REGEX and are contiguous. Cells are numbers >= 0 or blank. Text such as '12.4 (est)' is flagged, not parsed.", "Warning"),
    ("Q3", "Marks gaps", "Blank allowed only before the first mark. A later blank is a gap and is not interpolated.", "Warning"),
    ("Q4", "Marks first mark", "First mark quarter should equal the entry quarter (or the first quarter on the tab if entry was earlier).", "Warning"),
    ("Q5", "Marks QoQ", "Move of more than QOQ_LIMIT between consecutive valid marks is flagged.", "Warning"),
    ("Q6", "Marks vs Portfolio", "Latest mark must equal Current Valuation (Active) or Realised Proceeds (Exited) within TOLERANCE. Portfolio tab is the source of truth, so a break does not change any headline figure.", "Warning"),
    ("Q7", "Marks vs Portfolio", "First mark compared with Invested Capital (information only).", "Info"),
    ("F1", "Fund Summary", "Fund names unique, vintage integer, Net IRR numeric. Net IRR is shown as reported and cannot be verified without cash flows.", "Info"),
    ("F2", "Fund Summary totals", "SUMIF ranges must cover every Portfolio row.", "Warning"),
    ("F3", "Fund Summary totals", "Recompute invested, value, MOIC and DPI from validated rows and bridge to the workbook formula results (blanks counted as 0, as Excel does).", "Warning"),
    ("F4", "Fund Summary TVPI", "TVPI formula must differ from Gross MOIC; flagged if identical.", "Warning"),
    ("F5", "Whole workbook", "Formulas with no stored results are flagged and recomputed from their stated definitions.", "Info"),
    ("X1", "Whole workbook", "Coverage: sheets, columns or funds the rules do not cover are flagged and not processed.", "Warning"),
]
