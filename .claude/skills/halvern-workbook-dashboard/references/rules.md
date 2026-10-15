# Rules and coverage

`scripts/config.py` is the authority; this file explains it. Severity decides treatment:

- **Blocking:** the row is excluded from every headline figure and listed. Rules P1, P2, P3, P5, P6, P7, P8, P9.
- **Warning:** shown and flagged, the row stays in. Rules P4, Q1-Q6, F2-F4, X1.
- **Info:** shown for information. Rules Q7, F1, F5.

## Sheets and what each feeds

- **Portfolio:** one row per company. Source of truth for all money figures, headline totals, mixes and MOIC.
- **Quarterly Marks:** valuation history by quarter. Used for the trend chart and the reconciliation to Portfolio (Q6). A trend point is plotted only if every company expected to have a mark has a clean one.
- **Fund Summary:** fund names, vintage and hard-coded Net IRR. Its formulas are recomputed and bridged to the validated totals (F3).

## Things the rules deliberately do not do

- Parse a text date, treat a blank as 0, fix a sign, interpolate a gap, or parse text such as "12.4 (est)".
- Decide which of two near-duplicate rows is right, or which side of a MOIC mismatch is wrong.
- Verify Net IRR (no cash flows) or treat TVPI as net of fees (the workbook formula makes it gross).

## Not covered: raise these in chat

- Units, currency or reporting-date changes, and restated historic figures.
- Whether a flagged near-duplicate is one company or two.
- Anything reported under X1 (extra tabs or columns, hidden sheets, funds with no rows).
- Anything that passes every rule but looks implausible.
