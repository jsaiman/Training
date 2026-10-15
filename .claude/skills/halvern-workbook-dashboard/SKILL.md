---
description: Use only when Jeremie supplies a new or updated Halvern fund performance workbook (tabs Portfolio, Quarterly Marks, Fund Summary) to validate or turn into a dashboard; do not use for other workbooks, general questions about fund figures, or edits to the spreadsheet itself. It runs fixed validation code, flags every exception and never guesses a correction.
---

# Halvern workbook dashboard

The code decides what is valid. You do not.

## Operating rules

1. **Run the code. Do not interpret the data yourself.** Every figure, flag and exclusion comes from `scripts/process.py` using `scripts/config.py`. Do not calculate, correct, fill, round or tidy anything by hand.
2. **Never guess.** No corrected values, no parsed text dates, no blanks treated as 0, no interpolation, no choosing between near-duplicates or between disagreeing figures. Report it and ask.
3. **Flag anything the code does not cover.** If you notice something it does not report (new tab or column, changed layout, units, a figure that passes every rule but looks wrong), tell Jeremie in chat and say it is outside the rules. Do not fix it or ignore it. Read `references/rules.md` for what is and is not covered.
4. **Do not edit `scripts/` or `assets/` to make a file pass.** If a new version breaks the rules or layout, say so and ask whether the rules should change.
5. **If the script prints `FATAL` and exits with code 2, stop.** Report exactly what is missing and ask. Do not map headers by guesswork.

## Steps

1. Run, using a scratchpad path for the output (not this folder):

   ```
   python3 scripts/process.py scripts/config.py <workbook.xlsx> assets/template.html <output>/dashboard.html
   ```

2. Report the printed output in chat. The `ISSUES` list is the complete set of findings. Read `references/reporting.md` for the reply format and the dashboard check.
3. Save `dashboard.html` where Jeremie can open it.

## Load only when needed

- `references/rules.md`: what each rule means, what the code covers and what it does not. Read when explaining a finding or deciding whether something is out of scope.
- `references/reporting.md`: reply order, wording and dashboard render check. Read before replying.
- `references/maintenance.md`: how to change a rule or compare versions. Read only when Jeremie asks to change the rules or compare workbooks.
- `scripts/config.py`: the column mappings, thresholds and full rule table as code. Read only when you need the exact mapping or a threshold.
