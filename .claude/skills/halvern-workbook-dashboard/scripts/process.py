#!/usr/bin/env python3
"""Validate a Halvern workbook against config.py and build the dashboard.

Usage: python3 scripts/process.py scripts/config.py <workbook.xlsx> <template.html> <output.html>

All column mappings, thresholds and rules come from config.py (copied from SKILL.md).
Nothing is corrected, filled or interpolated. Anything outside the rules is recorded as an issue.
Exit code 2 means a required sheet or header is missing and nothing was processed.
"""
import datetime as dt
import importlib.util
import json
import math
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as L

issues = []


def issue(sev, rule, sheet, cell, entity, found, treatment):
    issues.append(dict(sev=sev, rule=rule, sheet=sheet, cell=cell, entity=entity,
                       found=found, treatment=treatment))


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def load_config(path):
    spec = importlib.util.spec_from_file_location("config", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def header_map(ws, mapping, sheet_label, strict_extra=True):
    """Return {key: 1-based column}. Missing required header is fatal. Unmapped headers are flagged (X1)."""
    hdr = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None:
            hdr[str(v).strip()] = c
    out, missing = {}, []
    for key, name in mapping.items():
        if name in hdr:
            out[key] = hdr[name]
        else:
            missing.append(name)
    if missing:
        print(f"FATAL: sheet '{sheet_label}' is missing required headers: {missing}. Found: {list(hdr)}")
        sys.exit(2)
    return out, hdr


def main(cfg_path, path, tpl, out):
    cfg = load_config(cfg_path)
    TODAY = cfg.TODAY or dt.date.today()
    SUFFIX = re.compile(r"\b(" + "|".join(cfg.NAME_SUFFIXES) + r")\b\.?", re.I)
    QHDR = re.compile(cfg.QUARTER_HEADER_REGEX)
    norm = lambda n: re.sub(r"\s+", " ", SUFFIX.sub("", n or "")).strip().lower()

    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)

    # ---- X1: coverage of sheets
    for need in (cfg.SHEET_PORTFOLIO, cfg.SHEET_MARKS, cfg.SHEET_FUNDS):
        if need not in wb.sheetnames:
            print(f"FATAL: required sheet '{need}' not found. Sheets: {wb.sheetnames}")
            sys.exit(2)
    for s in wb.worksheets:
        if s.title not in (cfg.SHEET_PORTFOLIO, cfg.SHEET_MARKS, cfg.SHEET_FUNDS):
            issue("Warning", "X1", s.title, "-", "-", f"Sheet '{s.title}' is not covered by the rules", "Not processed; needs a decision")
        if s.sheet_state != "visible":
            issue("Warning", "X1", s.title, "-", "-", f"Sheet is {s.sheet_state}", "Flagged")

    pf, qm, fs = wb[cfg.SHEET_PORTFOLIO], wb[cfg.SHEET_MARKS], wb[cfg.SHEET_FUNDS]
    pc, phdr = header_map(pf, cfg.PORTFOLIO_COLUMNS, cfg.SHEET_PORTFOLIO)
    fc, fhdr = header_map(fs, cfg.FUND_COLUMNS, cfg.SHEET_FUNDS)
    mkey, mhdr = header_map(qm, {"company": cfg.MARKS_COMPANY_HEADER}, cfg.SHEET_MARKS)
    for sheet, hdr, mp in ((cfg.SHEET_PORTFOLIO, phdr, cfg.PORTFOLIO_COLUMNS), (cfg.SHEET_FUNDS, fhdr, cfg.FUND_COLUMNS)):
        for h, c in hdr.items():
            if h not in mp.values():
                issue("Warning", "X1", sheet, f"{L(c)}1", "-", f"Column '{h}' is not covered by the rules", "Ignored; needs a decision")
    P = lambda k: L(pc[k])
    F = lambda k: L(fc[k])

    # ---- Fund Summary (F1)
    funds = []
    for r in range(2, fs.max_row + 1):
        name = fs.cell(r, fc["fund"]).value
        if name is None:
            continue
        vint, irr = fs.cell(r, fc["vintage"]).value, fs.cell(r, fc["net_irr"]).value
        if not isinstance(vint, int) or isinstance(vint, bool):
            issue("Warning", "F1", cfg.SHEET_FUNDS, f"{F('vintage')}{r}", name, f"Vintage {vint!r} is not an integer", "Fund vintage not used")
        if not is_num(irr):
            issue("Warning", "F1", cfg.SHEET_FUNDS, f"{F('net_irr')}{r}", name, f"Net IRR {irr!r} is not numeric", "Not shown")
        funds.append(dict(row=r, name=str(name).strip(), vintage=vint if isinstance(vint, int) else None,
                          irr=irr if is_num(irr) else None,
                          f_moic=fs.cell(r, fc["gross_moic"]).value, f_tvpi=fs.cell(r, fc["tvpi"]).value))
    lower = [f["name"].lower() for f in funds]
    for f in funds:
        if lower.count(f["name"].lower()) > 1:
            issue("Warning", "F1", cfg.SHEET_FUNDS, f"{F('fund')}{f['row']}", f["name"], "Duplicate fund name", "Flagged")
        if f["irr"] is not None:
            issue("Info", "F1", cfg.SHEET_FUNDS, f"{F('net_irr')}{f['row']}", f["name"],
                  f"Net IRR {f['irr']}% is hard-coded; no cash flows in the workbook to recompute it", "Shown as reported, labelled unverified")
    fund_map = {f["name"]: f for f in funds}

    last_pf = max((r for r in range(2, pf.max_row + 1) if pf.cell(r, pc["company"]).value is not None), default=1)
    for f in funds:
        for key in ("total_invested", "total_value"):
            m = re.findall(r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", str(fs.cell(f["row"], fc[key]).value))
            if not m or any(int(x[3]) < last_pf or int(x[1]) > 2 for x in m):
                issue("Warning", "F2", cfg.SHEET_FUNDS, f"{F(key)}{f['row']}", f["name"],
                      f"SUMIF range does not span Portfolio rows 2-{last_pf}", "Workbook totals may omit rows")
        if f["f_tvpi"] and f["f_tvpi"] == f["f_moic"]:
            issue("Warning", "F4", cfg.SHEET_FUNDS, f"{F('tvpi')}{f['row']}", f["name"],
                  f"TVPI formula {f['f_tvpi']} is identical to Gross MOIC; TVPI is normally net of fees and carry, so this is a gross figure under a net label",
                  "Shown as reported; not presented as a net measure")
    fsv = wbv[cfg.SHEET_FUNDS]
    if funds and all(fsv.cell(f["row"], fc["total_invested"]).value is None for f in funds):
        issue("Info", "F5", "Whole workbook", "-", "-", "Formula cells have no stored results (the file was never calculated in Excel)",
              "All formula results recomputed here from the stated formulas")

    # ---- Portfolio
    rows = []
    for r in range(2, last_pf + 1):
        g = lambda k: pf.cell(r, pc[k]).value
        rows.append(dict(row=r, company=g("company"), fund=g("fund"), sector=g("sector"), geo=g("geo"), entry=g("entry"),
                         inv=g("invested"), val=g("valuation"), real=g("realised"), moic_cell=g("moic"), status=g("status")))
    for p in rows:
        p["block"], p["entry_q"], n = [], None, p["row"]
        name = p["company"]

        def block(rule, key, found):
            p["block"].append(rule)
            issue("Blocking", rule, cfg.SHEET_PORTFOLIO, f"{P(key)}{n}", name, found, "Row excluded from all headline figures")

        if not isinstance(name, str) or not name.strip():
            name = f"(row {n})"; p["company"] = name
            block("P1", "company", "Company name missing")
        else:
            p["company"] = name = name.strip()
        if p["fund"] not in fund_map:
            block("P2", "fund", f"Fund {p['fund']!r} not on {cfg.SHEET_FUNDS}")
        for key, lab in (("sector", "Sector"), ("geo", "Geography")):
            if not isinstance(p[key], str) or not p[key].strip():
                block("P3", key, f"{lab} missing")
        e, vint = p["entry"], fund_map.get(p["fund"], {}).get("vintage")
        if isinstance(e, dt.datetime):
            d = e.date()
            if d > TODAY:
                issue("Warning", "P4", cfg.SHEET_PORTFOLIO, f"{P('entry')}{n}", name, f"Entry date {d} is in the future", "Excluded from date views")
            elif isinstance(vint, int) and d.year < vint:
                issue("Warning", "P4", cfg.SHEET_PORTFOLIO, f"{P('entry')}{n}", name, f"Entry date {d} is before fund vintage {vint}", "Excluded from date views")
            else:
                p["entry_q"] = (d.year, (d.month - 1) // 3 + 1)
        else:
            issue("Warning", "P4", cfg.SHEET_PORTFOLIO, f"{P('entry')}{n}", name,
                  f"Entry date is not a date value: {e!r}" if e is not None else "Entry date missing",
                  "Not parsed or corrected; excluded from date views only")
        if not is_num(p["inv"]):
            block("P5", "invested", f"Invested Capital is {p['inv']!r}, not a number")
        elif p["inv"] <= 0:
            block("P5", "invested", f"Invested Capital is {p['inv']}; must be above 0")
        for key, rule, lab in (("val", "P6", "Current Valuation"), ("real", "P7", "Realised Proceeds")):
            col = "valuation" if key == "val" else "realised"
            if p[key] is None:
                block(rule, col, f"{lab} is blank (not assumed to be 0)")
            elif not is_num(p[key]) or p[key] < 0:
                block(rule, col, f"{lab} is {p[key]!r}; must be a number of 0 or more")
        st = p["status"]
        if st not in cfg.STATUSES:
            block("P9", "status", f"Status is {st!r}; must be one of {list(cfg.STATUSES)}")
        elif is_num(p["val"]) and is_num(p["real"]):
            if st == cfg.STATUS_EXITED and not (p["val"] == 0 and p["real"] > 0):
                block("P9", "status", f"{st} but Valuation = {p['val']} and Realised = {p['real']}")
            if st == cfg.STATUS_ACTIVE and not p["val"] > 0:
                block("P9", "status", f"{st} but Valuation = {p['val']}")
        mc = p["moic_cell"]
        ok = is_num(p["inv"]) and p["inv"] > 0 and is_num(p["val"]) and is_num(p["real"])
        p["moic"] = (p["val"] + p["real"]) / p["inv"] if ok else None
        expected = f"=({P('valuation')}{n}+{P('realised')}{n})/{P('invested')}{n}"
        if isinstance(mc, str) and mc.replace(" ", "") == expected:
            pass
        elif is_num(mc):
            note = f"; recomputed (valuation + realised) / invested = {p['moic']:.2f}x" if p["moic"] is not None else ""
            block("P8", "moic", f"MOIC is a typed value {mc}x, not the formula {expected}{note}. Either the MOIC or the inputs are wrong; cannot tell which")
        else:
            block("P8", "moic", f"MOIC cell is {mc!r}; expected formula {expected}")

    seen = {}
    for p in rows:
        seen.setdefault(p["company"].lower(), []).append(p)
    for grp in seen.values():
        if len(grp) > 1:
            for p in grp:
                if "P1" not in p["block"]:
                    p["block"].append("P1")
                issue("Blocking", "P1", cfg.SHEET_PORTFOLIO, f"{P('company')}{p['row']}", p["company"],
                      f"Exact duplicate company name (rows {', '.join(str(x['row']) for x in grp)})", "All duplicate rows excluded")
    groups = {}
    for p in rows:
        groups.setdefault((norm(p["company"]), p["fund"], p["sector"]), []).append(p)
    for grp in groups.values():
        if len(grp) > 1 and len({x["company"].lower() for x in grp}) > 1:
            desc = "; ".join(f"row {x['row']} '{x['company']}' invested {x['inv']}, valued {x['val']}, entry {x['entry'].date() if isinstance(x['entry'], dt.datetime) else x['entry']}" for x in grp)
            for x in grp:
                if "P1" not in x["block"]:
                    x["block"].append("P1")
                issue("Blocking", "P1", cfg.SHEET_PORTFOLIO, f"{P('company')}{x['row']}", x["company"],
                      f"Possible duplicate: same company, fund and sector under two names. {desc}", "Both rows excluded; cannot tell which is correct, or whether both are real")

    # ---- Quarterly Marks
    qcols = []
    for c in range(1, qm.max_column + 1):
        h = qm.cell(1, c).value
        if c == mkey["company"] or h is None:
            continue
        m = QHDR.match(str(h).strip())
        if m:
            qcols.append((c, str(h).strip(), (int(m.group(2)), int(m.group(1)))))
        else:
            issue("Warning", "Q2", cfg.SHEET_MARKS, f"{L(c)}1", "-", f"Header {h!r} is not a quarter label", "Column ignored; needs a decision")
    quarters = [q[1] for q in qcols]
    qkeys = [q[2] for q in qcols]
    for a, b in zip(qkeys, qkeys[1:]):
        if (b[0] * 4 + b[1]) - (a[0] * 4 + a[1]) != 1:
            issue("Warning", "Q2", cfg.SHEET_MARKS, "row 1", "-", f"Quarter headers not contiguous at {a} -> {b}", "Flagged")
    first_q = qkeys[0] if qkeys else None
    pmap = {p["company"].lower(): p for p in rows}
    qm_names = set()
    for r in range(2, qm.max_row + 1):
        nm = qm.cell(r, mkey["company"]).value
        if nm is None:
            continue
        nm = str(nm).strip()
        qm_names.add(nm.lower())
        p = pmap.get(nm.lower())
        if p is None:
            issue("Warning", "Q1", cfg.SHEET_MARKS, f"{L(mkey['company'])}{r}", nm, "Company not on Portfolio tab", "Ignored")
            continue
        raw = [qm.cell(r, c).value for c, _, _ in qcols]
        vals = []
        for i, v in enumerate(raw):
            cell = f"{L(qcols[i][0])}{r}"
            if v is None:
                vals.append(None)
            elif is_num(v) and v >= 0:
                vals.append(float(v))
            else:
                vals.append(None)
                issue("Warning", "Q2", cfg.SHEET_MARKS, cell, nm, f"{quarters[i]}: {v!r} is not a valid number", "Excluded from the series (shown as a gap)")
        first_i = next((i for i, v in enumerate(raw) if v is not None), None)
        if first_i is None:
            issue("Warning", "Q3", cfg.SHEET_MARKS, f"{L(mkey['company'])}{r}", nm, "No marks at all", "Flagged")
        else:
            for i in range(first_i, len(quarters)):
                if raw[i] is None:
                    issue("Warning", "Q3", cfg.SHEET_MARKS, f"{L(qcols[i][0])}{r}", nm,
                          f"{quarters[i]}: blank after first mark ({quarters[first_i]})", "Gap kept; not interpolated; trend point omitted")
            if p["entry_q"] and first_q:
                exp = max(p["entry_q"], first_q)
                if qkeys[first_i] != exp:
                    issue("Warning", "Q4", cfg.SHEET_MARKS, f"{L(mkey['company'])}{r}", nm,
                          f"First mark is {quarters[first_i]} but entry date implies Q{exp[1]} {exp[0]}", "Flagged")
            fv = vals[first_i]
            if fv is not None and is_num(p["inv"]) and abs(fv - p["inv"]) > cfg.TOLERANCE:
                issue("Info", "Q7", cfg.SHEET_MARKS, f"{L(qcols[first_i][0])}{r}", nm,
                      f"First mark {fv} differs from Invested Capital {p['inv']}", "Information only")
        prev = None
        for i, v in enumerate(vals):
            if v is None:
                continue
            if prev is not None and (qkeys[i][0] * 4 + qkeys[i][1]) - (qkeys[prev[0]][0] * 4 + qkeys[prev[0]][1]) == 1 and prev[1] > 0:
                ch = v / prev[1] - 1
                if abs(ch) > cfg.QOQ_LIMIT:
                    issue("Warning", "Q5", cfg.SHEET_MARKS, f"{L(qcols[i][0])}{r}", nm, f"{quarters[i]} mark moves {ch:+.0%} on the prior quarter", "Flagged")
            prev = (i, v)
        last_i = max((i for i, v in enumerate(vals) if v is not None), default=None)
        ref = lab = None
        if p["status"] == cfg.STATUS_ACTIVE:
            ref, lab = p["val"], "Current Valuation"
        elif p["status"] == cfg.STATUS_EXITED:
            ref, lab = p["real"], "Realised Proceeds"
        if last_i is not None and is_num(ref):
            if last_i != len(quarters) - 1:
                issue("Warning", "Q6", cfg.SHEET_MARKS, f"{L(mkey['company'])}{r}", nm, f"Latest valid mark is {quarters[last_i]}, not {quarters[-1]}", "Compared at latest valid mark")
            if abs(vals[last_i] - ref) > cfg.TOLERANCE:
                issue("Warning", "Q6", cfg.SHEET_MARKS, f"{L(qcols[last_i][0])}{r}", nm,
                      f"Latest mark {vals[last_i]} vs Portfolio {lab} {ref}: difference {vals[last_i] - ref:+.1f}", cfg.Q6_TREATMENT)
        p["marks"] = vals
    for p in rows:
        if p["company"].lower() not in qm_names:
            issue("Warning", "Q1", cfg.SHEET_PORTFOLIO, f"{P('company')}{p['row']}", p["company"], f"No row on {cfg.SHEET_MARKS} tab", "No trend shown")
            p["marks"] = [None] * len(quarters)

    # ---- Validated figures
    good = [p for p in rows if not p["block"]]

    def agg(ps):
        inv, val, real = (sum(p[k] for p in ps) for k in ("inv", "val", "real"))
        return dict(n=len(ps), inv=inv, val=val, real=real, tv=val + real,
                    moic=(val + real) / inv if inv else None, dpi=real / inv if inv else None)

    xl = lambda ps, k: sum(p[k] for p in ps if is_num(p[k]))  # Excel SUMIF treats blanks as 0

    fund_out = []
    for f in funds:
        allf = [p for p in rows if p["fund"] == f["name"]]
        gf = [p for p in good if p["fund"] == f["name"]]
        v = agg(gf) if gf else None
        wi, wr = xl(allf, "inv"), xl(allf, "real")
        wv = xl(allf, "val") + wr
        wk = dict(inv=wi, tv=wv, moic=wv / wi if wi else None, dpi=wr / wi if wi else None)
        fund_out.append(dict(name=f["name"], vintage=f["vintage"], irr=f["irr"], validated=v, workbook=wk,
                             rows_total=len(allf), rows_valid=len(gf)))
        if not allf:
            issue("Warning", "X1", cfg.SHEET_FUNDS, f"{F('fund')}{f['row']}", f["name"], "Fund has no rows on Portfolio tab", "Flagged")
        if v and (abs(v["inv"] - wk["inv"]) > cfg.TOLERANCE or abs(v["tv"] - wk["tv"]) > cfg.TOLERANCE):
            issue("Warning", "F3", cfg.SHEET_FUNDS, f"{F('total_invested')}{f['row']}:{F('dpi')}{f['row']}", f["name"],
                  f"Workbook formulas give invested {wk['inv']:.1f}, value {wk['tv']:.1f}, MOIC {wk['moic']:.2f}x over all {len(allf)} rows; "
                  f"validated rows ({len(gf)}) give invested {v['inv']:.1f}, value {v['tv']:.1f}, MOIC {v['moic']:.2f}x",
                  "Dashboard shows validated figures with the workbook figures beside them")

    trend = []
    for f in funds:
        gf = [p for p in good if p["fund"] == f["name"]]
        pts = []
        for i, q in enumerate(quarters):
            exp = [p for p in gf if next((j for j, v in enumerate(p["marks"]) if v is not None), 10**6) <= i]
            have = [p for p in exp if p["marks"][i] is not None]
            complete = bool(exp) and len(have) == len(exp)
            pts.append(dict(q=q, v=sum(p["marks"][i] for p in have) if complete else None, n=len(have), expected=len(exp)))
        trend.append(dict(fund=f["name"], pts=pts))

    def mix(key):
        d = {}
        for p in good:
            d[p[key]] = d.get(p[key], 0) + p["inv"]
        return sorted(({"k": k, "v": v} for k, v in d.items()), key=lambda x: -x["v"])

    comp_issues = {}
    for i in issues:
        comp_issues.setdefault(i["entity"], []).append(i["rule"])
    companies = [dict(
        company=p["company"], fund=p["fund"], sector=p["sector"], geo=p["geo"], status=p["status"],
        entry=p["entry"].date().isoformat() if isinstance(p["entry"], dt.datetime) else None,
        inv=p["inv"] if is_num(p["inv"]) else None, val=p["val"] if is_num(p["val"]) else None,
        real=p["real"] if is_num(p["real"]) else None, moic=p["moic"] if not p["block"] else None,
        excluded=bool(p["block"]), block=sorted(set(p["block"])), rules=sorted(set(comp_issues.get(p["company"], []))),
        row=p["row"], last_mark=next((v for v in reversed(p["marks"]) if v is not None), None)) for p in rows]

    order = {"Blocking": 0, "Warning": 1, "Info": 2}
    issues.sort(key=lambda i: (order[i["sev"]], i["rule"], i["sheet"], i["cell"]))
    wk_all = dict(inv=xl(rows, "inv"), tv=xl(rows, "val") + xl(rows, "real"))
    wk_all["moic"] = wk_all["tv"] / wk_all["inv"] if wk_all["inv"] else None
    wk_all["dpi"] = xl(rows, "real") / wk_all["inv"] if wk_all["inv"] else None
    data = dict(
        generated=TODAY.isoformat(), source=path.split("/")[-1], rows_total=len(rows), rows_valid=len(good),
        sev={s: sum(1 for i in issues if i["sev"] == s) for s in order},
        headline=agg(good) if good else None, workbook_all=wk_all, funds=fund_out,
        sector=mix("sector"), geo=mix("geo"), companies=companies, quarters=quarters, trend=trend, issues=issues,
        rules=[dict(id=a, where=b, rule=c, sev=d) for a, b, c, d in cfg.RULES],
        excluded=[dict(company=c["company"], rules=c["block"]) for c in companies if c["excluded"]])
    html = open(tpl, encoding="utf-8").read().replace("__DATA__", json.dumps(data, ensure_ascii=False))
    open(out, "w", encoding="utf-8").write(html)
    print(json.dumps(dict(valid=len(good), total=len(rows), severity=data["sev"], headline=data["headline"],
                          workbook_all=wk_all, excluded=data["excluded"]), indent=1, default=str))
    print("\nISSUES")
    for i in issues:
        print(f"[{i['sev']}] {i['rule']} {i['sheet']} {i['cell']} | {i['entity']} | {i['found']}")


if __name__ == "__main__":
    main(*sys.argv[1:5])
