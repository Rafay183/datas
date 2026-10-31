"""Compute Table-2 'Actual to date' from a CFR, with a full reconciliation.

Matching rule: a CFR code belongs to a detail line when the code STARTS WITH one
of the line's patterns (case-insensitive). This equals the source workbook's
SUMIF(…,"*"&code&"*",…) because every pattern starts with a letter and codes
are letter + 5 characters, so a pattern can only match at the start.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

from .charter import Charter
from .cfr import CFRData


@dataclass
class Table2Result:
    table: pd.DataFrame            # one row per line, in report order
    detail: pd.DataFrame           # one row per CFR code with its line
    recon: pd.DataFrame            # reconciliation lines
    checks: List[dict]             # PASS/FAIL checks
    charter_check: pd.DataFrame    # CFR codes vs charter
    unmapped: pd.DataFrame
    warnings: List[str] = field(default_factory=list)


def _match(code: str, patterns: List[str]) -> bool:
    c = code.upper()
    return any(c.startswith(p.upper()) for p in patterns)


def compute(cfr: CFRData, mapping: dict, charter: Optional[Charter] = None,
            unit: float = 1e6, receipts_positive: bool = True) -> Table2Result:
    lines = mapping["lines"]
    ids = {l["id"]: l for l in lines}
    excl = mapping.get("exclusions", [])
    sign = -1.0 if receipts_positive else 1.0

    # aggregate CFR by code (duplicates are separate IFMIS entries -> summed)
    agg = (cfr.rows.groupby("Code", sort=True)
           .agg(Description=("Description", "first"), Amount=("Amount", "sum"),
                Entries=("Row", "size"), Rows=("Row", lambda s: ", ".join(map(str, s))))
           .reset_index())

    detail_lines = [l for l in lines if l["type"] == "detail"]
    recs, warnings = [], []
    for _, r in agg.iterrows():
        hits = [l["id"] for l in detail_lines if _match(r["Code"], l.get("patterns", []))]
        ex = next((e for e in excl if r["Code"].upper().startswith(e["pattern"].upper())), None)
        if len(hits) > 1:
            status = "Double-mapped"
        elif hits:
            status = "Mapped"
        elif ex:
            status = "Excluded by rule"
        else:
            status = "Unmapped"
        recs.append({"Code": r["Code"], "CFR Description": r["Description"], "Amount (Rs)": r["Amount"],
                     "Entries": r["Entries"], "CFR Rows": r["Rows"], "Status": status,
                     "Line": ", ".join(ids[h]["label"] for h in hits),
                     "Line Id": ",".join(hits),
                     "Rule": ex["reason"] if (ex and not hits) else ""})
    detail = pd.DataFrame(recs)

    # detail line values (in rupees, CFR sign)
    raw: Dict[str, float] = {l["id"]: 0.0 for l in lines}
    count: Dict[str, int] = {l["id"]: 0 for l in lines}
    for rec in recs:
        if rec["Status"] == "Mapped":
            raw[rec["Line Id"]] += rec["Amount (Rs)"]
            count[rec["Line Id"]] += 1

    def total(i):
        l = ids[i]
        if l["type"] == "detail":
            return raw[i], count[i]
        s, n = 0.0, 0
        for c in l.get("children", []):
            a, b = total(c)
            s += a
            n += b
        raw[i], count[i] = s, n
        return s, n

    root = next(l["id"] for l in lines if l["id"] not in {c for x in lines for c in x.get("children", [])})
    total(root)

    tab = pd.DataFrame([{
        "Line Id": l["id"], "Code": l.get("code", ""), "Description": l["label"], "Level": l["level"],
        "Type": l["type"], "Actual to date": sign * raw[l["id"]] / unit, "Actual (Rs)": sign * raw[l["id"]],
        "CFR codes": count[l["id"]], "Patterns": ", ".join(l.get("patterns", [])), "Note": l.get("note", ""),
    } for l in lines])

    # reconciliation
    cfr_total = float(cfr.rows["Amount"].sum())
    by = detail.groupby("Status")["Amount (Rs)"].sum() if len(detail) else pd.Series(dtype=float)
    mapped = float(by.get("Mapped", 0.0))
    recon = pd.DataFrame([
        {"Item": "CFR grand total (all rows)", "Codes": len(detail), "Rs": cfr_total},
        {"Item": "Mapped to Table-2 lines", "Codes": int((detail["Status"] == "Mapped").sum()), "Rs": mapped},
        {"Item": "Excluded by rule", "Codes": int((detail["Status"] == "Excluded by rule").sum()), "Rs": float(by.get("Excluded by rule", 0.0))},
        {"Item": "Unmapped (not in any line)", "Codes": int((detail["Status"] == "Unmapped").sum()), "Rs": float(by.get("Unmapped", 0.0))},
        {"Item": "Double-mapped (ERROR)", "Codes": int((detail["Status"] == "Double-mapped").sum()), "Rs": float(by.get("Double-mapped", 0.0))},
        {"Item": "Table-2 TOTAL line", "Codes": count[root], "Rs": raw[root]},
    ])
    recon[f"Signed ÷ unit"] = sign * recon["Rs"] / unit
    checks = []

    def add(name, ok, det):
        checks.append({"Check": name, "Result": "PASS" if ok else "FAIL", "Detail": det})
    add("Table-2 total equals mapped CFR codes", abs(raw[root] - mapped) < 0.5,
        f"Total line {raw[root]:,.0f} vs mapped {mapped:,.0f}")
    add("Every CFR rupee accounted for", abs(cfr_total - sum(by)) < 0.5,
        f"CFR {cfr_total:,.0f} = mapped + excluded + unmapped + double ({sum(by):,.0f})")
    add("No code counted in two lines", (detail["Status"] == "Double-mapped").sum() == 0,
        f"{int((detail['Status'] == 'Double-mapped').sum())} double-mapped codes")
    um = detail[detail["Status"] == "Unmapped"]
    add("No unmapped code with a value", (um["Amount (Rs)"].abs() > 0).sum() == 0,
        f"{len(um)} unmapped codes, {int((um['Amount (Rs)'].abs() > 0).sum())} with a non-zero amount "
        f"(Rs {um['Amount (Rs)'].sum():,.0f})")

    # charter check
    cc = pd.DataFrame()
    if charter is not None:
        lk = charter.lookup()
        rows = []
        for rec in recs:
            ch = lk.get(rec["Code"].upper())
            if ch is None:
                st, cdesc, cmin, cst = "Not in charter", "", "", ""
            else:
                cdesc, cmin, cst = ch["Description"], ch["Minor"], ch["Status"]
                a, b = _norm(rec["CFR Description"]), _norm(cdesc)
                st = "Same" if a == b else ("Same (CFR text shortened)" if b.startswith(a) or a.startswith(b) else "Different wording")
                if cst == "Blocked":
                    st = "Blocked in charter"
            rows.append({"Code": rec["Code"], "CFR Description": rec["CFR Description"],
                         "Charter Description": cdesc, "Charter Minor": cmin, "Charter Status": cst,
                         "Check": st, "Amount (Rs)": rec["Amount (Rs)"], "Table-2 Line": rec["Line"]})
        cc = pd.DataFrame(rows)
        n_missing = int((cc["Check"] == "Not in charter").sum())
        if n_missing:
            warnings.append(f"{n_missing} CFR codes are not in the charter file {charter.file}.")
    # suggestions for unmapped codes: lines whose code equals the code's minor
    um = detail[detail["Status"] == "Unmapped"].copy()
    if len(um):
        parent = {c: l["id"] for l in lines for c in l.get("children", [])}

        def section(i):
            while i in parent and ids[parent[i]]["level"] >= 1:
                i = parent[i]
            return ids[i]["label"] if ids[i]["level"] == 1 else ""
        sugg = []
        for code in um["Code"]:
            cands = [f"{l['label']} [{section(l['id'])}]" for l in detail_lines
                     if l.get("code") and code.upper().startswith(l["code"].upper())]
            sugg.append(" | ".join(cands) if cands else "(no line with this minor head)")
        um["Suggested line(s)"] = sugg
    for _, r in um.iterrows():
        if abs(r["Amount (Rs)"]) > 0:
            warnings.append(f"Unmapped code {r['Code']} ({r['CFR Description']}) Rs {r['Amount (Rs)']:,.0f} is not in Table-2.")
    return Table2Result(tab, detail, recon, checks, cc, um, warnings)


def _norm(s: str) -> str:
    import re
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def compare(a: Table2Result, b: Table2Result, label_a: str, label_b: str) -> pd.DataFrame:
    t = a.table[["Line Id", "Code", "Description", "Level", "Type", "Actual to date"]].rename(
        columns={"Actual to date": label_a})
    t[label_b] = b.table.set_index("Line Id").loc[t["Line Id"], "Actual to date"].values
    t["Difference"] = t[label_b] - t[label_a]
    return t
