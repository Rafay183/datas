"""Read and update Civil Account CFR workbooks (IFMIS 'Consolidated Fund Receipts').

Expected layout (any sheet, position detected automatically):
    title rows ... 'FOR THE MONTH OF MARCH 2026'
    header row:   'Detail Item Desc' | ... | 'Progressive March'
    data rows:    'B01108-SHARE OF NET PROCEEDS…' | ... | -894730700000
Amounts are rupees; receipts are negative.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import pandas as pd
from openpyxl import load_workbook

from .files import open_workbook

CODE_RX = re.compile(r"^\s*\**\s*([A-Za-z]\d{2}[0-9A-Za-z]{3})\s*-\s*(.*?)\s*$")
PERIOD_RX = re.compile(r"FOR THE MONTH OF\s+(.+)", re.I)


class CFRFormatError(Exception):
    pass


@dataclass
class CFRData:
    file: str
    sheet: str
    header_row: int          # Excel row number of the header
    code_col: int            # 1-based column holding 'CODE-DESC'
    amount_col: int          # 1-based column holding the amount
    amount_header: str
    period: str
    rows: pd.DataFrame       # Row, Code, Description, Amount, Raw Text
    skipped: pd.DataFrame    # non-code text found in the data area (footers etc.)
    notes: List[str] = field(default_factory=list)

    @property
    def total(self) -> float:
        return float(self.rows["Amount"].sum())


def _num(v):
    if v is None or v == "":
        return math.nan
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace(",", "").strip()
    if t.startswith("(") and t.endswith(")"):
        t = "-" + t[1:-1]
    try:
        return float(t)
    except ValueError:
        return math.nan


def _find_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60)):
        for c in row:
            if isinstance(c.value, str) and c.value.strip().lower() == "detail item desc":
                return c.row, c.column
    return None


def parse_cfr(path: Path) -> CFRData:
    path = Path(path)
    wb = open_workbook(path, data_only=True)
    found = None
    for ws in wb.worksheets:
        h = _find_header(ws)
        if h:
            found = (ws, *h)
            break
    if not found:
        raise CFRFormatError(f"{path.name}: no sheet has a 'Detail Item Desc' header.")
    ws, hr, cc = found
    notes = []
    # amount column: header containing 'prog' (progressive) to the right, else first numeric column
    amount_col, amount_header = None, ""
    for c in range(cc + 1, ws.max_column + 1):
        v = ws.cell(hr, c).value
        if isinstance(v, str) and "prog" in v.lower():
            amount_col, amount_header = c, " ".join(v.split())
            break
    if amount_col is None:
        for c in range(cc + 1, min(ws.max_column, cc + 10) + 1):
            vals = [ws.cell(r, c).value for r in range(hr + 1, min(ws.max_row, hr + 50) + 1)]
            if sum(isinstance(x, (int, float)) for x in vals) >= 5:
                amount_col = c
                amount_header = str(ws.cell(hr, c).value or f"Column {c}")
                notes.append(f"No 'Progressive' header found; using column {c} as amount.")
                break
    if amount_col is None:
        raise CFRFormatError(f"{path.name}: amount column not found next to the header.")
    period = ""
    for r in range(1, hr):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and PERIOD_RX.search(v):
                period = PERIOD_RX.search(v).group(1).strip()
    rows, skipped = [], []
    for r in range(hr + 1, ws.max_row + 1):
        t = ws.cell(r, cc).value
        a = ws.cell(r, amount_col).value
        if t is None or str(t).strip() == "":
            continue
        text = str(t)
        lead = re.match(r"^\s*(\*+)", text)
        stars = len(lead.group(1)) if lead else 0
        if stars >= 2:
            # hierarchical IFMIS report: '**' minor, '***' major, … are subtotals of the '*' detail lines
            skipped.append({"Row": r, "Text": text, "Amount": a, "Reason": f"subtotal level ({'*' * stars})"})
            continue
        m = CODE_RX.match(text)
        if not m:
            skipped.append({"Row": r, "Text": text, "Amount": a, "Reason": "no object code"})
            continue
        val = _num(a)
        rows.append({"Row": r, "Code": m.group(1).upper(), "Description": m.group(2),
                     "Amount": val, "Raw Text": str(t)})
        if math.isnan(val):
            notes.append(f"Row {r} ({m.group(1)}): amount is blank or not a number – counted as 0.")
    df = pd.DataFrame(rows, columns=["Row", "Code", "Description", "Amount", "Raw Text"])
    df["Amount"] = df["Amount"].fillna(0.0)
    return CFRData(path.name, ws.title, hr, cc, amount_col, amount_header, period, df,
                   pd.DataFrame(skipped, columns=["Row", "Text", "Amount", "Reason"]), notes)


# --------------------------------------------------------------------------- editing
def diff_rows(before: pd.DataFrame, after: pd.DataFrame) -> List[dict]:
    """Compare the original rows with an edited grid (columns Row, Code, Description,
    Amount). Rows with a blank 'Row' are new. Returns change records for the log."""
    changes = []
    b = before.set_index("Row")
    after = after.copy()
    kept = set()
    for _, r in after.iterrows():
        row = r.get("Row")
        if row is None or (isinstance(row, float) and math.isnan(row)) or row == "":
            changes.append({"action": "CFR row added", "location": "new row", "field": "Code / Description / Amount",
                            "old_value": "", "new_value": f"{r['Code']}-{r['Description']} | {r['Amount']}"})
            continue
        row = int(row)
        kept.add(row)
        if row not in b.index:
            continue
        for fld in ("Code", "Description", "Amount"):
            old, new = b.at[row, fld], r[fld]
            if fld == "Amount":
                if (pd.isna(old) and pd.isna(new)) or (not pd.isna(old) and not pd.isna(new) and float(old) == float(new)):
                    continue
            elif str(old).strip() == str(new).strip():
                continue
            changes.append({"action": "CFR value edited", "location": f"row {row} ({b.at[row, 'Code']})",
                            "field": fld, "old_value": old, "new_value": new})
    for row in b.index:
        if row not in kept:
            changes.append({"action": "CFR row deleted", "location": f"row {row}", "field": "",
                            "old_value": f"{b.at[row, 'Code']}-{b.at[row, 'Description']} | {b.at[row, 'Amount']}",
                            "new_value": ""})
    return changes


def validate_grid(grid: pd.DataFrame) -> List[str]:
    errs = []
    for i, r in grid.reset_index(drop=True).iterrows():
        code = str(r.get("Code") or "").strip().upper()
        if not re.fullmatch(r"[A-Z]\d{2}[0-9A-Z]{3}", code):
            errs.append(f"Line {i + 1}: code '{r.get('Code')}' is not a 6-character object code (e.g. B01108).")
        if pd.isna(_num(r.get("Amount"))):
            errs.append(f"Line {i + 1} ({code}): amount is not a number.")
    return errs


def save_updated(src: Path, data: CFRData, grid: pd.DataFrame, target: Path) -> Path:
    """Write a NEW workbook: the source file with its data block (code + amount
    columns, below the header) rewritten from `grid`. Titles and every other cell
    are kept. The source file is never modified."""
    src = Path(src)
    if src.suffix.lower() == ".xls":
        wb = open_workbook(src, data_only=False)
    else:
        wb = load_workbook(src)  # keep formulas/styles
    ws = wb[data.sheet]
    last = max(ws.max_row, data.header_row + 1)
    for r in range(data.header_row + 1, last + 1):
        ws.cell(r, data.code_col).value = None
        ws.cell(r, data.amount_col).value = None
    for i, (_, r) in enumerate(grid.iterrows()):
        rr = data.header_row + 1 + i
        ws.cell(rr, data.code_col).value = f"{str(r['Code']).strip().upper()}-{str(r['Description']).strip()}"
        amt = _num(r["Amount"])
        ws.cell(rr, data.amount_col).value = int(amt) if float(amt).is_integer() else amt
        ws.cell(rr, data.amount_col).number_format = "#,##0;-#,##0"
    wb.save(target)
    return target
