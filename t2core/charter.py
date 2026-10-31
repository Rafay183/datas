"""Read and update the NAM Chart of Accounts ('charter') workbook.

Object sheets (A=Expenditure, B=Tax Revenue, C=Non Tax Revenue, E=Capital
Receipts, F=Assets, G=Liabilities, H=Equities) share one layout:

    A Major code | B Major desc | C Minor code | D Minor desc | E New (detailed) code | F Description | G Old code / CS No.

Major/minor are written only on the first row of their group, so they are
forward-filled. Rows with F but no E are group headings and are skipped.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List

import pandas as pd
from openpyxl import load_workbook

from .config import OBJECT_SHEET_PREFIXES
from .files import open_workbook

COLS = ["Sheet", "Row", "Major", "Major Description", "Minor", "Minor Description",
        "Code", "Description", "Old Code / CS", "Status"]


@dataclass
class Charter:
    file: str
    codes: pd.DataFrame                        # one row per detailed code
    sheets: List[str]
    issues: List[str] = field(default_factory=list)

    def lookup(self) -> Dict[str, dict]:
        """CODE -> first record (codes are matched case-insensitively)."""
        out = {}
        for rec in self.codes.to_dict("records"):
            out.setdefault(str(rec["Code"]).upper(), rec)
        return out


def _s(v):
    return "" if v is None else " ".join(str(v).split())


def status_of(old: str) -> str:
    o = (old or "").lower()
    if "unblock" in o:
        return "Unblocked"
    if "block" in o:
        return "Blocked"
    if "proposed" in o:
        return "Proposed"
    return "Active"


def object_sheets(wb) -> List[str]:
    return [n for n in wb.sheetnames if n.startswith(OBJECT_SHEET_PREFIXES)]


def parse_charter(path: Path) -> Charter:
    wb = open_workbook(Path(path), data_only=True)
    names = object_sheets(wb)
    if not names:
        raise ValueError(f"{Path(path).name}: no object-code sheets (A=…, B=…, C=…) found.")
    recs, issues = [], []
    for name in names:
        ws = wb[name]
        maj = majd = mn = mnd = ""
        for r in range(3, ws.max_row + 1):
            A, B, C, D, E, F, G = [ws.cell(r, c).value for c in range(1, 8)]
            if A is not None and _s(A):
                maj, majd = _s(A).upper(), _s(B)
            if C is not None and _s(C):
                mn, mnd = _s(C).upper(), _s(D)
            if E is None or not _s(E):
                continue
            code = _s(E)
            recs.append({"Sheet": name.strip(), "Row": r, "Major": maj, "Major Description": majd,
                         "Minor": mn, "Minor Description": mnd, "Code": code, "Description": _s(F),
                         "Old Code / CS": _s(G), "Status": status_of(_s(G))})
            if not re.fullmatch(r"[A-Z]\d{2}[0-9A-Z]{3}", code):
                issues.append(f"{name.strip()} row {r}: code '{code}' is not in the standard 6-character format.")
    df = pd.DataFrame(recs, columns=COLS)
    dup = df[df.duplicated("Code", keep=False)]
    for code, g in dup.groupby("Code"):
        issues.append(f"Code {code} appears {len(g)} times ({g['Sheet'].iloc[0]} rows {', '.join(map(str, g['Row']))}).")
    return Charter(Path(path).name, df, [n.strip() for n in names], issues)


# --------------------------------------------------------------------------- editing
def diff_codes(before: pd.DataFrame, after: pd.DataFrame, sheet: str) -> List[dict]:
    """Changes between original rows of one sheet and the edited grid.
    Grid columns: Row, Major, Minor, Code, Description, Old Code / CS (Row blank = new)."""
    ch = []
    b = before.set_index("Row")
    kept = set()
    for _, r in after.iterrows():
        row = r.get("Row")
        if row is None or pd.isna(row) or row == "":
            ch.append({"op": "add", "action": "Charter code added", "location": f"{sheet}", "field": "Code",
                       "old_value": "", "new_value": f"{r['Code']} – {r['Description']} (minor {r['Minor']})",
                       "rec": r.to_dict()})
            continue
        row = int(row)
        kept.add(row)
        if row not in b.index:
            continue
        for fld in ("Code", "Description", "Old Code / CS"):
            old, new = _s(b.at[row, fld]), _s(r.get(fld))
            if old != new:
                ch.append({"op": "edit", "action": "Charter code edited", "location": f"{sheet} row {row} ({b.at[row, 'Code']})",
                           "field": fld, "old_value": old, "new_value": new, "row": row, "col": fld})
    for row in b.index:
        if row not in kept:
            ch.append({"op": "delete", "action": "Charter code deleted", "location": f"{sheet} row {row}",
                       "field": "Code", "old_value": f"{b.at[row, 'Code']} – {b.at[row, 'Description']}",
                       "new_value": "", "row": row})
    return ch


def validate_additions(changes: List[dict], charter: Charter) -> List[str]:
    errs = []
    existing = set(charter.codes["Code"].str.upper())
    for c in changes:
        if c["op"] == "add":
            r = c["rec"]
            code = _s(r.get("Code")).upper()
            if not re.fullmatch(r"[A-Z]\d{2}[0-9A-Z]{3}", code):
                errs.append(f"New code '{r.get('Code')}' must be 6 characters like C03923.")
            elif code in existing:
                errs.append(f"New code {code} already exists in the charter.")
            minor = _s(r.get("Minor")).upper()
            if minor and not code.startswith(minor):
                errs.append(f"New code {code} does not start with its minor head {minor}.")
            if not _s(r.get("Description")):
                errs.append(f"New code {code}: description is required.")
        if c["op"] == "edit" and c["col"] == "Code":
            nv = c["new_value"].upper()
            if not re.fullmatch(r"[A-Z]\d{2}[0-9A-Z]{3}", nv):
                errs.append(f"Edited code '{c['new_value']}' is not a valid 6-character code.")
            elif nv in existing:
                errs.append(f"Edited code {nv} already exists in the charter.")
    return errs


def _shift_merged(ws, at_row: int) -> list:
    """openpyxl insert/delete_rows does not move merged ranges: unmerge every range
    at/below `at_row` and return it so _remerge can put it back shifted."""
    moved = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= at_row:
            moved.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
            ws.unmerge_cells(str(rng))
    return moved


def _remerge(ws, moved, n):
    from openpyxl.utils import get_column_letter as L
    for r1, c1, r2, c2 in moved:
        ws.merge_cells(f"{L(c1)}{r1 + n}:{L(c2)}{r2 + n}")


def save_updated(src: Path, charter: Charter, sheet: str, changes: List[dict], target: Path) -> Path:
    """Apply edits/additions/deletions of ONE object sheet to a copy of the source
    workbook and save it as `target`. The source file is not modified."""
    src = Path(src)
    wb = load_workbook(src) if src.suffix.lower() != ".xls" else open_workbook(src, data_only=False)
    full = next(n for n in wb.sheetnames if n.strip() == sheet.strip())
    ws = wb[full]
    colmap = {"Code": 5, "Description": 6, "Old Code / CS": 7}
    # 1) in-place edits
    for c in changes:
        if c["op"] == "edit":
            ws.cell(c["row"], colmap[c["col"]]).value = c["new_value"]
    # 2) deletions (bottom-up so row numbers stay valid)
    for c in sorted([c for c in changes if c["op"] == "delete"], key=lambda c: -c["row"]):
        r = c["row"]
        # If the deleted row carries the major/minor label (cols A–D) and the next
        # row is a code of the same group without its own label, move the label down
        # so the group keeps its heading.
        nxt = r + 1
        if ws.cell(nxt, 5).value not in (None, ""):
            for col in (1, 2, 3, 4):
                v = ws.cell(r, col).value
                if v not in (None, "") and ws.cell(nxt, col).value in (None, ""):
                    ws.cell(nxt, col).value = v
        for rng in list(ws.merged_cells.ranges):       # a merge spanning the deleted row
            if rng.min_row <= r <= rng.max_row:          # cannot survive; unmerge it
                ws.unmerge_cells(str(rng))
        moved = _shift_merged(ws, r + 1)
        ws.delete_rows(r, 1)
        _remerge(ws, moved, -1)
    # 3) additions: insert after the last row of the same minor (or major), else append
    for c in [c for c in changes if c["op"] == "add"]:
        rec = c["rec"]
        minor, major = _s(rec.get("Minor")).upper(), _s(rec.get("Major")).upper()
        live = _current_rows(ws)
        same_minor = [r for r, (mj, mn) in live.items() if mn == minor and minor]
        same_major = [r for r, (mj, mn) in live.items() if mj == major and major]
        if same_minor:
            at, write_minor, write_major = max(same_minor) + 1, False, False
        elif same_major:
            at, write_minor, write_major = max(same_major) + 1, True, False
        else:
            at, write_minor, write_major = ws.max_row + 1, True, True
        spanning = [(g.min_row, g.min_col, g.max_row, g.max_col) for g in list(ws.merged_cells.ranges)
                    if g.min_row < at <= g.max_row]
        for g in spanning:
            ws.unmerge_cells(start_row=g[0], start_column=g[1], end_row=g[2], end_column=g[3])
        moved = _shift_merged(ws, at)
        ws.insert_rows(at, 1)
        _remerge(ws, moved, 1)
        for r1, c1, r2, c2 in spanning:                  # stretch merges that span the new row
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2 + 1, end_column=c2)
        known = charter.codes
        maj_desc = _s(rec.get("Major Description")) or next(iter(known.loc[known["Major"] == major, "Major Description"]), "")
        min_desc = _s(rec.get("Minor Description")) or next(iter(known.loc[known["Minor"] == minor, "Minor Description"]), "")
        if write_major:
            ws.cell(at, 1).value = major
            ws.cell(at, 2).value = maj_desc
        if write_minor:
            ws.cell(at, 3).value = minor
            ws.cell(at, 4).value = min_desc
        ws.cell(at, 5).value = _s(rec.get("Code")).upper()
        ws.cell(at, 6).value = _s(rec.get("Description"))
        ws.cell(at, 7).value = _s(rec.get("Old Code / CS")) or "added via Table-2 app"
        c["location"] = f"{sheet} row {at}"
    wb.save(target)
    return target


def _current_rows(ws) -> Dict[int, tuple]:
    out, maj, mn = {}, "", ""
    for r in range(3, ws.max_row + 1):
        A, C, E = ws.cell(r, 1).value, ws.cell(r, 3).value, ws.cell(r, 5).value
        if A not in (None, "") and _s(A):
            maj = _s(A).upper()
        if C not in (None, "") and _s(C):
            mn = _s(C).upper()
        if E not in (None, "") and _s(E):
            out[r] = (maj, mn)
    return out
