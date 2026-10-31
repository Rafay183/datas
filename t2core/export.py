"""Formatted Excel export of Table-2 (Actual to date) plus supporting sheets."""
from __future__ import annotations

import io
from datetime import datetime
from typing import Dict, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .table2 import Table2Result

NAVY = "1F4E79"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
LEVEL_FILL = {0: PatternFill("solid", fgColor="BDD7EE"), 1: PatternFill("solid", fgColor="DDEBF7"),
              2: PatternFill("solid", fgColor="F2F2F2")}
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _num_fmt(decimals: int) -> str:
    d = "." + "0" * decimals if decimals else ""
    return f"#,##0{d};[Red]-#,##0{d}"


def _write_df(ws, df: pd.DataFrame, r0: int, num_cols=(), decimals=3, widths: Optional[Dict[str, int]] = None):
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(r0, j, c)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(df.itertuples(index=False), r0 + 1):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and pd.isna(v):
                v = None
            cell = ws.cell(i, j, v.item() if hasattr(v, "item") else v)
            cell.border = BOX
            if df.columns[j - 1] in num_cols:
                cell.number_format = _num_fmt(decimals)
    for j, c in enumerate(df.columns, 1):
        w = (widths or {}).get(c) or min(max([len(str(c))] + [len(str(x)) for x in df[c].head(500)]) + 2, 60)
        ws.column_dimensions[get_column_letter(j)].width = max(w, 10)
    ws.freeze_panes = ws.cell(r0 + 1, 1)
    ws.auto_filter.ref = f"A{r0}:{get_column_letter(len(df.columns))}{r0 + max(len(df), 1)}"


def build_workbook(res: Table2Result, *, unit_name: str, decimals: int, period: str, cfr_file: str,
                   charter_file: str, mapping_version, hide_zero: bool = False,
                   compare_df: Optional[pd.DataFrame] = None, sources: Optional[pd.DataFrame] = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Table 2"
    ws.sheet_view.showGridLines = False
    title = f"TABLE 2 REVENUE BY OBJECT – ACTUAL TO DATE{(' (' + period + ')') if period else ''}"
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=13, color=NAVY)
    ws["A2"] = f"( Rs. in {unit_name.lower()} )   Source: {cfr_file}   •   Generated {datetime.now():%d %b %Y %H:%M}"
    ws["A2"].font = Font(italic=True, size=9, color="595959")
    hdr = ["Code", "Description", f"Actual to date (Rs {unit_name.lower()})"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(4, j, h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 30
    r = 5
    t = res.table
    if hide_zero:
        t = t[(t["Type"] == "total") | (t["Actual to date"].round(decimals) != 0)]
    for _, row in t.iterrows():
        vals = [row["Code"] if row["Type"] == "detail" else "", ("   " * int(row["Level"])) + row["Description"],
                float(row["Actual to date"])]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.border = BOX
            if row["Type"] == "total":
                c.font = Font(bold=True)
                c.fill = LEVEL_FILL.get(int(row["Level"]), LEVEL_FILL[2])
        ws.cell(r, 3).number_format = _num_fmt(decimals)
        r += 1
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 66
    ws.column_dimensions["C"].width = 26
    ws.freeze_panes = "A5"
    ws.print_title_rows = "4:4"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    if compare_df is not None:
        wc = wb.create_sheet("Comparison")
        _write_df(wc, compare_df.drop(columns=["Line Id", "Level", "Type"]), 1,
                  num_cols=[c for c in compare_df.columns if c not in ("Line Id", "Code", "Description", "Level", "Type")],
                  decimals=decimals, widths={"Description": 60})

    wd = wb.create_sheet("Code Detail")
    d = res.detail.drop(columns=["Line Id"])
    _write_df(wd, d, 1, num_cols=["Amount (Rs)"], decimals=0, widths={"Line": 50, "CFR Rows": 30})

    wr = wb.create_sheet("Reconciliation")
    _write_df(wr, res.recon.rename(columns={"Signed ÷ unit": f"Rs {unit_name.lower()} (receipts +)"}), 1,
              num_cols=["Rs", f"Rs {unit_name.lower()} (receipts +)"], decimals=decimals)
    chk = pd.DataFrame(res.checks)
    r0 = len(res.recon) + 4
    _write_df(wr, chk, r0)
    for i, v in enumerate(chk["Result"], r0 + 1):
        wr.cell(i, 2).fill = PatternFill("solid", fgColor="C6EFCE" if v == "PASS" else "FFC7CE")
    wr.freeze_panes = None

    if not res.unmapped.empty:
        wu = wb.create_sheet("Unmapped Codes")
        _write_df(wu, res.unmapped.drop(columns=["Line Id"], errors="ignore"), 1, num_cols=["Amount (Rs)"], decimals=0)

    if not res.charter_check.empty:
        wcc = wb.create_sheet("Charter Check")
        _write_df(wcc, res.charter_check, 1, num_cols=["Amount (Rs)"], decimals=0,
                  widths={"Charter Description": 60, "Table-2 Line": 45})

    wm = wb.create_sheet("Line Mapping")
    _write_df(wm, res.table[["Line Id", "Code", "Description", "Type", "CFR codes", "Patterns", "Note"]], 1,
              widths={"Patterns": 80, "Description": 55})

    wsrc = wb.create_sheet("Sources")
    info = pd.DataFrame([
        {"Item": "CFR file", "Value": cfr_file}, {"Item": "Period", "Value": period},
        {"Item": "Charter file", "Value": charter_file}, {"Item": "Mapping version", "Value": str(mapping_version)},
        {"Item": "Unit", "Value": unit_name}, {"Item": "Sign", "Value": "Receipts shown positive (CFR × −1)"},
        {"Item": "Generated", "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ])
    _write_df(wsrc, info, 1, widths={"Value": 70})
    if sources is not None and not sources.empty:
        _write_df(wsrc, sources, len(info) + 4, widths={"SHA-256": 70})
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
