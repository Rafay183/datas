"""One-step intake: recognise an uploaded workbook (CFR or Chart of Accounts),
store it in the right resources folder (without duplicates) and make it active."""
from __future__ import annotations

import hashlib
import io
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from .charter import object_sheets
from .files import open_workbook, safe_copy_into, sha256_of


MAX_CFR_SHEETS = 3   # a real CFR export has 1–3 sheets; bigger workbooks are reports that merely contain a CFR copy


@dataclass
class IntakeResult:
    name: str
    kind: Optional[str]      # "cfr" | "charter" | None
    status: str              # "saved" | "already in folder" | "imported to DB" | "error"
    stored_as: str = ""
    message: str = ""
    cfr_candidate: str = ""  # sheet that looks like a CFR inside an 'other' workbook


def classify(name: str, data: bytes) -> tuple[Optional[str], str]:
    """Look inside the workbook: object-code sheets (A=…, B=…, C=…) → charter;
    a 'Detail Item Desc' header → CFR."""
    suffix = Path(name).suffix.lower() or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(data)
        p = Path(tmp.name)
    try:
        wb = open_workbook(p, data_only=True)
        if len(object_sheets(wb)) >= 3:
            return "charter", f"{len(wb.sheetnames)} sheets incl. object-code sheets ({', '.join(object_sheets(wb)[:3])}…)"
        hits = []
        for ws in wb.worksheets:
            found = False
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60)):
                for c in row:
                    if isinstance(c.value, str) and c.value.strip().lower() == "detail item desc":
                        hits.append((ws.title, c.row))
                        found = True
                        break
                if found:
                    break
        if hits and len(wb.sheetnames) <= MAX_CFR_SHEETS:
            return "cfr", f"'Detail Item Desc' header on sheet {hits[0][0]}, row {hits[0][1]}"
        if hits:
            return None, (f"CANDIDATE:{hits[0][0]}|Workbook has {len(wb.sheetnames)} sheets; sheet "
                          f"'{hits[0][0]}' looks like a CFR but the file is treated as general data.")
        return None, "Neither a CFR ('Detail Item Desc' header) nor a Chart of Accounts (A=…/B=…/C=… sheets)."
    except Exception as e:
        return None, f"Could not open the file: {e}"
    finally:
        p.unlink(missing_ok=True)


def ingest(name: str, data: bytes, folders: dict, state, log, user=None, db=None,
           force_kind: Optional[str] = None) -> IntakeResult:
    if force_kind:
        kind, why = force_kind, f"Marked as {force_kind} by the user"
    else:
        kind, why = classify(name, data) if Path(name).suffix.lower() in (".xlsx", ".xlsm", ".xls") else (None, "CSV/text file")
    candidate = ""
    if kind is None and why.startswith("CANDIDATE:"):
        candidate, why = why[len("CANDIDATE:"):].split("|", 1)
    if kind is None:
        # Not a CFR or charter: keep the file and load every sheet into the database.
        if db is None:
            return IntakeResult(name, None, "error", message=why)
        sha = hashlib.sha256(data).hexdigest()
        try:
            _, tables, new = db.load_any(name, data, sha)
        except Exception as e:
            return IntakeResult(name, None, "error", message=f"Could not import: {e}")
        stored = ""
        if new:
            stored = safe_copy_into(folders["other"], name, data).name
            log.add("Other file imported into database", "database", stored, user=user, field="tables",
                    new_value=", ".join(t["table"] for t in tables))
        msg = "; ".join(f"{t['table']} ({t['rows']} rows × {t['columns']} cols)" for t in tables) or "no data found"
        if candidate:
            msg = why + " Tables: " + msg
        return IntakeResult(name, "other", "imported to DB" if new else "already in DB", stored, msg, candidate)
    folder = folders[kind]
    sha = hashlib.sha256(data).hexdigest()
    existing = next((p for p in sorted(Path(folder).iterdir())
                     if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".xls") and sha256_of(p) == sha), None)
    if existing is not None:
        changed = state.get_active(kind) != existing.name
        state.set_active(kind, existing.name)
        if changed:
            log.add("Uploaded file already in folder – made active", kind, existing.name, user=user,
                    field="sha256", new_value=sha[:16])
        return IntakeResult(name, kind, "already in folder", existing.name, why)
    dst = safe_copy_into(folder, name, data)
    state.set_active(kind, dst.name)
    log.add("File uploaded and made active", kind, dst.name, user=user, field="size",
            new_value=len(data), reason=why)
    return IntakeResult(name, kind, "saved", dst.name, why)
