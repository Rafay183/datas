"""Folder tracking: list workbooks, fingerprint them, remember the active one,
detect new/changed/removed files, and create versioned file names."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import Workbook, load_workbook

from .config import EXCEL_EXT, STATE_FILE


@dataclass
class TrackedFile:
    kind: str            # "charter" | "cfr"
    path: Path
    size: int
    modified: datetime
    sha256: str
    parent: str = ""     # file this one was derived from (for app-made versions)
    note: str = ""

    @property
    def name(self) -> str:
        return self.path.name


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class State:
    """Small JSON file in resources/ that remembers active files and fingerprints."""

    def __init__(self, root: Path):
        self.path = Path(root) / STATE_FILE
        self.data = {"active": {}, "seen": {}, "lineage": {}}
        if self.path.exists():
            try:
                self.data.update(json.loads(self.path.read_text(encoding="utf-8")))
            except Exception:
                pass  # corrupt state: start fresh, files are the source of truth

    def save(self) -> None:
        self.path.write_text(json.dumps(self.data, indent=1), encoding="utf-8")

    # active file per kind ("" = automatic: newest)
    def get_active(self, kind: str) -> str:
        return self.data["active"].get(kind, "")

    def set_active(self, kind: str, name: str) -> None:
        self.data["active"][kind] = name
        self.save()

    def lineage(self, name: str) -> dict:
        return self.data["lineage"].get(name, {})

    def add_lineage(self, name: str, parent: str, note: str) -> None:
        self.data["lineage"][name] = {"parent": parent, "note": note,
                                      "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        self.save()


def scan(folder: Path, kind: str, state: State) -> List[TrackedFile]:
    out = []
    for p in sorted(Path(folder).iterdir()):
        if not p.is_file() or p.name.startswith(("~$", ".")) or p.suffix.lower() not in EXCEL_EXT:
            continue
        st = p.stat()
        lin = state.lineage(p.name)
        out.append(TrackedFile(kind, p, st.st_size, datetime.fromtimestamp(st.st_mtime), sha256_of(p),
                               lin.get("parent", ""), lin.get("note", "")))
    return sorted(out, key=lambda f: f.modified, reverse=True)


def detect_changes(files: List[TrackedFile], kind: str, state: State) -> List[dict]:
    """Compare with fingerprints from the previous scan. Returns change events and
    updates the state. Events: new, changed (same name, different content), removed."""
    seen: Dict[str, str] = state.data["seen"].setdefault(kind, {})
    events = []
    current = {f.name: f.sha256 for f in files}
    for name, sha in current.items():
        if name not in seen:
            events.append({"event": "new", "file": name, "sha": sha})
        elif seen[name] != sha:
            events.append({"event": "changed", "file": name, "old": seen[name], "sha": sha})
    for name in list(seen):
        if name not in current:
            events.append({"event": "removed", "file": name, "old": seen[name]})
    state.data["seen"][kind] = current
    state.save()
    return events


def resolve_active(files: List[TrackedFile], kind: str, state: State) -> Optional[TrackedFile]:
    if not files:
        return None
    want = state.get_active(kind)
    if want:
        for f in files:
            if f.name == want:
                return f
    return files[0]  # newest


_VER_RX = re.compile(r"\s*\(updated \d{4}-\d{2}-\d{2} \d{4}(?:\d{2})?(?: \d+)?\)$")


def versioned_name(folder: Path, original_name: str) -> Path:
    """'cfr.xlsx' -> 'cfr (updated 2026-10-31 1455).xlsx' (never overwrites)."""
    stem = _VER_RX.sub("", Path(original_name).stem)
    stamp = datetime.now().strftime("%Y-%m-%d %H%M")
    cand = Path(folder) / f"{stem} (updated {stamp}).xlsx"
    i = 2
    while cand.exists():
        cand = Path(folder) / f"{stem} (updated {stamp} {i}).xlsx"
        i += 1
    return cand


def safe_copy_into(folder: Path, name: str, data: bytes) -> Path:
    """Save an uploaded file into a folder without overwriting anything."""
    target = Path(folder) / Path(name).name
    i = 2
    while target.exists():
        target = Path(folder) / f"{Path(name).stem} ({i}){Path(name).suffix}"
        i += 1
    target.write_bytes(data)
    return target


def open_workbook(path: Path, data_only: bool = True):
    """openpyxl workbook for .xlsx/.xlsm; .xls is converted in memory via xlrd."""
    path = Path(path)
    if path.suffix.lower() != ".xls":
        return load_workbook(path, data_only=data_only)
    try:
        import xlrd  # noqa: F401
    except ImportError as e:
        raise RuntimeError(f"{path.name} is an old .xls file. Install xlrd (pip install xlrd) "
                           "or save it as .xlsx in Excel.") from e
    import xlrd
    book = xlrd.open_workbook(str(path))
    wb = Workbook()
    wb.remove(wb.active)
    for sh in book.sheets():
        ws = wb.create_sheet(sh.name[:31])
        if sh.visibility:
            ws.sheet_state = "hidden"
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                if v != "":
                    ws.cell(r + 1, c + 1, v)
    return wb


def backup(path: Path, folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    dst = folder / f"{Path(path).stem} {datetime.now():%Y%m%d_%H%M%S}{Path(path).suffix}"
    shutil.copy2(path, dst)
    return dst
