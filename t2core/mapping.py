"""Table-2 line definitions (labels, structure and code patterns).

The default mapping reproduces the column-E formulas of the source workbook
'June_2025-26 (T01 to T03)' sheet 'Table-2(2QTR)'. It is stored as JSON in
resources/mapping/ so it can be reviewed and changed without touching code.
"""
from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple

from .config import MAPPING_FILE

DEFAULT_PATH = Path(__file__).resolve().parent / "default_mapping.json"


def load_default() -> dict:
    return json.loads(DEFAULT_PATH.read_text(encoding="utf-8"))


def load(mapping_dir: Path) -> Tuple[dict, bool]:
    """Load resources/mapping/table2_mapping.json, creating it from the default
    on first run. Returns (mapping, created)."""
    p = Path(mapping_dir) / MAPPING_FILE
    if not p.exists():
        m = load_default()
        p.write_text(json.dumps(m, indent=1), encoding="utf-8")
        return m, True
    return json.loads(p.read_text(encoding="utf-8")), False


def save(mapping_dir: Path, m: dict) -> Path:
    from datetime import datetime
    p = Path(mapping_dir) / MAPPING_FILE
    if p.exists():
        hist = Path(mapping_dir) / "history" / f"table2_mapping {datetime.now():%Y%m%d_%H%M%S}.json"
        hist.write_text(p.read_text(encoding="utf-8"), encoding="utf-8")
    m = copy.deepcopy(m)
    m["version"] = int(m.get("version", 1)) + 1
    p.write_text(json.dumps(m, indent=1), encoding="utf-8")
    return p


def parse_patterns(text: str) -> List[str]:
    return [t.strip().upper().strip('"').strip("'") for t in re.split(r"[,\s;]+", str(text or "")) if t.strip()]


def validate(m: dict) -> List[str]:
    errs = []
    ids = {l["id"]: l for l in m.get("lines", [])}
    if not ids:
        return ["Mapping has no lines."]
    seen_child = {}
    for l in m["lines"]:
        if l["type"] == "total":
            for c in l.get("children", []):
                if c not in ids:
                    errs.append(f"{l['label']}: child {c} does not exist.")
                if c in seen_child:
                    errs.append(f"Line {c} is a child of both {seen_child[c]} and {l['id']}.")
                seen_child[c] = l["id"]
        else:
            for p in l.get("patterns", []):
                if not re.fullmatch(r"[A-Z]\d{1,2}[0-9A-Z]{0,4}", p):
                    errs.append(f"{l['label']}: pattern '{p}' does not look like a code or code prefix.")
    roots = [l for l in m["lines"] if l["id"] not in seen_child]
    if len(roots) != 1:
        errs.append(f"Mapping must have exactly one top line (found {len(roots)}).")
    # every detail line must be reachable exactly once from the root
    if len(roots) == 1:
        reached = []

        def walk(i):
            l = ids[i]
            if l["type"] == "total":
                for c in l.get("children", []):
                    if c in ids:
                        walk(c)
            else:
                reached.append(i)
        walk(roots[0]["id"])
        details = [l["id"] for l in m["lines"] if l["type"] == "detail"]
        missing = set(details) - set(reached)
        if missing:
            errs.append("Detail lines not under the total: " + ", ".join(sorted(missing)))
    return errs


def line_of_code(m: dict) -> Dict[str, List[str]]:
    """pattern -> [line ids] (for overlap checks)."""
    out: Dict[str, List[str]] = {}
    for l in m["lines"]:
        for p in l.get("patterns", []):
            out.setdefault(p, []).append(l["id"])
    return out
