"""Folder layout and shared constants.

resources/
├── charter/          Chart of Accounts workbooks (original + updated versions)
├── civil_accounts/   CFR workbooks (original + updated versions)
├── mapping/          table2_mapping.json (+ history/ backups)
├── logs/             change_log.csv  (append-only audit trail)
├── outputs/          saved Table-2 exports (optional)
├── other_uploads/    any other uploaded workbook/CSV (also imported into the DB)
├── table2.db         SQLite database mirroring everything
└── .app_state.json   which file is active, file fingerprints seen
"""
from __future__ import annotations

import os
from pathlib import Path

APP_VERSION = "1.0.0"

DEFAULT_RESOURCES = Path(os.environ.get("T2_RESOURCES", Path(__file__).resolve().parent.parent / "resources"))

CHARTER_DIR = "charter"
CFR_DIR = "civil_accounts"
OTHER_DIR = "other_uploads"
MAPPING_DIR = "mapping"
LOG_DIR = "logs"
OUTPUT_DIR = "outputs"
STATE_FILE = ".app_state.json"
LOG_FILE = "change_log.csv"
MAPPING_FILE = "table2_mapping.json"

EXCEL_EXT = (".xlsx", ".xlsm", ".xls")

# Object-code sheets of the NAM Chart of Accounts (sheet name starts with the letter)
OBJECT_SHEET_PREFIXES = ("A=", "B=", "C=", "E=", "F=", "G=", "H=")

UNITS = {"Rupees": 1.0, "Thousand": 1e3, "Million": 1e6, "Billion": 1e9}


def resource_paths(root: Path | str | None = None) -> dict[str, Path]:
    root = Path(root) if root else DEFAULT_RESOURCES
    p = {
        "root": root,
        "charter": root / CHARTER_DIR,
        "cfr": root / CFR_DIR,
        "mapping": root / MAPPING_DIR,
        "logs": root / LOG_DIR,
        "outputs": root / OUTPUT_DIR,
        "other": root / OTHER_DIR,
    }
    for k, v in p.items():
        v.mkdir(parents=True, exist_ok=True)
    (p["mapping"] / "history").mkdir(exist_ok=True)
    return p
