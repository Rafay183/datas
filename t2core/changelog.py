"""Append-only change log (CSV, opens directly in Excel)."""
from __future__ import annotations

import csv
import getpass
from datetime import datetime
from pathlib import Path

import pandas as pd

FIELDS = ["timestamp", "user", "action", "target", "file", "location", "field", "old_value", "new_value", "reason"]


class ChangeLog:
    def __init__(self, path: Path, db=None):
        self.path = Path(path)
        self.db = db
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            with self.path.open("w", newline="", encoding="utf-8-sig") as f:
                csv.DictWriter(f, FIELDS).writeheader()

    def add(self, action: str, target: str = "", file: str = "", location: str = "", field: str = "",
            old_value="", new_value="", reason: str = "", user: str | None = None) -> None:
        row = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": user or _default_user(),
            "action": action, "target": target, "file": file, "location": location, "field": field,
            "old_value": _txt(old_value), "new_value": _txt(new_value), "reason": reason,
        }
        with self.path.open("a", newline="", encoding="utf-8-sig") as f:
            csv.DictWriter(f, FIELDS).writerow(row)
        if self.db is not None:
            try:
                self.db.add_log(row)
            except Exception:
                pass  # the CSV log is the master copy; never fail an action because of the DB

    def add_many(self, rows: list[dict], **common) -> None:
        for r in rows:
            self.add(**{**common, **r})

    def read(self) -> pd.DataFrame:
        try:
            return pd.read_csv(self.path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except Exception:
            return pd.DataFrame(columns=FIELDS)


def _txt(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _default_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return "unknown"
