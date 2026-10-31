"""SQLite database that mirrors everything the app sees.

File: resources/table2.db  (stdlib sqlite3 – nothing to install; open it with
DB Browser for SQLite, Excel/Power BI ODBC, Python, etc.)

Tables
------
files            one row per workbook ever seen (kind, name, sha256, size, folder, loaded_at)
active_files     which file is active per kind (cfr / charter)
cfr_rows         every CFR data row (file_id, excel_row, code, description, amount)
charter_codes    every Chart of Accounts detailed code (file_id, sheet, major, minor, code, …)
table2_runs      one row per distinct Table-2 result (cfr file + charter file + mapping version)
table2_lines     the 66 lines of each run
imported_tables  catalog of tables created from any OTHER uploaded workbook / CSV
change_log       copy of every change-log entry
raw_*            one table per sheet of any other uploaded file (auto-created)

Views:  v_cfr_active, v_charter_active, v_table2_latest
Re-loading the same file (same SHA-256) is a no-op, so the DB never duplicates data.
"""
from __future__ import annotations

import io
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd

DB_NAME = "table2.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(
    file_id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL, name TEXT NOT NULL, sha256 TEXT NOT NULL, size INTEGER,
    folder TEXT, period TEXT, rows_loaded INTEGER, loaded_at TEXT,
    UNIQUE(kind, sha256));
CREATE TABLE IF NOT EXISTS active_files(kind TEXT PRIMARY KEY, file_id INTEGER, set_at TEXT);
CREATE TABLE IF NOT EXISTS cfr_rows(
    file_id INTEGER, excel_row INTEGER, code TEXT, major TEXT, minor TEXT,
    description TEXT, amount REAL, raw_text TEXT);
CREATE INDEX IF NOT EXISTS ix_cfr_file_code ON cfr_rows(file_id, code);
CREATE TABLE IF NOT EXISTS charter_codes(
    file_id INTEGER, sheet TEXT, excel_row INTEGER, major TEXT, major_description TEXT,
    minor TEXT, minor_description TEXT, code TEXT, description TEXT, old_code_cs TEXT, status TEXT);
CREATE INDEX IF NOT EXISTS ix_charter_file_code ON charter_codes(file_id, code);
CREATE TABLE IF NOT EXISTS table2_runs(
    run_id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT,
    cfr_file_id INTEGER, charter_file_id INTEGER, mapping_version INTEGER,
    period TEXT, total_rs REAL, checks_passed INTEGER, checks_total INTEGER,
    UNIQUE(cfr_file_id, charter_file_id, mapping_version));
CREATE TABLE IF NOT EXISTS table2_lines(
    run_id INTEGER, line_order INTEGER, line_id TEXT, code TEXT, description TEXT,
    level INTEGER, line_type TEXT, actual_rs REAL, actual_million REAL, cfr_codes INTEGER);
CREATE TABLE IF NOT EXISTS imported_tables(
    table_name TEXT PRIMARY KEY, file_id INTEGER, sheet TEXT, header_row INTEGER,
    rows INTEGER, columns INTEGER, created_at TEXT);
CREATE TABLE IF NOT EXISTS change_log(
    timestamp TEXT, user TEXT, action TEXT, target TEXT, file TEXT, location TEXT,
    field TEXT, old_value TEXT, new_value TEXT, reason TEXT);
DROP VIEW IF EXISTS v_cfr_active;
CREATE VIEW v_cfr_active AS
    SELECT f.name AS file, f.period, r.* FROM cfr_rows r
    JOIN active_files a ON a.kind='cfr' AND a.file_id=r.file_id JOIN files f ON f.file_id=r.file_id;
DROP VIEW IF EXISTS v_charter_active;
CREATE VIEW v_charter_active AS
    SELECT f.name AS file, c.* FROM charter_codes c
    JOIN active_files a ON a.kind='charter' AND a.file_id=c.file_id JOIN files f ON f.file_id=c.file_id;
DROP VIEW IF EXISTS v_table2_latest;
CREATE VIEW v_table2_latest AS
    SELECT r.period, r.created_at, l.* FROM table2_lines l
    JOIN table2_runs r ON r.run_id=l.run_id
    WHERE l.run_id=(SELECT MAX(run_id) FROM table2_runs);
"""


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_ident(text: str, maxlen: int = 50) -> str:
    s = re.sub(r"[^0-9a-zA-Z]+", "_", str(text)).strip("_").lower()
    if not s or s[0].isdigit():
        s = "c_" + s
    return s[:maxlen]


class DB:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.conn() as c:
            c.executescript(SCHEMA)

    @contextmanager
    def conn(self):
        c = sqlite3.connect(self.path, timeout=30)
        try:
            yield c
            c.commit()
        finally:
            c.close()

    # ------------------------------------------------------------------ files
    def file_id(self, kind: str, sha: str) -> Optional[int]:
        with self.conn() as c:
            r = c.execute("SELECT file_id FROM files WHERE kind=? AND sha256=?", (kind, sha)).fetchone()
        return r[0] if r else None

    def _register(self, c, kind, name, sha, size, folder, period, rows) -> int:
        cur = c.execute("INSERT INTO files(kind,name,sha256,size,folder,period,rows_loaded,loaded_at) "
                        "VALUES(?,?,?,?,?,?,?,?)", (kind, name, sha, size, folder, period, rows, _now()))
        return cur.lastrowid

    def set_active(self, kind: str, sha: str) -> None:
        fid = self.file_id(kind, sha)
        if fid is None:
            return
        with self.conn() as c:
            c.execute("INSERT INTO active_files(kind,file_id,set_at) VALUES(?,?,?) "
                      "ON CONFLICT(kind) DO UPDATE SET file_id=excluded.file_id, set_at=excluded.set_at",
                      (kind, fid, _now()))

    # ------------------------------------------------------------------ loaders (idempotent)
    def load_cfr(self, tracked, cfr) -> tuple[int, bool]:
        fid = self.file_id("cfr", tracked.sha256)
        if fid:
            return fid, False
        with self.conn() as c:
            fid = self._register(c, "cfr", tracked.name, tracked.sha256, tracked.size, "civil_accounts",
                                 cfr.period, len(cfr.rows))
            c.executemany("INSERT INTO cfr_rows VALUES(?,?,?,?,?,?,?,?)",
                          [(fid, int(row), code, code[:3], code[:4], desc, float(amt), raw)
                           for row, code, desc, amt, raw in cfr.rows[["Row", "Code", "Description", "Amount", "Raw Text"]]
                           .itertuples(index=False, name=None)])
        return fid, True

    def load_charter(self, tracked, charter) -> tuple[int, bool]:
        fid = self.file_id("charter", tracked.sha256)
        if fid:
            return fid, False
        df = charter.codes
        with self.conn() as c:
            fid = self._register(c, "charter", tracked.name, tracked.sha256, tracked.size, "charter", "", len(df))
            c.executemany("INSERT INTO charter_codes VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                          [(fid, r["Sheet"], int(r["Row"]), r["Major"], r["Major Description"], r["Minor"],
                            r["Minor Description"], r["Code"], r["Description"], r["Old Code / CS"], r["Status"])
                           for r in df.to_dict("records")])
        return fid, True

    def load_any(self, name: str, data: bytes, sha: str, folder: str = "other") -> tuple[int, List[dict], bool]:
        """Import ANY workbook/CSV: every sheet becomes table raw_<file>_<sheet>,
        with a detected header row and an extra source_row column."""
        fid = self.file_id("other", sha)
        if fid:
            with self.conn() as c:
                rows = c.execute("SELECT table_name, sheet, rows, columns FROM imported_tables WHERE file_id=?",
                                 (fid,)).fetchall()
            return fid, [dict(zip(("table", "sheet", "rows", "columns"), r)) for r in rows], False
        frames = read_any(name, data)
        made = []
        with self.conn() as c:
            fid = self._register(c, "other", name, sha, len(data), folder, "", sum(len(f) for _, f, _ in frames))
            for sheet, df, hdr in frames:
                base = f"raw_{safe_ident(Path(name).stem, 30)}_{safe_ident(sheet, 25)}"
                tname, i = base, 2
                while c.execute("SELECT 1 FROM sqlite_master WHERE name=?", (tname,)).fetchone():
                    tname, i = f"{base}_{i}", i + 1
                df.to_sql(tname, c, index=False)
                c.execute("INSERT INTO imported_tables VALUES(?,?,?,?,?,?,?)",
                          (tname, fid, sheet, hdr, len(df), len(df.columns), _now()))
                made.append({"table": tname, "sheet": sheet, "rows": len(df), "columns": len(df.columns)})
        return fid, made, True

    # ------------------------------------------------------------------ results / log
    def save_table2(self, cfr_sha, charter_sha, mapping_version, period, res) -> Optional[int]:
        cf = self.file_id("cfr", cfr_sha)
        ch = self.file_id("charter", charter_sha) if charter_sha else None
        if cf is None:
            return None
        with self.conn() as c:
            exists = c.execute("SELECT run_id FROM table2_runs WHERE cfr_file_id=? AND charter_file_id IS ? "
                               "AND mapping_version=?", (cf, ch, mapping_version)).fetchone()
            if exists:
                return exists[0]
            t = res.table
            passed = sum(1 for x in res.checks if x["Result"] == "PASS")
            cur = c.execute("INSERT INTO table2_runs(created_at,cfr_file_id,charter_file_id,mapping_version,period,"
                            "total_rs,checks_passed,checks_total) VALUES(?,?,?,?,?,?,?,?)",
                            (_now(), cf, ch, mapping_version, period, float(t.iloc[0]["Actual (Rs)"]), passed,
                             len(res.checks)))
            rid = cur.lastrowid
            c.executemany("INSERT INTO table2_lines VALUES(?,?,?,?,?,?,?,?,?,?)",
                          [(rid, i, r["Line Id"], r["Code"], r["Description"], int(r["Level"]), r["Type"],
                            float(r["Actual (Rs)"]), float(r["Actual (Rs)"]) / 1e6, int(r["CFR codes"]))
                           for i, r in enumerate(t.to_dict("records"))])
        return rid

    def add_log(self, row: dict) -> None:
        with self.conn() as c:
            c.execute("INSERT INTO change_log VALUES(?,?,?,?,?,?,?,?,?,?)",
                      tuple(row.get(k, "") for k in ("timestamp", "user", "action", "target", "file", "location",
                                                     "field", "old_value", "new_value", "reason")))

    def backfill_log(self, df: pd.DataFrame) -> None:
        with self.conn() as c:
            if c.execute("SELECT COUNT(*) FROM change_log").fetchone()[0] == 0 and len(df):
                df.to_sql("change_log", c, if_exists="append", index=False)

    # ------------------------------------------------------------------ browsing
    def tables(self) -> pd.DataFrame:
        with self.conn() as c:
            names = c.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table','view') "
                              "AND name NOT LIKE 'sqlite_%' ORDER BY type, name").fetchall()
            out = []
            for n, t in names:
                cnt = c.execute(f'SELECT COUNT(*) FROM "{n}"').fetchone()[0]
                out.append({"Name": n, "Type": t, "Rows": cnt})
        return pd.DataFrame(out)

    def query(self, sql: str, limit: int = 5000) -> pd.DataFrame:
        """Read-only query (SELECT / WITH only)."""
        s = sql.strip().rstrip(";")
        if not re.match(r"(?is)^\s*(select|with)\b", s) or ";" in s:
            raise ValueError("Only a single SELECT (or WITH … SELECT) statement is allowed here.")
        uri = f"file:{self.path.as_posix()}?mode=ro"
        c = sqlite3.connect(uri, uri=True)
        try:
            return pd.read_sql_query(f"SELECT * FROM ({s}) LIMIT {int(limit)}", c)
        finally:
            c.close()


# ---------------------------------------------------------------------- generic reader
def read_any(name: str, data: bytes) -> List[tuple]:
    """Return [(sheet, DataFrame, header_row_excel)] for any xlsx/xlsm/xls/csv."""
    from .files import open_workbook
    import tempfile
    suffix = Path(name).suffix.lower()
    grids = []
    if suffix in (".csv", ".txt"):
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(data), header=None, dtype=object, encoding=enc, sep=None, engine="python")
                break
            except Exception:
                df = None
        if df is None:
            raise ValueError(f"{name}: could not read as CSV.")
        grids.append(("csv", df.values.tolist()))
    else:
        with tempfile.NamedTemporaryFile(suffix=suffix or ".xlsx", delete=False) as tmp:
            tmp.write(data)
            p = Path(tmp.name)
        try:
            wb = open_workbook(p, data_only=True)
            for ws in wb.worksheets:
                grids.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
        finally:
            p.unlink(missing_ok=True)
    out = []
    for sheet, grid in grids:
        grid = [r for r in grid]
        if not any(any(v not in (None, "") for v in r) for r in grid):
            continue
        hdr = _header_row(grid)
        width = max(len(r) for r in grid)
        heads, seen = [], {}
        for j in range(width):
            v = grid[hdr][j] if j < len(grid[hdr]) else None
            h = safe_ident(v, 40) if v not in (None, "") else f"col_{j + 1}"
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 1
            heads.append(h)
        rows = []
        for i in range(hdr + 1, len(grid)):
            r = list(grid[i]) + [None] * (width - len(grid[i]))
            if all(v in (None, "") for v in r):
                continue
            rows.append([i + 1] + [(v if isinstance(v, (int, float, str)) or v is None else str(v)) for v in r])
        df = pd.DataFrame(rows, columns=["source_row"] + heads)
        df = df.loc[:, [c for c in df.columns if c == "source_row" or df[c].notna().any()]]
        out.append((sheet, df, hdr + 1))
    return out


def _header_row(grid) -> int:
    best, score_best = 0, -1
    for i, r in enumerate(grid[:40]):
        texts = [v for v in r if isinstance(v, str) and v.strip() and not re.fullmatch(r"[-+()\d.,%\s]+", v)]
        nums = [v for v in r if isinstance(v, (int, float))]
        score = len(texts) - 1.5 * len(nums)
        if len(texts) >= 2 and score > score_best:
            best, score_best = i, score
    return best
