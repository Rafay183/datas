"""Run:  pytest -q   (from the project folder)"""
import json
import shutil
from pathlib import Path

import pandas as pd
import pytest

from t2core.cfr import diff_rows, parse_cfr, save_updated as save_cfr, validate_grid
from t2core.changelog import ChangeLog
from t2core.charter import diff_codes, parse_charter, save_updated as save_charter, validate_additions
from t2core.config import resource_paths
from t2core.export import build_workbook
from t2core.files import State, detect_changes, resolve_active, scan, versioned_name
from t2core.mapping import load_default, validate
from t2core.table2 import compute

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(__file__).parent / "data"
EXPECTED = json.loads((Path(__file__).parent / "expected_table2_june2026.json").read_text())


@pytest.fixture
def res_dir(tmp_path):
    dst = tmp_path / "resources"
    P = resource_paths(dst)
    shutil.copy(DATA / "15 Dec 2024.xlsx", P["charter"])
    shutil.copy(DATA / "cfr.xlsx", P["cfr"])
    return P


@pytest.fixture
def cfr(res_dir):
    return parse_cfr(res_dir["cfr"] / "cfr.xlsx")


@pytest.fixture
def charter(res_dir):
    return parse_charter(res_dir["charter"] / "15 Dec 2024.xlsx")


def test_default_mapping_is_valid():
    assert validate(load_default()) == []


def test_cfr_parsing(cfr):
    assert (cfr.header_row, cfr.code_col, cfr.amount_col) == (6, 2, 4)
    assert cfr.period == "MARCH 2026"
    assert len(cfr.rows) == 518
    assert cfr.total == -3_091_210_164_753


def test_table2_reproduces_source_workbook(cfr, charter):
    """Every line equals column E of Table-2(2QTR) in the source workbook."""
    res = compute(cfr, load_default(), charter)
    got = dict(zip(res.table["Line Id"], res.table["Actual to date"]))
    for line, value in EXPECTED.items():
        assert got[line] == pytest.approx(value, abs=1e-6), line
    assert all(c["Result"] == "PASS" for c in res.checks)


def test_reconciliation_accounts_for_every_rupee(cfr):
    res = compute(cfr, load_default())
    r = res.recon.set_index("Item")["Rs"]
    assert r["Mapped to Table-2 lines"] + r["Excluded by rule"] + r["Unmapped (not in any line)"] == pytest.approx(cfr.total)
    assert r["Excluded by rule"] == pytest.approx(-110_675_524_743)   # E012x + E032x, as in the source


def test_unit_and_sign_options(cfr):
    res = compute(cfr, load_default(), unit=1e9, receipts_positive=False)
    assert res.table.iloc[0]["Actual to date"] == pytest.approx(-2980.53464001)


def test_double_mapping_is_detected(cfr):
    m = load_default()
    next(l for l in m["lines"] if l["id"] == "r67")["patterns"].append("C03908")  # also in r17
    res = compute(cfr, m)
    assert (res.detail["Status"] == "Double-mapped").sum() == 1
    assert any(c["Result"] == "FAIL" for c in res.checks)


def test_cfr_edit_creates_new_file(res_dir, cfr):
    g = cfr.rows[["Row", "Code", "Description", "Amount"]].copy()
    g.loc[g.index[g.Code == "B02384"][0], "Amount"] = -16_000_000_000
    g = pd.concat([g.iloc[:-1], pd.DataFrame([{"Row": None, "Code": "C03923", "Description": "NEW", "Amount": -5}])])
    assert validate_grid(g) == []
    ch = diff_rows(cfr.rows, g)
    assert {c["action"] for c in ch} == {"CFR value edited", "CFR row added", "CFR row deleted"}
    tgt = versioned_name(res_dir["cfr"], cfr.file)
    save_cfr(res_dir["cfr"] / cfr.file, cfr, g, tgt)
    new = parse_cfr(tgt)
    assert (res_dir["cfr"] / "cfr.xlsx").exists()               # original untouched
    assert parse_cfr(res_dir["cfr"] / "cfr.xlsx").total == cfr.total
    assert new.total == pytest.approx(cfr.total + (-16_000_000_000 + 15_986_892_133) + 50_118 - 5)
    res = compute(new, load_default())
    assert "C03923" in set(res.unmapped["Code"])


def test_invalid_cfr_code_rejected():
    g = pd.DataFrame([{"Row": None, "Code": "XYZ", "Description": "bad", "Amount": "abc"}])
    assert len(validate_grid(g)) == 2


def test_charter_edit_add_delete(res_dir, charter):
    sheet = "C=Non Tax Revenue"
    base = charter.codes[charter.codes.Sheet == sheet][["Row", "Major", "Minor", "Code", "Description", "Old Code / CS"]]
    g = base.copy()
    g.loc[g.index[g.Code == "C03123"][0], "Description"] = "Receipts from Plant Protection Operations"
    g = g.drop(g.index[g.Code == "C013aa"][0])
    g = pd.concat([g, pd.DataFrame([{"Row": None, "Major": "C03", "Minor": "C039", "Code": "C03922",
                                     "Description": "Net Proceeds of Wellhead Value Receipts", "Old Code / CS": ""}])])
    ch = diff_codes(base, g, sheet)
    assert sorted(c["op"] for c in ch) == ["add", "delete", "edit"]
    assert validate_additions(ch, charter) == []
    tgt = versioned_name(res_dir["charter"], charter.file)
    save_charter(res_dir["charter"] / charter.file, charter, sheet, ch, tgt)
    new = parse_charter(tgt).lookup()
    assert new["C03123"]["Description"] == "Receipts from Plant Protection Operations"
    assert "C013AA" not in new and new["C03922"]["Minor"] == "C039"
    assert len(parse_charter(res_dir["charter"] / charter.file).codes) == len(charter.codes)  # original untouched


def test_charter_rejects_bad_addition(charter):
    ch = [{"op": "add", "rec": {"Code": "C03101", "Minor": "C031", "Description": "dup"}}]
    assert any("already exists" in e for e in validate_additions(ch, charter))


def test_file_tracking_and_log(res_dir, tmp_path):
    st = State(res_dir["root"])
    files = scan(res_dir["cfr"], "cfr", st)
    assert [e["event"] for e in detect_changes(files, "cfr", st)] == ["new"]
    assert detect_changes(scan(res_dir["cfr"], "cfr", st), "cfr", st) == []
    (res_dir["cfr"] / "cfr.xlsx").write_bytes((res_dir["cfr"] / "cfr.xlsx").read_bytes() + b"")
    st.set_active("cfr", "cfr.xlsx")
    assert resolve_active(files, "cfr", st).name == "cfr.xlsx"
    log = ChangeLog(res_dir["logs"] / "change_log.csv")
    log.add("Test", "cfr", "cfr.xlsx", old_value=1, new_value=2, reason="r", user="tester")
    assert log.read().iloc[-1]["user"] == "tester"


def test_export(cfr, charter):
    from openpyxl import load_workbook
    import io
    x = build_workbook(compute(cfr, load_default(), charter), unit_name="Million", decimals=3, period=cfr.period,
                       cfr_file=cfr.file, charter_file=charter.file, mapping_version=1)
    wb = load_workbook(io.BytesIO(x))
    assert wb.sheetnames[0] == "Table 2"
    assert wb["Table 2"]["C5"].value == pytest.approx(2_980_534.64001)


def test_one_step_upload(tmp_path):
    from t2core.intake import ingest
    P = resource_paths(tmp_path / "resources")
    st, log = State(P["root"]), ChangeLog(P["logs"] / "change_log.csv")
    r1 = ingest("cfr.xlsx", (DATA / "cfr.xlsx").read_bytes(), P, st, log)
    r2 = ingest("15 Dec 2024.xlsx", (DATA / "15 Dec 2024.xlsx").read_bytes(), P, st, log)
    assert (r1.kind, r1.status, r2.kind, r2.status) == ("cfr", "saved", "charter", "saved")
    again = ingest("copy of cfr.xlsx", (DATA / "cfr.xlsx").read_bytes(), P, st, log)
    assert again.status == "already in folder" and len(list(P["cfr"].glob("*.xlsx"))) == 1
    assert st.get_active("cfr") == "cfr.xlsx" and st.get_active("charter") == "15 Dec 2024.xlsx"


def test_database_loads_everything_once(tmp_path, res_dir):
    from t2core.db import DB
    from t2core.intake import ingest
    P = res_dir
    st, db = State(P["root"]), DB(P["root"] / "table2.db")
    log = ChangeLog(P["logs"] / "change_log.csv", db=db)
    files = scan(P["cfr"], "cfr", st)
    cfr = parse_cfr(files[0].path)
    fid, new = db.load_cfr(files[0], cfr)
    assert new and db.load_cfr(files[0], cfr) == (fid, False)          # same file twice = no duplicate
    ch_file = scan(P["charter"], "charter", st)[0]
    db.load_charter(ch_file, parse_charter(ch_file.path))
    db.set_active("cfr", files[0].sha256)
    db.set_active("charter", ch_file.sha256)
    res = compute(cfr, load_default())
    rid = db.save_table2(files[0].sha256, ch_file.sha256, 1, cfr.period, res)
    assert db.save_table2(files[0].sha256, ch_file.sha256, 1, cfr.period, res) == rid
    q = db.query("SELECT COUNT(*) n, SUM(amount) s FROM v_cfr_active")
    assert q.iat[0, 0] == 518 and q.iat[0, 1] == pytest.approx(-3_091_210_164_753)
    assert db.query("SELECT actual_million FROM v_table2_latest WHERE line_id='r5'").iat[0, 0] == pytest.approx(2_980_534.64001)
    # any other file becomes tables
    csv_bytes = b"District,Receipts\nKarachi,10\nSukkur,5\n"
    r = ingest("districts.csv", csv_bytes, P, st, log, db=db)
    assert r.kind == "other" and r.status == "imported to DB"
    assert db.query("SELECT SUM(receipts) s FROM raw_districts_csv").iat[0, 0] == 15
    assert ingest("districts.csv", csv_bytes, P, st, log, db=db).status == "already in DB"
    log.add("Test entry", "test")
    assert db.query("SELECT COUNT(*) FROM change_log WHERE action='Test entry'").iat[0, 0] == 1
    with pytest.raises(ValueError):
        db.query("DROP TABLE files")


def test_big_report_workbook_is_not_taken_as_cfr():
    from openpyxl import Workbook
    import io
    from t2core.intake import classify
    wb = Workbook()
    wb.active["B6"] = "Detail Item Desc"
    for i in range(5):
        wb.create_sheet(f"Report {i}")
    bio = io.BytesIO(); wb.save(bio)
    kind, why = classify("report.xlsx", bio.getvalue())
    assert kind is None and why.startswith("CANDIDATE:")


def test_hierarchical_cfr_skips_subtotals(tmp_path):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(["Detail Item Desc", "Prog Jun Suppl"])
    ws.append(["*******", -30])
    ws.append(["**     Minor Item Desc B011-TAXES ON INCOME", -30])
    ws.append(["*      B01108-SHARE OF NET PROCEEDS", -20])
    ws.append(["*      B01176-AGRI TAX", -10])
    p = tmp_path / "h.xlsx"; wb.save(p)
    c = parse_cfr(p)
    assert len(c.rows) == 2 and c.total == -30 and len(c.skipped) == 2
