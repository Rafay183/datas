"""Table-2 'Actual to date' builder – Streamlit app.

Run:   streamlit run app.py
Data:  resources/charter/        Chart of Accounts workbook(s)
       resources/civil_accounts/ CFR workbook(s)
"""
from __future__ import annotations

import json
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from t2core import config
from t2core.changelog import ChangeLog
from t2core.charter import diff_codes, parse_charter, save_updated as save_charter, validate_additions
from t2core.cfr import CFRFormatError, diff_rows, parse_cfr, save_updated as save_cfr, validate_grid
from t2core.export import build_workbook
from t2core.files import (State, detect_changes, resolve_active, safe_copy_into, scan, sha256_of, versioned_name)
from t2core.intake import ingest
from t2core.db import DB, DB_NAME
from t2core import mapping as mp
from t2core.table2 import compare, compute

st.set_page_config(page_title="Table-2 Actual to Date", page_icon="📊", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1.4rem;max-width:1450px}
h1,h2,h3{color:#1F4E79}
div[data-testid="stMetric"]{background:#F3F6FA;border:1px solid #DCE3EC;border-radius:8px;padding:8px 12px}
</style>""", unsafe_allow_html=True)


# =========================================================================== cached loaders
@st.cache_data(show_spinner="Reading CFR…")
def load_cfr(path: str, mtime: float, size: int):
    return parse_cfr(Path(path))


@st.cache_data(show_spinner="Reading charter…")
def load_charter(path: str, mtime: float, size: int):
    return parse_charter(Path(path))


def stamp(f):
    return str(f.path), f.path.stat().st_mtime, f.path.stat().st_size


# =========================================================================== setup / sidebar
ss = st.session_state
ss.setdefault("user", "")
ss.setdefault("edit_nonce", 0)
ss.setdefault("processed", {})      # (name, sha) -> IntakeResult, so each upload is handled once

with st.sidebar:
    st.markdown("### 📊 Table-2 Builder")
    st.caption(f"Actual to date from CFR • v{config.APP_VERSION}")
    root = Path(st.text_input("Resources folder", str(config.DEFAULT_RESOURCES)))
    ss.user = st.text_input("Your name (for the change log)", ss.user)
    if st.button("🔄 Rescan folders", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    P = config.resource_paths(root)
except Exception as e:
    st.error(f"Cannot use resources folder '{root}': {e}")
    st.stop()
state = State(P["root"])
db = DB(P["root"] / DB_NAME)
log = ChangeLog(P["logs"] / config.LOG_FILE, db=db)
db.backfill_log(log.read())          # first run with an existing CSV log: copy it into the DB
user = ss.user or None

st.title("Table 2 – Revenue by Object: Actual to Date")
with st.container(border=True):
    st.markdown("### ⬆️ Upload Files")
    st.caption("Drop the Chart of Accounts and the CFR here. The app identifies each file, stores it in the right "
               "folder, makes it active, builds Table-2 and saves everything in the **database** (resources/table2.db). "
               "Any other Excel/CSV file is imported too – each sheet becomes a database table.")
    ups = st.file_uploader("Files", type=["xlsx", "xlsm", "xls", "csv", "txt"], accept_multiple_files=True,
                           key="quick_up", label_visibility="collapsed")
    import hashlib as _h
    for up in ups or []:
        data = up.getvalue()
        k = f"{up.name}|{_h.sha256(data).hexdigest()}"
        if k in ss.processed:
            continue
        r = ingest(up.name, data, P, state, log, user, db=db)
        ss.processed[k] = r
        if r.kind in ("cfr", "charter"):
            ss[f"pick_{r.kind}"] = r.stored_as        # keep the sidebar selector in sync
            st.cache_data.clear()
    shown = [r for k, r in ss.processed.items() if any(k.startswith(u.name + "|") for u in (ups or []))]
    if shown:
        icon = {"saved": "✅", "already in folder": "♻️", "imported to DB": "🗄️", "already in DB": "♻️", "error": "❌"}
        st.dataframe(pd.DataFrame([{"": icon.get(r.status, ""), "Uploaded file": r.name,
                                    "Detected as": {"cfr": "CFR (Civil Accounts)", "charter": "Chart of Accounts",
                                                    "other": "Other data → database"}.get(r.kind, "—"),
                                    "Result": r.status, "Stored as": r.stored_as, "Why": r.message} for r in shown]),
                     hide_index=True, use_container_width=True)
        for r in shown:
            if r.status == "error":
                st.error(f"{r.name}: {r.message}")
            if r.cfr_candidate:
                st.info(f"**{r.name}** contains a CFR-style sheet (*{r.cfr_candidate}*) but has many other sheets, "
                        "so it was saved as general data in the database. If it really is your CFR, press the button.")
                if st.button(f"Use {r.name} as the CFR", key=f"force_{r.name}"):
                    up = next(u for u in ups if u.name == r.name)
                    fr = ingest(up.name, up.getvalue(), P, state, log, user, db=db, force_kind="cfr")
                    ss.processed[f"{up.name}|forced"] = fr
                    ss["pick_cfr"] = fr.stored_as
                    st.cache_data.clear()
                    st.rerun()


def track(kind: str):
    files = scan(P[kind], kind, state)
    for ev in detect_changes(files, kind, state):
        if ev["event"] == "new":
            log.add("File detected", kind, ev["file"], field="sha256", new_value=ev["sha"][:16], user=user)
        elif ev["event"] == "changed":
            log.add("File changed outside the app", kind, ev["file"], field="sha256",
                    old_value=ev["old"][:16], new_value=ev["sha"][:16], user=user)
            st.toast(f"{ev['file']} was changed outside the app – reloaded.", icon="⚠️")
        else:
            log.add("File removed", kind, ev["file"], field="sha256", old_value=ev["old"][:16], user=user)
    return files


charter_files, cfr_files = track("charter"), track("cfr")
m, created = mp.load(P["mapping"])
if created:
    log.add("Mapping created from default", "mapping", config.MAPPING_FILE, user=user,
            reason="First run – reproduces Table-2(2QTR) column E of the source workbook")


for _k in ("cfr", "charter"):                      # a save in a tab asked to switch the active file
    if f"_pending_pick_{_k}" in ss:
        ss[f"pick_{_k}"] = ss.pop(f"_pending_pick_{_k}")


def picker(kind: str, files, label: str):
    names = ["Automatic (newest file)"] + [f.name for f in files]
    cur = state.get_active(kind)
    idx = names.index(cur) if cur in names else 0
    sel = st.selectbox(label, names, index=idx, key=f"pick_{kind}")
    want = "" if sel == names[0] else sel
    if want != cur:
        state.set_active(kind, want)
        log.add("Active file changed", kind, want or "(automatic)", old_value=cur or "(automatic)",
                new_value=want or "(automatic)", user=user)
    return resolve_active(files, kind, state)


with st.sidebar:
    st.divider()
    act_cfr = picker("cfr", cfr_files, "CFR file (civil_accounts)")
    act_ch = picker("charter", charter_files, "Charter file (charter)")
    if act_cfr:
        st.caption(f"Using **{act_cfr.name}**")
    if act_ch:
        st.caption(f"Using **{act_ch.name}**")

if not cfr_files:
    st.warning(f"No CFR workbook found in `{P['cfr']}`. Upload it in the box above.")
if not charter_files:
    st.info("No Chart of Accounts yet – Table-2 still works; upload it above to enable the charter check.")

cfr = charter = None
if act_cfr:
    try:
        cfr = load_cfr(*stamp(act_cfr))
    except (CFRFormatError, Exception) as e:
        st.error(f"Could not read {act_cfr.name}: {e}")
if act_ch:
    try:
        charter = load_charter(*stamp(act_ch))
    except Exception as e:
        st.error(f"Could not read {act_ch.name}: {e}")

# ---------------------------------------------------------------- keep the database in sync with the folders
def sync_db():
    for f in cfr_files:
        if db.file_id("cfr", f.sha256) is None:
            try:
                _, new = db.load_cfr(f, load_cfr(*stamp(f)))
                if new:
                    log.add("CFR loaded into database", "database", f.name, field="rows",
                            new_value=db.query(f"SELECT rows_loaded FROM files WHERE kind='cfr' AND sha256='{f.sha256}'").iat[0, 0],
                            user=user)
            except Exception as e:
                st.warning(f"{f.name} not loaded into the database: {e}")
    for f in charter_files:
        if db.file_id("charter", f.sha256) is None:
            try:
                _, new = db.load_charter(f, load_charter(*stamp(f)))
                if new:
                    log.add("Charter loaded into database", "database", f.name, user=user)
            except Exception as e:
                st.warning(f"{f.name} not loaded into the database: {e}")
    if act_cfr:
        db.set_active("cfr", act_cfr.sha256)
    if act_ch:
        db.set_active("charter", act_ch.sha256)


sync_db()

if cfr:
    st.caption(f"CFR **{cfr.file}** · sheet {cfr.sheet} · column '{cfr.amount_header}' · period **{cfr.period or '—'}** · "
               f"{len(cfr.rows):,} rows · mapping v{m.get('version', 1)}")

# ---------------------------------------------------------------- one-glance result
if cfr:
    try:
        quick = compute(cfr, m, charter, unit=1e6)
        qfails = [c for c in quick.checks if c["Result"] == "FAIL"]
        run_id = db.save_table2(act_cfr.sha256, act_ch.sha256 if act_ch else None, int(m.get("version", 1)),
                                cfr.period, quick)
        with st.container(border=True):
            st.markdown(f"### ✅ Table-2 Ready – {cfr.period or 'Actual to Date'}")
            qt = quick.table
            cols = st.columns(4)
            for col, (_, r) in zip(cols, qt[qt["Level"] <= 1].iterrows()):
                col.metric(r["Description"], f"{r['Actual to date']:,.3f}")
            st.caption("Rs. in million · " + (f"{len(quick.checks)}/{len(quick.checks)} checks pass" if not qfails
                                               else f"⚠️ {len(qfails)} check(s) need attention – see the Detail tab"))
            for w in quick.warnings:
                st.warning(w)
            qx = build_workbook(quick, unit_name="Million", decimals=3, period=cfr.period, cfr_file=cfr.file,
                                charter_file=charter.file if charter else "(none)", mapping_version=m.get("version", 1))
            qname = f"Table2_Actual_{(cfr.period or 'to_date').replace(' ', '_')}.xlsx"
            if st.download_button("⬇ Download Table-2 (Excel)", qx, file_name=qname, type="primary",
                                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  key="quick_dl"):
                log.add("Table-2 downloaded (Excel)", "table2", qname, location=cfr.file, user=user)
            st.caption("More options (units, decimals, comparison, editing, database) are in the tabs below.")
    except Exception as e:
        st.error(f"Table-2 could not be built: {e}")

tabs = st.tabs(["📊 Table 2", "🔎 Detail & Checks", "✏️ Edit CFR", "📘 Edit Charter", "🧭 Mapping", "📁 Files", "📝 Change Log", "🗄️ Database"])


# =========================================================================== Table 2
with tabs[0]:
    if not cfr:
        st.info("Load a CFR file to build Table 2.")
    else:
        mapping_errors = mp.validate(m)
        if mapping_errors:
            st.error("Mapping problems: " + " | ".join(mapping_errors))
        c1, c2, c3, c4 = st.columns(4)
        unit_name = c1.selectbox("Units", list(config.UNITS), index=2)
        decimals = c2.number_input("Decimals", 0, 6, 3)
        hide_zero = c3.checkbox("Hide zero detail lines", False)
        show_extra = c4.checkbox("Show code count & patterns", False)
        cmp_opts = ["(none)"] + [f.name for f in cfr_files if f.name != cfr.file]
        cmp_name = st.selectbox("Compare with another CFR file (optional)", cmp_opts)
        try:
            res = compute(cfr, m, charter, unit=config.UNITS[unit_name])
        except Exception as e:
            st.error(f"Calculation failed: {e}")
            st.code(traceback.format_exc())
            st.stop()
        fails = [c for c in res.checks if c["Result"] == "FAIL"]
        top = res.table.iloc[0]
        sec = res.table[res.table["Level"] == 1]
        k = st.columns(2 + len(sec))
        k[0].metric(top["Description"], f"{top['Actual to date']:,.{decimals}f}")
        for col, (_, r) in zip(k[1:], sec.iterrows()):
            col.metric(r["Description"], f"{r['Actual to date']:,.{decimals}f}")
        k[-1].metric("Checks", f"{len(res.checks) - len(fails)}/{len(res.checks)} pass")
        for w in res.warnings:
            st.warning(w)

        cmp_df = None
        if cmp_name != "(none)":
            other = next(f for f in cfr_files if f.name == cmp_name)
            try:
                res_b = compute(load_cfr(*stamp(other)), m, charter, unit=config.UNITS[unit_name])
                cmp_df = compare(res, res_b, cfr.file, other.name)
            except Exception as e:
                st.error(f"Comparison file could not be read: {e}")

        view = res.table.copy()
        if hide_zero:
            view = view[(view["Type"] == "total") | (view["Actual to date"].round(decimals) != 0)]
        view["Description"] = view.apply(lambda r: " " * int(r["Level"]) + r["Description"], axis=1)
        view["Code"] = view.apply(lambda r: r["Code"] if r["Type"] == "detail" else "", axis=1)
        cols = ["Code", "Description", "Actual to date"] + (["CFR codes", "Patterns"] if show_extra else [])
        shown = view[cols].rename(columns={"Actual to date": f"Actual to date (Rs {unit_name.lower()})"})
        if cmp_df is not None:
            c = cmp_df.set_index("Line Id").loc[view["Line Id"]]
            shown[cmp_name] = c[cmp_name].values
            shown["Difference"] = c["Difference"].values
        bold = view["Type"].eq("total").values

        shown = shown.reset_index(drop=True)
        shown.index.name = None
        num_cols = [c for c in shown.columns if c not in ("Code", "Description", "Patterns", "CFR codes")]
        sty = shown.style.apply(lambda r: ["font-weight:700;background-color:#DDEBF7" if bold[r.name] else "" for _ in r],
                                axis=1).format({c: f"{{:,.{decimals}f}}" for c in num_cols}, na_rep="")
        st.dataframe(sty, hide_index=True, use_container_width=True, height=min(38 * len(shown) + 40, 900))

        st.markdown("**Checks**")
        st.dataframe(pd.DataFrame(res.checks), hide_index=True, use_container_width=True)

        sources = pd.DataFrame([{"Role": f.kind, "File": f.name, "Modified": f.modified.strftime("%Y-%m-%d %H:%M"),
                                 "SHA-256": f.sha256} for f in [act_cfr, act_ch] if f])
        xlsx = build_workbook(res, unit_name=unit_name, decimals=int(decimals), period=cfr.period, cfr_file=cfr.file,
                              charter_file=charter.file if charter else "(none)", mapping_version=m.get("version", 1),
                              hide_zero=hide_zero, compare_df=cmp_df, sources=sources)
        fname = f"Table2_Actual_{(cfr.period or 'to_date').replace(' ', '_')}_{datetime.now():%Y%m%d_%H%M}.xlsx"
        d1, d2, d3 = st.columns(3)
        if d1.download_button("⬇ Download Excel", xlsx, file_name=fname, type="primary", use_container_width=True,
                              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
            log.add("Table-2 downloaded (Excel)", "table2", fname, location=cfr.file, user=user)
        csv = res.table[["Code", "Description", "Actual to date"]].to_csv(index=False).encode("utf-8-sig")
        if d2.download_button("⬇ Download CSV", csv, file_name=fname.replace(".xlsx", ".csv"), use_container_width=True):
            log.add("Table-2 downloaded (CSV)", "table2", fname.replace(".xlsx", ".csv"), location=cfr.file, user=user)
        if d3.button("💾 Save a copy in resources/outputs", use_container_width=True):
            out = P["outputs"] / fname
            out.write_bytes(xlsx)
            log.add("Table-2 saved to outputs", "table2", out.name, location=cfr.file, user=user)
            st.success(f"Saved {out}")


# =========================================================================== Detail & checks
with tabs[1]:
    if not cfr:
        st.info("Load a CFR file first.")
    else:
        res = compute(cfr, m, charter)
        st.markdown("**Reconciliation (Rs)**")
        st.dataframe(res.recon.drop(columns=["Signed ÷ unit"]).style.format({"Rs": "{:,.0f}"}),
                     hide_index=True, use_container_width=True)
        a, b = st.columns([1, 2])
        stat = a.multiselect("Status", sorted(res.detail["Status"].unique()), default=sorted(res.detail["Status"].unique()))
        q = b.text_input("Search code / description / line")
        d = res.detail[res.detail["Status"].isin(stat)]
        if q:
            d = d[d.apply(lambda r: q.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
        st.dataframe(d.drop(columns=["Line Id"]).style.format({"Amount (Rs)": "{:,.0f}"}), hide_index=True,
                     use_container_width=True, height=420)
        if not res.unmapped.empty:
            st.markdown("**Unmapped Codes** – add them to a line in the *Mapping* tab")
            st.dataframe(res.unmapped.drop(columns=["Line Id", "Line"]), hide_index=True, use_container_width=True)
        if not res.charter_check.empty:
            st.markdown("**CFR Codes vs Charter**")
            ch_f = st.multiselect("Show", sorted(res.charter_check["Check"].unique()),
                                  default=[x for x in res.charter_check["Check"].unique() if x not in ("Same", "Same (CFR text shortened)")])
            st.dataframe(res.charter_check[res.charter_check["Check"].isin(ch_f)].style.format({"Amount (Rs)": "{:,.0f}"}),
                         hide_index=True, use_container_width=True, height=360)
        if cfr.notes or len(cfr.skipped):
            with st.expander("CFR reading notes"):
                for n in cfr.notes:
                    st.write("•", n)
                if len(cfr.skipped):
                    st.dataframe(cfr.skipped, hide_index=True)


# =========================================================================== Edit CFR
with tabs[2]:
    if not cfr or not act_cfr:
        st.info("Load a CFR file first.")
    else:
        st.markdown(f"Editing **{cfr.file}**. Saving creates a **new file** in `civil_accounts/` "
                    "(the original is never overwritten) and makes it the active CFR.")
        q = st.text_input("Filter rows (code or description)", key="cfr_q")
        base = cfr.rows[["Row", "Code", "Description", "Amount"]].copy()
        if q:
            base = base[base["Code"].str.contains(q, case=False) | base["Description"].str.contains(q, case=False)]
        st.caption(f"{len(base):,} rows shown. Add rows at the bottom; delete with the row checkbox + trash icon. "
                   "With a filter on, only the visible rows can be changed or deleted.")
        grid = st.data_editor(base, num_rows="dynamic", use_container_width=True, height=440, hide_index=True,
                              key=f"cfr_ed_{cfr.file}_{ss.edit_nonce}_{q}",
                              column_config={"Row": st.column_config.NumberColumn("Excel row", disabled=True),
                                             "Amount": st.column_config.NumberColumn(format="%.0f")})
        changes = diff_rows(base, grid)
        errs = validate_grid(grid) if changes else []
        if changes:
            st.markdown(f"**{len(changes)} Pending Change(s)**")
            st.dataframe(pd.DataFrame(changes).astype(str), hide_index=True, use_container_width=True)
        for e in errs:
            st.error(e)
        reason = st.text_input("Reason for the change (required)", key="cfr_reason")
        if st.button("💾 Save as updated CFR file", type="primary", disabled=not changes or bool(errs) or not reason.strip()):
            full = cfr.rows[["Row", "Code", "Description", "Amount"]].copy()
            if q:  # merge filtered edits back into the full list
                shown_rows = set(base["Row"])
                rest = full[~full["Row"].isin(shown_rows)]
                full = pd.concat([rest, grid], ignore_index=True)
                full["_o"] = full["Row"].fillna(10 ** 9)
                full = full.sort_values("_o", kind="stable").drop(columns="_o")
            else:
                full = grid
            target = versioned_name(P["cfr"], cfr.file)
            try:
                save_cfr(act_cfr.path, cfr, full, target)
            except Exception as e:
                st.error(f"Save failed: {e}")
                st.stop()
            state.add_lineage(target.name, cfr.file, reason)
            state.set_active("cfr", target.name)
            ss["_pending_pick_cfr"] = target.name
            log.add("CFR updated file created", "cfr", target.name, location=f"from {cfr.file}",
                    field="changes", new_value=len(changes), reason=reason, user=user)
            log.add_many(changes, target="cfr", file=target.name, reason=reason, user=user)
            ss.edit_nonce += 1
            st.cache_data.clear()
            st.success(f"Saved {target.name} – it is now the active CFR.")
            st.rerun()


# =========================================================================== Edit charter
with tabs[3]:
    if not charter or not act_ch:
        st.info("Load a charter file first.")
    else:
        st.markdown(f"Editing **{charter.file}** (object-code sheets). Saving creates a **new workbook** in `charter/` "
                    "with all 15 sheets; only the edited sheet changes.")
        if charter.issues:
            with st.expander(f"⚠️ {len(charter.issues)} issues found in this charter"):
                for i in charter.issues:
                    st.write("•", i)
        c1, c2 = st.columns([1, 2])
        sheet = c1.selectbox("Sheet", charter.sheets, index=charter.sheets.index("C=Non Tax Revenue") if "C=Non Tax Revenue" in charter.sheets else 0)
        q = c2.text_input("Filter (code, minor or description)", key="ch_q")
        base = charter.codes[charter.codes["Sheet"] == sheet][["Row", "Major", "Minor", "Code", "Description", "Old Code / CS"]]
        if q:
            base = base[base.apply(lambda r: q.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
        st.caption(f"{len(base):,} codes shown. For a new code fill Major, Minor, Code and Description – it is inserted "
                   "after the last code of the same minor head.")
        grid = st.data_editor(base, num_rows="dynamic", use_container_width=True, height=440, hide_index=True,
                              key=f"ch_ed_{charter.file}_{sheet}_{ss.edit_nonce}_{q}",
                              column_config={"Row": st.column_config.NumberColumn("Excel row", disabled=True)})
        changes = diff_codes(base, grid, sheet)
        # existing rows: Major/Minor are informational (not editable through the grid)
        errs = validate_additions(changes, charter) if changes else []
        if changes:
            st.markdown(f"**{len(changes)} Pending Change(s)**")
            st.dataframe(pd.DataFrame([{k: v for k, v in c.items() if k not in ("rec",)} for c in changes]).astype(str),
                         hide_index=True, use_container_width=True)
        for e in errs:
            st.error(e)
        reason = st.text_input("Reason / reference (e.g. CS No.) – required", key="ch_reason")
        if st.button("💾 Save as updated charter file", type="primary", disabled=not changes or bool(errs) or not reason.strip()):
            target = versioned_name(P["charter"], charter.file)
            try:
                save_charter(act_ch.path, charter, sheet, changes, target)
            except Exception as e:
                st.error(f"Save failed: {e}")
                st.code(traceback.format_exc())
                st.stop()
            state.add_lineage(target.name, charter.file, reason)
            state.set_active("charter", target.name)
            ss["_pending_pick_charter"] = target.name
            log.add("Charter updated file created", "charter", target.name, location=f"from {charter.file}",
                    field="changes", new_value=len(changes), reason=reason, user=user)
            log.add_many([{k: v for k, v in c.items() if k in ("action", "location", "field", "old_value", "new_value")}
                          for c in changes], target="charter", file=target.name, reason=reason, user=user)
            ss.edit_nonce += 1
            st.cache_data.clear()
            st.success(f"Saved {target.name} – it is now the active charter.")
            st.rerun()


# =========================================================================== Mapping
with tabs[4]:
    st.markdown("Each Table-2 detail line sums the CFR codes that **start with** one of its patterns "
                "(e.g. `B027` = every B027xx code, `C03604` = that code only). Totals are built from their child lines.")
    det = pd.DataFrame([{"Line Id": l["id"], "Code": l.get("code", ""), "Description": l["label"],
                         "Patterns": ", ".join(l.get("patterns", [])), "Note": l.get("note", "")}
                        for l in m["lines"] if l["type"] == "detail"])
    ed = st.data_editor(det, hide_index=True, use_container_width=True, height=460, num_rows="fixed",
                        disabled=["Line Id", "Code"], key=f"map_ed_{m.get('version', 1)}_{ss.edit_nonce}")
    exc = pd.DataFrame(m.get("exclusions", []), columns=["pattern", "reason"])
    st.markdown("**Exclusions** – codes deliberately left out of Table-2")
    exc_ed = st.data_editor(exc, hide_index=True, num_rows="dynamic", use_container_width=True,
                            key=f"exc_ed_{m.get('version', 1)}_{ss.edit_nonce}")
    newm = json.loads(json.dumps(m))
    changes = []
    byid = {l["id"]: l for l in newm["lines"]}
    for _, r in ed.iterrows():
        l = byid[r["Line Id"]]
        new_p = mp.parse_patterns(r["Patterns"])
        if new_p != l.get("patterns", []):
            changes.append({"action": "Mapping patterns changed", "location": f"{l['id']} {l['label']}",
                            "field": "patterns", "old_value": ", ".join(l.get("patterns", [])), "new_value": ", ".join(new_p)})
            l["patterns"] = new_p
        if str(r["Description"]).strip() != l["label"]:
            changes.append({"action": "Mapping label changed", "location": l["id"], "field": "label",
                            "old_value": l["label"], "new_value": str(r["Description"]).strip()})
            l["label"] = str(r["Description"]).strip()
        if str(r["Note"] or "") != l.get("note", ""):
            l["note"] = str(r["Note"] or "")
    new_exc = [{"pattern": str(p).strip().upper(), "reason": str(rs or "")} for p, rs in zip(exc_ed["pattern"], exc_ed["reason"])
               if str(p or "").strip() and str(p) != "nan"]
    if new_exc != m.get("exclusions", []):
        changes.append({"action": "Mapping exclusions changed", "location": "exclusions", "field": "exclusions",
                        "old_value": json.dumps(m.get("exclusions", [])), "new_value": json.dumps(new_exc)})
        newm["exclusions"] = new_exc
    errs = mp.validate(newm)
    if changes:
        st.dataframe(pd.DataFrame(changes).astype(str), hide_index=True, use_container_width=True)
        if cfr:
            before = compute(cfr, m).table.set_index("Line Id")["Actual to date"]
            after = compute(cfr, newm).table.set_index("Line Id")["Actual to date"]
            diff = (after - before)
            moved = diff[diff.abs() > 1e-9]
            if len(moved):
                st.markdown("**Effect on Table-2 (Rs million)**")
                st.dataframe(pd.DataFrame({"Line": [byid[i]["label"] for i in moved.index], "Before": before[moved.index].values,
                                           "After": after[moved.index].values, "Change": moved.values}),
                             hide_index=True, use_container_width=True)
    for e in errs:
        st.error(e)
    reason = st.text_input("Reason for the mapping change (required)", key="map_reason")
    c1, c2 = st.columns(2)
    if c1.button("💾 Save mapping", type="primary", disabled=not changes or bool(errs) or not reason.strip()):
        p = mp.save(P["mapping"], newm)
        log.add_many(changes, target="mapping", file=p.name, reason=reason, user=user)
        ss.edit_nonce += 1
        st.success("Mapping saved (previous version kept in mapping/history).")
        st.rerun()
    if c2.button("↩ Restore default mapping (from the source workbook)"):
        p = mp.save(P["mapping"], {**mp.load_default(), "version": m.get("version", 1)})
        log.add("Mapping restored to default", "mapping", p.name, reason="user request", user=user)
        ss.edit_nonce += 1
        st.rerun()


# =========================================================================== Files
with tabs[5]:
    for kind, files, label in (("cfr", cfr_files, "Civil accounts (CFR)"), ("charter", charter_files, "Charter")):
        st.subheader(label)
        active = resolve_active(files, kind, state)
        rows = [{"Active": "✅" if active and f.name == active.name else "", "File": f.name,
                 "Modified": f.modified.strftime("%Y-%m-%d %H:%M"), "Size (KB)": round(f.size / 1024, 1),
                 "Derived from": f.parent, "Reason": f.note, "SHA-256": f.sha256[:16] + "…"} for f in files]
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        up = st.file_uploader(f"Add a {label} workbook to `{P[kind].name}/`", type=["xlsx", "xlsm", "xls"], key=f"up_{kind}")
        if up is not None and st.button(f"Add {up.name}", key=f"add_{kind}"):
            dst = safe_copy_into(P[kind], up.name, up.getvalue())
            log.add("File uploaded", kind, dst.name, field="size", new_value=len(up.getvalue()), user=user)
            st.success(f"Added {dst.name}")
            st.rerun()
    st.caption(f"Folder: `{P['root']}`")


# =========================================================================== Change log
with tabs[6]:
    lg = log.read()
    st.caption(f"`{log.path}` – {len(lg):,} entries (append-only)")
    if len(lg):
        c1, c2 = st.columns(2)
        tg = c1.multiselect("Target", sorted(lg["target"].unique()), default=sorted(lg["target"].unique()))
        q = c2.text_input("Search", key="log_q")
        v = lg[lg["target"].isin(tg)]
        if q:
            v = v[v.apply(lambda r: q.lower() in " ".join(r.values).lower(), axis=1)]
        st.dataframe(v.iloc[::-1], hide_index=True, use_container_width=True, height=480)
        st.download_button("⬇ Download change log (CSV)", log.path.read_bytes(), file_name="change_log.csv")


# =========================================================================== Database
with tabs[7]:
    st.markdown(f"SQLite database: `{db.path}` – updated automatically on every upload, edit and Table-2 build. "
                "Open it with *DB Browser for SQLite*, Power BI / Excel (ODBC) or Python.")
    tb = db.tables()
    c1, c2 = st.columns([1, 2])
    with c1:
        st.dataframe(tb, hide_index=True, use_container_width=True, height=420)
        st.download_button("⬇ Download database (.db)", db.path.read_bytes(), file_name="table2.db",
                           use_container_width=True)
    with c2:
        pick = st.selectbox("Preview a table / view", tb["Name"].tolist() if len(tb) else [])
        if pick:
            prev = db.query(f'SELECT * FROM "{pick}"', limit=1000)
            st.caption(f"First {len(prev):,} rows of {pick}")
            st.dataframe(prev, hide_index=True, use_container_width=True, height=300)
            x1, x2 = st.columns(2)
            full = db.query(f'SELECT * FROM "{pick}"', limit=1_000_000)
            x1.download_button(f"⬇ {pick}.csv", full.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"{pick}.csv", use_container_width=True)
            import io as _io
            bio = _io.BytesIO()
            full.to_excel(bio, index=False, sheet_name=pick[:31])
            x2.download_button(f"⬇ {pick}.xlsx", bio.getvalue(), file_name=f"{pick}.xlsx", use_container_width=True)
    st.markdown("**Run a Query (Read-Only)**")
    examples = {
        "Active CFR – total by major head": "SELECT major, COUNT(*) AS entries, SUM(amount) AS amount_rs FROM v_cfr_active GROUP BY major ORDER BY major",
        "Latest Table-2": "SELECT line_order, code, description, actual_million FROM v_table2_latest ORDER BY line_order",
        "Table-2 total of every run": "SELECT r.run_id, r.period, f.name AS cfr_file, r.mapping_version, r.total_rs/1e6 AS total_million, r.checks_passed||'/'||r.checks_total AS checks FROM table2_runs r JOIN files f ON f.file_id=r.cfr_file_id ORDER BY r.run_id",
        "CFR codes missing from the active charter": "SELECT DISTINCT c.code, c.description FROM v_cfr_active c LEFT JOIN v_charter_active h ON h.code=c.code WHERE h.code IS NULL ORDER BY c.code",
        "Compare two CFR files code by code": "SELECT code, SUM(CASE WHEN file_id=(SELECT MIN(file_id) FROM files WHERE kind='cfr') THEN amount END) AS first_file, SUM(CASE WHEN file_id=(SELECT MAX(file_id) FROM files WHERE kind='cfr') THEN amount END) AS latest_file FROM cfr_rows GROUP BY code ORDER BY code",
        "All files in the database": "SELECT file_id, kind, name, period, rows_loaded, loaded_at FROM files ORDER BY file_id",
        "Tables made from other uploads": "SELECT * FROM imported_tables ORDER BY created_at",
    }
    ex = st.selectbox("Example", list(examples))
    sql = st.text_area("SQL", examples[ex], height=110, key=f"sql_{ex}")
    if st.button("▶ Run query"):
        try:
            out = db.query(sql)
            st.caption(f"{len(out):,} rows")
            st.dataframe(out, hide_index=True, use_container_width=True)
            st.download_button("⬇ Result as CSV", out.to_csv(index=False).encode("utf-8-sig"), file_name="query.csv")
        except Exception as e:
            st.error(f"Query failed: {e}")
