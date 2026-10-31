# Table-2 Builder: Actual to date from the CFR

A Streamlit app that reads the **Chart of Accounts** and the **CFR (Civil Accounts)** from a
`resources` folder and builds **Table 2 – Revenue by Object, "Actual to date"**. It uses the same
lines and descriptions as the `Table-2(2QTR)` sheet of the June 2025-26 workbook.

You can also do the following in the app:
- **edit** the CFR or the Chart of Accounts. Each save creates a **new updated file** in the same
  folder, and the app switches to that file.
- change the line mapping (which codes go into which line)
- keep a **change log** of every action

---

## ⚡ Quick start: just upload the two files

1. **Windows:** double-click `Run Table2 App (Windows).bat`. On the first run it installs Python
   packages by itself. **Mac/Linux:** run `./run.sh`.
2. The app opens in your browser. Drag **both files** into the box at the top:
   - the Chart of Accounts (e.g. `15 Dec 2024.xlsx`)
   - the CFR (e.g. `cfr.xlsx`)
3. The app then works on its own:
   - It **recognises** which file is which by looking inside it, not by its name.
   - It **stores** each file in `resources/charter/` or `resources/civil_accounts/`. If the same file
     is already there, it does not make a copy.
   - It **makes both files active** and **builds Table-2**.
   - It shows totals and checks, and a **⬇ Download Table-2 (Excel)** button.
   - It writes every step in `resources/logs/change_log.csv`.
4. Next month, upload just the new CFR. The new file becomes active and Table-2 is rebuilt.

All other options (units, comparison, editing the CFR or Chart of Accounts, mapping, files, log) are
in the tabs below the result.

## 🗄️ Database (automatic)

Everything you upload is also stored in **`resources/table2.db`** (SQLite – part of Python, nothing to install).

| What you upload | What happens |
|---|---|
| **CFR** (1–3 sheets, with a `Detail Item Desc` header) | Saved in `civil_accounts/`; every row goes to table **`cfr_rows`** |
| **Chart of Accounts** (A=…, B=…, C=… sheets) | Saved in `charter/`; every code goes to **`charter_codes`** |
| **Any other Excel / CSV** | Saved in `other_uploads/`; **every sheet becomes its own table** `raw_<file>_<sheet>` (header row detected automatically, plus a `source_row` column) |

Also stored:
- **`table2_runs`** and **`table2_lines`**: every distinct Table-2 result. A new run is added when the CFR, charter or mapping changes, so month-by-month history builds up by itself.
- **`files`**: every file ever loaded, with its SHA-256, and **`active_files`**.
- **`change_log`**: a copy of the CSV log.
- **Views:** `v_cfr_active`, `v_charter_active`, `v_table2_latest`.

Uploading the same file again does nothing, so data is never duplicated. Hierarchical IFMIS
reports (`*`, `**`, `***` levels) are read without their subtotal rows, so nothing is counted twice.

A big report workbook that merely *contains* a CFR sheet (for example the 19-sheet T01–T03
workbook) is stored as general data, not as the CFR. If it really is your CFR, a
**"Use … as the CFR"** button appears.

In the **🗄️ Database** tab you can:
- see every table with its row count
- preview a table, and download it as CSV or Excel
- run read-only SQL (ready-made examples are included)
- download the `.db` file

You can also open `table2.db` in *DB Browser for SQLite* or Power BI.

---

## 1. Manual setup

Python 3.10 or newer is required.

```bash
cd table2_app
python -m venv .venv
# Windows:  .venv\Scripts\activate
# Mac/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

The browser opens at http://localhost:8501.

## 2. Folder layout

```
table2_app/
├── app.py                         # Streamlit app
├── t2core/                        # engine (no Streamlit code, unit-tested)
│   ├── cfr.py                     # read CFR, apply edits, save updated file
│   ├── charter.py                 # read Chart of Accounts, add/edit/delete codes, save updated file
│   ├── table2.py                  # Table-2 calculation + reconciliation + charter check
│   ├── mapping.py                 # Table-2 lines (labels, structure, code patterns)
│   ├── default_mapping.json       # taken from the source workbook's Table-2 formulas
│   ├── files.py                   # folder tracking, active file, versioned names
│   ├── changelog.py               # append-only CSV log
│   └── export.py                  # formatted Excel download
├── resources/
│   ├── charter/                   # uploaded Chart of Accounts + updated versions
│   ├── civil_accounts/            # uploaded CFRs + updated versions
│   ├── mapping/                   # table2_mapping.json (+ history/)      ← created on first run
│   ├── logs/change_log.csv        # every change                         ← created on first run
│   ├── outputs/                   # saved Table-2 exports (optional)
│   └── .app_state.json            # active files and file fingerprints
└── tests/                         # tests + sample data in tests/data
```

To keep the resources folder somewhere else, type its path in the sidebar, or set the environment
variable `T2_RESOURCES`.

## 3. Tabs

| Tab | What it does |
|---|---|
| **📊 Table 2** | "Actual to date" for 66 lines (53 detail lines plus subtotals). Options: units (Rupees, Thousand, Million, Billion), decimals, hide zero lines, show code count and patterns, **compare with another CFR file**. Downloads: **Excel** and **CSV**, or save a copy in `resources/outputs`. |
| **🔎 Detail & checks** | Reconciliation (CFR total = mapped + excluded + unmapped), the Table-2 line for each code, unmapped codes with suggested lines, and the CFR vs Chart of Accounts check (missing codes, different wording, blocked codes). |
| **✏️ Edit CFR** | Edit an amount, code or description; add or delete a row. Saving with a **reason** creates `cfr (updated YYYY-MM-DD HHMM).xlsx` in `civil_accounts/`. The original file is never changed, and the new file becomes the active CFR. |
| **📘 Edit Charter** | Choose an object sheet (A, B, C, E, F, G, H). Edit a code's description or Old Code/CS, add a new code (it goes after the last code of the same minor head), or delete a code. Saving creates `15 Dec 2024 (updated …).xlsx` with all 15 sheets and makes it the active charter. |
| **🧭 Mapping** | Edit the code patterns of each line (`B027` means every B027xx code). You see the effect on Table-2 before saving. Previous versions are kept in `mapping/history/`. You can restore the default. |
| **📁 Files** | Every file in both folders, with the active one, modified time, the file it came from, the reason and its SHA-256. You can also upload a new file into a folder without overwriting anything. |
| **📝 Change log** | Every entry, with search and filter, and a CSV download. |

**Active file:** in the sidebar, choose "Automatic (newest file)" or a specific file. After each
save the app switches to the new file.

**Tracking:** each time the app runs it scans both folders. It writes a log entry when a file is
new, has been changed outside the app (its SHA-256 differs), or has been removed.

## 4. Calculation rules (same as the source workbook)

- A line's amount is the total of every CFR code that **starts with** one of that line's patterns.
  This gives the same result as the workbook's `SUMIF(…,"*"&code&"*",…)`.
- Duplicate codes in the CFR are separate IFMIS entries, so they are **added together**.
- Receipts are negative in the CFR, so **Actual = −(sum)**. The amount is then divided by the unit
  (million by default).
- **Excluded:** E012 (Recoveries of investment) and E032 (Cash credit). These are left out of
  Table-2 in the source workbook too, and come to −110,675.5 million.
- **Checks:**
  - Table-2 total = mapped codes
  - every rupee in the CFR is accounted for
  - no code is counted in two lines
  - no unmapped code has a value

  If a new code arrives that is in no line, it shows as a **warning**. Add it to a line in the
  Mapping tab.
- Results for **CFR March 2026** match all 66 lines of the source workbook's `Table-2(2QTR)`
  column E (TOTAL = **2,980,534.640** million). This is checked in `tests/test_table2.py`.

## 5. Change log columns

`timestamp, user, action, target (cfr/charter/mapping/table2), file, location, field, old_value, new_value, reason`

The file opens directly in Excel.

## 6. Tests

```bash
pytest -q
```

The tests cover:
- matching the source workbook
- reconciliation
- units and sign
- detecting double mapping
- CFR edit → new file (the original stays the same)
- charter add/edit/delete
- file tracking and the log
- the Excel export

## 7. Notes

- `.xls` files can be read if `xlrd` is installed. Updated files are always saved as `.xlsx`.
- Editing works on the **object-code sheets** of the Chart of Accounts (A to H). Fund, Function and
  Entity sheets are copied into the updated file unchanged.
- The source workbook writes "E031-FROM OTHERS" as the label of the line that adds up E02701
  (domestic loans). That label is kept here as it is. You can change it in the Mapping tab.
"# datas" 
