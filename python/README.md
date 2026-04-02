# Python Implementation

Two scripts + optional Streamlit web UI.
Status: **Secondary implementation — parallel to the Excel primary.**

---

## Files

| File | Version | Purpose |
|---|---|---|
| `run.py` | v1.3.0 | Invoice checker — main pipeline |
| `ppi_dashboard.py` | v1.3.0 | PPI Quota Dashboard — standalone, do not import from run.py |
| `app.py` | v1.4.0 | Streamlit web UI — wraps run.py + ppi_dashboard.py |
| `START_CHECKER.bat` | — | Windows launcher for run.py |
| `START_DASHBOARD.bat` | — | Windows launcher for ppi_dashboard.py |
| `START_APP.bat` | v1.4.0 | Windows launcher for Streamlit (`streamlit run app.py`) |
| `requirements.txt` | — | Python dependencies |

---

## Folder Structure (self-contained — zip and move)

```
python/
├── run.py
├── ppi_dashboard.py
├── app.py
├── START_CHECKER.bat
├── START_DASHBOARD.bat
├── START_APP.bat
├── requirements.txt
├── db/
│   └── db.xlsx                  ← 3,537-item Part No → HS Code DB (not in repo)
├── ppi/
│   ├── ppi_cvc.xlsx             ← 76 authorized HS codes, CVC/Citi (not in repo)
│   ├── ppi_hsbc.xlsx            ← 177 authorized HS codes, HSBC (not in repo)
│   └── ppi_approval_q1_2026.xlsx  ← gov-approved Q1 quantities (not in repo)
├── invoices/
│   └── incoming/                ← drop invoice .xlsx here before running
└── output/
    └── YYYY-MM/                 ← THO_CHECK_*.xlsx files, auto-organized by month
```

---

## Quick Start

```bash
pip install -r requirements.txt

# 1. Drop your invoice into invoices/incoming/
# 2. Run the checker
python run.py              # GUI popup asks CVC or HSBC
# or double-click START_CHECKER.bat on Windows
```

Output: `output/YYYY-MM/THO_CHECK_<invoice>_<timestamp>.xlsx`

---

## run.py — Invoice Checker

### Pipeline

```
load_invoice(invoice_path)    → df_input
                                    │
                    ┌───────────────┴───────────────┐
                    ↓                               ↓
          run_hs_check(df_input, db_lookup)   run_ppi_check(df_input, bank, ...)
                    ↓  df_hs                        ↓  df_ppi
                    └───────────────┬───────────────┘
                                    ↓
                       build_tho_output(df_hs, df_ppi)
                                    ↓
                          write Excel (5 sheets)
```

Dependencies loaded separately before both checks:
- `load_db()` → `db_lookup: dict {Part No → HS Code}`
- `load_ppi()` × 2 → `ppi_set: set {HS Code}` (one per bank)

### Output Excel — 5 sheets

| Sheet | Contents |
|---|---|
| DASHBOARD | Summary: total lines, HS pass rate, PPI pass rate |
| HS_CHECK | Per-line HS verdict: MATCH / MISMATCH / NEW ITEM |
| PPI_CHECK | Per-line PPI verdict: IN PPI / NOT IN PPI / NO HS CODE |
| THO_OUTPUT | Flagged lines only (union of HS failures and PPI failures) |
| PPI_REF | Copy of the selected PPI list (CVC or HSBC) |

### Verdicts

**HS_CHECK:**
- `MATCH` — invoice HS matches DB historical HS
- `MISMATCH` — invoice HS differs from DB
- `NEW ITEM` — Part No not in DB → escalate to warehouse team

**PPI_CHECK:**
- `IN PPI` — HS code is on the authorized PPI list
- `NOT IN PPI` — HS code not authorized
- `NO HS CODE` — no HS code present on invoice line

### Bank selection

At runtime a GUI popup appears (tkinter). Falls back to CLI prompt if tkinter unavailable:
- `CVC` → uses `ppi/ppi_cvc.xlsx` (spare parts, Citi bank)
- `HSBC` → uses `ppi/ppi_hsbc.xlsx` (kits, HSBC bank)

### HS Normalization (`normalize_hs()`)

- Strips dots and spaces
- Pads to 10 digits with leading zeros
- Handles 11-digit trailing-zero variant (truncates to 10)
- Returns `"-"` for empty / null / nan

### Critical implementation note

`build_tho_output()` uses `pd.concat(axis=1)` — **not** `.merge()`. This is intentional: both `df_hs` and `df_ppi` are filtered with identical Part No masks from the same `df_input`, so row order is guaranteed. Do not change to a key-based merge.

### Invoice format flexibility

`load_invoice()` tries header row 0, then row 1, then falls back to positional column assignment. Detects any Qty column (matches: qty, quant, qté, qte) and normalizes it to `"Qty"` — additive, non-breaking if no Qty column present.

---

## ppi_dashboard.py — PPI Quota Dashboard

Standalone script. Reads a government-approved PPI quota file and compares against consumed quantities from processed output files.

> **Note:** Rename your approval file to `ppi_approval_q1_2026.xlsx` before placing it in `ppi/`.

### Pipeline

```
load_quota(ppi/ppi_approval_q1_2026.xlsx)
  sheet: "PM Data (2)", filter: Remark == 'Approved'
  Category == 'SVC'   → cvc_dict  {hs: {qty, desc}}
  Category != 'SVC'   → hsbc_dict {hs: {qty, desc}}

load_consumption(output/)
  scans THO_CHECK_*.xlsx → PPI_CHECK sheet
  sums Qty by HS code + Bank where PPI Status == 'IN PPI'

build_dashboard() → merge quota + consumption → Remaining + % Used per bank
```

Output: `output/PPI_QUOTA_DASHBOARD_YYYY-MM.xlsx` (overwrites monthly)

### PPI Quota Logic

- `SVC` category in approval file → CVC bank quota (spare parts)
- All other categories (TV, REF, WM, AC, …) → HSBC bank quota (kits)
- Same HS code can appear in both banks with separate quantities
- Quota tracked at HS code level — multiple product lines with same HS code are summed

### Known limitations

- "Q1 2026" is hardcoded in the dashboard title — update manually for Q2
- Dashboard scans all output files across all subdirs — Q2 will include Q1 files; reset scope manually for Q2

---

## app.py — Streamlit Web UI (v1.4.0)

Wraps `run.py` and `ppi_dashboard.py` in a browser interface.

```bash
streamlit run app.py
# or double-click START_APP.bat
# Browser opens at http://localhost:8501
```

Requires `run.py` and `ppi_dashboard.py` in the same directory.
Theme: deep navy ops console, electric teal accent (Syne + JetBrains Mono fonts).

---

## Dependencies

```
pandas>=2.0
openpyxl>=3.1
streamlit          # only needed for app.py
```

```bash
pip install -r requirements.txt
```

---

## Building the .exe (PyInstaller)

```bash
pyinstaller --onefile --windowed run.py --name THO_Checker
```

The existing `THO_Checker.exe` (33 MB) was built 2026-03-31. It needs a rebuild after any `run.py` changes (Qty patch added 2026-04-01).
