# Architecture

## Overview

The HS Compliance Checker validates import invoices against two independent checks:

1. **HS Check** — Does the invoice HS code match the historical database?
2. **Authorization Check** — Is the invoice HS code on the bank's authorized import list?

Both checks are independent. THO_OUTPUT is the union of failures from either check.

---

## Excel Architecture

Single workbook. All logic is formula-driven. No macros, no Power Query.

```
CONFIG (B4: CVC|HSBC)
       ↓
INPUT (paste invoice here)
       ↓
HS_CHECK  ←  DB tab (3,537 items)
PPI_CHECK ←  PPI_CVC (76 codes) or PPI_HSBC (177 codes), switched by CONFIG B4
       ↓
THO_OUTPUT  (flagged lines only)
DASHBOARD   (auto-updating summary)
```

### Tab dependency order

1. **CONFIG** — controls PPI list selection (B4)
2. **INPUT** — raw invoice data (paste values only)
3. **DB** — historical Part No → HS Code lookup table
4. **PPI_CVC / PPI_HSBC** — authorized HS code lists
5. **HS_CHECK** — formula output: HS verdict per line
6. **PPI_CHECK** — formula output: PPI verdict per line
7. **THO_OUTPUT** — formula output: flagged lines only
8. **DASHBOARD** — formula output: summary counts and rates

---

## Python Architecture

### run.py pipeline

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
             write Excel → output/YYYY-MM/THO_CHECK_<name>_<timestamp>.xlsx
                           Sheets: DASHBOARD, HS_CHECK, PPI_CHECK, THO_OUTPUT, PPI_REF
```

Dependencies:
- `load_db()` → `db_lookup: dict {Part No → HS Code}` (3,537 items)
- `load_ppi()` × 2 → `ppi_set: set {HS Code}` (76 or 177 codes)

### Critical design decisions

**`build_tho_output` uses `pd.concat(axis=1)`, not `.merge()`**
Both `df_hs` and `df_ppi` are produced from the same `df_input` with identical Part No filtering, so row order is guaranteed identical. A key-based merge would be wrong here — do not change this.

**`normalize_hs()` is duplicated in both `run.py` and `ppi_dashboard.py`**
Intentional — `ppi_dashboard.py` is standalone and must not import from `run.py`.

---

## Check Logic

### HS Check

```
For each invoice line:
  Part No → lookup in DB
  if not found:              → NEW ITEM
  elif HS codes match:       → MATCH
  else:                      → MISMATCH
```

### PPI Check

```
For each invoice line:
  HS Code (normalized) → membership in selected PPI set
  if HS code is empty:        → NO HS CODE
  elif HS code in PPI set:    → IN PPI
  else:                       → NOT IN PPI
```

### THO_OUTPUT filter

A line appears in THO_OUTPUT if:
- `HS Verdict` is `MISMATCH` or `NEW ITEM`, **OR**
- `PPI Status` is `NOT IN PPI`

Action Requise (French, for freight forwarder submission):
- NEW ITEM → "Nouvel article — Validation requise"
- MISMATCH → "Corriger code HS — voir suggestion DB"
- NOT IN PPI → "Code HS non autorisé dans PPI"

---

## HS Code Normalization

Handles multiple formats encountered in invoice files:

| Input | Output | Note |
|---|---|---|
| `8450.90.9900` | `8450909900` | Strip dots |
| `8450 90 9900` | `8450909900` | Strip spaces |
| `84509099` | `0084509099` | Pad to 10 digits |
| `84509099000` | `8450909900` | 11-digit trailing-zero → truncate to 10 |
| `""` / `NaN` / `None` | `"-"` | Empty sentinel |

---

## Import Quota Dashboard (ppi_dashboard.py)

Tracks authorized import quantities vs consumed quantities per HS code per bank.

```
Approval file (gov-approved quantities)
  Sheet: "PM Data (2)", filter: Remark == 'Approved'
  Category == 'SVC'   → CVC quota
  Category != 'SVC'   → HSBC quota (TV, REF, WM, AC, ...)

Output files (actual consumption)
  Scan: output/**/THO_CHECK_*.xlsx
  Sheet: PPI_CHECK, filter: PPI Status == 'IN PPI'
  Sum Qty by HS Code + Bank

Dashboard = quota - consumed = Remaining + % Used
```

Output columns: HS Code, Description, Q1 Auth CVC, Consumed CVC, Remaining CVC, % CVC Used, Q1 Auth HSBC, Consumed HSBC, Remaining HSBC, % HSBC Used

Color coding: ≥90% used → red, ≥70% → orange, <70% → green.

---

## Color Scheme (Python output)

| Verdict | Fill | Font |
|---|---|---|
| MATCH / IN PPI | Green `#C6EFCE` | Dark green `#006100` |
| MISMATCH / NOT IN PPI / CHECK - HS MISMATCH | Red `#FFC7CE` | Dark red `#9C0006` |
| NEW ITEM | Orange `#FFEB9C` | Dark orange `#9C5700` |

Tab colors: DASHBOARD `#2E4057`, HS_CHECK `#375623`, PPI_CHECK `#7030A0`, THO_OUTPUT `#C55A11`, PPI_REF `#4472C4`
