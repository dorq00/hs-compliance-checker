# Data Sources

## Files Required (not in repo — supply separately)

All data files contain proprietary part numbers and customs data. They are not committed to the repository.

---

## db.xlsx — Historical Database

| Property | Value |
|---|---|
| Location (Python) | `python/db/db.xlsx` |
| Location (Excel) | Embedded in `DB` tab |
| Row count | 3,537 items |
| Columns | Part No, HS Code, Designation, Class |
| Header row | Row 2 (row 1 is title) |
| Data starts | Row 3 |

**Purpose:** Maps Part Numbers to their validated 10-digit HS codes. Used by the HS Check to detect MATCH / MISMATCH / NEW ITEM.

**Maintenance:** New items not in this DB will be flagged as `NEW ITEM`. Add them only after validation by the warehouse team.

**Swap procedure (Python):** Replace `db/db.xlsx` with any clean DB file — keep the filename `db.xlsx` and the same column structure (Part No, HS Code in first two columns).

---

## ppi_cvc.xlsx — CVC/Citi Authorized HS Codes

| Property | Value |
|---|---|
| Location (Python) | `python/ppi/ppi_cvc.xlsx` |
| Location (Excel) | Embedded in `PPI_CVC` tab |
| Row count | 76 HS codes |
| Columns | HS Code, Description |
| Header row | Row 2 (row 1 is title) |
| Data starts | Row 3 |

**Purpose:** Presence check only — is the invoice HS code on this list? No quantity limits.

**Use case:** Spare parts invoices imported via Citi bank. Select `CVC` in CONFIG B4 (Excel) or at bank selection prompt (Python).

---

## ppi_hsbc.xlsx — HSBC Authorized HS Codes

| Property | Value |
|---|---|
| Location (Python) | `python/ppi/ppi_hsbc.xlsx` |
| Location (Excel) | Embedded in `PPI_HSBC` tab |
| Row count | 177 HS codes |
| Columns | HS Code, Description |
| Header row | Row 2 (row 1 is title) |
| Data starts | Row 3 |

**Purpose:** Presence check only — is the invoice HS code on this list? No quantity limits.

**Use case:** Kit invoices imported via HSBC. Select `HSBC` in CONFIG B4 (Excel) or at bank selection prompt (Python).

---

## ppi_approval_q1_2026.xlsx — Government-Approved PPI Quantities

| Property | Value |
|---|---|
| Location | `python/ppi/ppi_approval_q1_2026.xlsx` |
| Sheet | `PM Data (2)` |
| Filter | `Remark == 'Approved'` |

**Purpose:** Government-approved import quantities per HS code per bank for Q1 2026. Used exclusively by `ppi_dashboard.py` — not referenced by `run.py` or the Excel checker.

**Note:** Rename your source file to `ppi_approval_q1_2026.xlsx` before placing it in `ppi/`.

**Category routing:**
- `Category == 'SVC'` → CVC quota (spare parts, Citi)
- `Category != 'SVC'` (TV, REF, WM, AC, …) → HSBC quota (kits)

**Key columns** (detected by keyword, not exact name):
- `category` — product category
- `sous-position` / `sous position` — HS code
- `désignation` / `designation` — product description
- `quantité à importer` / `quantite a importer` — approved quantity
- `remark` — filter for "Approved"

---

## Invoice Files

| Property | Value |
|---|---|
| Location (Python) | `python/invoices/incoming/` |
| Format | `.xlsx` only |

**Purpose:** Raw import invoices from the freight forwarder. Format varies — `run.py` handles flexible header detection.

**Minimum required columns:** `Part No`, `HS Code`

**Optional:** `Line N°`, `Designation`, `Class`, `Qty` (any of: qty, quant, qté, qte)

**Note:** Invoice format changes regularly. If a new format breaks loading, update the `_read_excel()` fallback logic in `run.py`.

---

## Output Files

| Pattern | Location | Contents |
|---|---|---|
| `THO_CHECK_<invoice>_<timestamp>.xlsx` | `python/output/YYYY-MM/` | 5-sheet compliance report |
| `PPI_QUOTA_DASHBOARD_YYYY-MM.xlsx` | `python/output/` | Quota vs consumption per HS code |

Output files are consumed by `ppi_dashboard.py` to calculate consumption — do not move or rename `THO_CHECK_*.xlsx` files out of the `output/` directory tree.
