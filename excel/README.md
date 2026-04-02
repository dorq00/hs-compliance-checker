# Excel Implementation — v2.1

File: `hs_checker_v2.1.xlsx`
Status: **Primary implementation — 99% working, production use.**

---

## Tab Overview

| Tab | Purpose |
|---|---|
| **DASHBOARD** | Auto-updating summary: total lines, HS matches/mismatches/new items, PPI compliant/non-compliant counts, pass rates |
| **CONFIG** | Single configuration cell (B4): set to `CVC` or `HSBC` before each run |
| **INPUT** | Paste zone — invoice data goes here (values only, no formulas). Columns: Line N°, Part No, Designation, Class, HS Code (Invoice), Qty, Bank |
| **DB** | Historical database — 3,537 Part No → HS Code mappings. Columns: Part No, HS Code, Designation, Class. Swappable: paste any clean DB here, keep same column order. |
| **HS_CHECK** | HS code validation engine. Columns: Line N°, Part No, Class, Designation, HS Invoice (raw), HS Invoice (norm), DB Historical HS, HS Match Result, Volatile Item |
| **PPI_CHECK** | PPI compliance validation. Columns: Line N°, Part No, Designation, Class, HS Code (Normalized), PPI Status, —, Invoice Qty, — |
| **THO_OUTPUT** | Flagged lines only (non-compliant or new items). Bilingual headers (FR/EN). Columns: Ligne/Line N°, Code Article/Item, Désignation, Class, Code HS Facture/Invoice HS, Statut PPI/PPI Status, Code HS Historique/DB HS, Action Requise/Required Action |
| **PPI_CVC** | Reference list — 76 authorized HS codes for CVC/Citi (spare parts). Columns: HS Code, Description |
| **PPI_HSBC** | Reference list — 177 authorized HS codes for HSBC (kits). Columns: HS Code, Description |

---

## How to Use

1. **Set bank** — CONFIG tab, cell B4: type `CVC` (spare parts) or `HSBC` (kits)
2. **Paste invoice** — INPUT tab, starting row 3. Paste as **values only** (`Ctrl+Alt+V` → select Values, or right-click → Paste Special → Values). Do not paste formulas from the invoice source.
3. **Read results** — all other tabs auto-update via formulas.

---

## CONFIG Tab

| Cell | Value | Options | Effect |
|---|---|---|---|
| B4 | `CVC` | `CVC` or `HSBC` | PPI_CHECK reads PPI_CVC or PPI_HSBC accordingly |

Warning (from the sheet): *"Change B4 ONCE per invoice run. CVC for spare parts invoices. HSBC for kit invoices."*

---

## Verdicts

**HS_CHECK — HS Match Result column:**

| Verdict | Meaning |
|---|---|
| `MATCH` | HS code on invoice matches DB historical HS code |
| `MISMATCH` | HS code on invoice differs from DB — needs correction |
| `NEW ITEM` | Part No not found in DB — escalate to warehouse team |

**PPI_CHECK — PPI Status column:**

| Verdict | Meaning |
|---|---|
| `IN PPI` | HS code is on the authorized PPI list for the selected bank |
| `NOT IN PPI` | HS code is not authorized — compliance failure |

**THO_OUTPUT** shows only flagged lines: MISMATCH, NEW ITEM, or NOT IN PPI.

---

## DASHBOARD Metrics

**HS Code Validation:**
- Total Lines Checked
- HS Matches / HS Mismatches / New Items (Flagged)
- HS Pass Rate

**PPI Compliance:**
- PPI Compliant / PPI Non-Compliant
- Qty Exceeded
- PPI Pass Rate

---

## DB Tab

- Columns: Part No, HS Code, Designation, Class
- Header row 2, data from row 3
- **In this repo:** 3 dummy rows are included to illustrate format. Replace with your real DB (3,537+ items) starting row 3.
- HS codes are 10-digit standard

---

## PPI Tabs

- **PPI_CVC**: 76 HS codes — Citi bank, spare parts
- **PPI_HSBC**: 177 HS codes — HSBC bank, kits
- Both tabs: header row 2, data from row 3. Columns: HS Code, Description
- PPI check is **presence-only** — no quantity limits in these tabs

---

## Known Constraints

- **No Power Query** — removed, do not add back.
- **No FILTER formula** — injecting FILTER via XML manipulation corrupts the xlsx file structure. Use VLOOKUP/INDEX-MATCH or manual paste instead.
- Paste invoice data as **values only** — pasting with formulas can break the input range.
