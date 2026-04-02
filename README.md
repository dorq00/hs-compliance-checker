# HS Compliance Checker

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-v2.1-217346?logo=microsoft-excel&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-v1.4.0-FF4B4B?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)

Import compliance tooling for electronics supply chains. Takes a raw invoice, validates every line against a parts database and government-authorized import lists, tracks quarterly quotas per bank (CVC for spare parts, HSBC for kits), and generates a freight-forwarder-ready compliance report.

Built to replace a manual, error-prone process that runs on every incoming shipment — one wrong HS code or unauthorized item means a global logistics delay.

## What It Does

1. Normalizes HS codes to the 10-digit standard (strips dots, spaces, inconsistent formatting)
2. Matches each part number against the historical parts database → resolves the correct HS code
3. Routes to the right authorization list — **CVC** (spare parts) or **HSBC** (kits) — and checks compliance
4. Cross-references government-approved import quantities for the current quarter
5. Flags anything not in the DB as `NEW ITEM` — escalates to warehouse, never auto-approves
6. Outputs per-line verdicts + a summary dashboard

## Two Implementations

| | Excel v2.1 | Python v1.3.0 / v1.4.0 |
|---|---|---|
| **Status** | Primary — production use | Secondary — parallel |
| **Location** | `excel/` | `python/` |
| **How to run** | Open xlsx, paste invoice into INPUT, set CONFIG B4 | `START_CHECKER.bat` or `python run.py` |
| **Output** | In-place, same workbook | New file in `output/YYYY-MM/` |
| **Web UI** | — | `START_APP.bat` (Streamlit, v1.4.0) |

The Excel version is the daily driver — no install, no setup, paste and go. The Python version handles scale, automation, and adds a web UI when you want to move off spreadsheets.

## Quick Start — Excel

```
1. Open   excel/hs_checker_v2.1.xlsx
2. CONFIG tab → cell B4 → type CVC or HSBC
3. INPUT tab  → paste invoice data (values only)
4. Done  — all tabs auto-update
```

→ [`excel/README.md`](excel/README.md) for tab-by-tab details

## Quick Start — Python

```
python/
├── run.py                         ← invoice checker
├── ppi_dashboard.py               ← import quota dashboard
├── app.py                         ← Streamlit web UI (v1.4.0)
├── db/db.xlsx                     ← parts database (not in repo)
├── ppi/
│   ├── ppi_cvc.xlsx               ← authorized list, CVC bank
│   ├── ppi_hsbc.xlsx              ← authorized list, HSBC bank
│   └── ppi_approval_q1_2026.xlsx  ← gov-approved quantities
└── invoices/incoming/             ← drop invoice here before running
```

```bash
pip install -r requirements.txt
python run.py          # GUI popup selects bank, output in output/YYYY-MM/
```

Windows: double-click `START_CHECKER.bat`, `START_DASHBOARD.bat`, or `START_APP.bat`

→ [`python/README.md`](python/README.md) for full pipeline and critical implementation notes

## Data Files (not in repo — supply separately)

| File | Location | What it is |
|---|---|---|
| `db.xlsx` | `python/db/` | Part No → HS Code mappings |
| `ppi_cvc.xlsx` | `python/ppi/` | Authorized HS codes, CVC/Citi |
| `ppi_hsbc.xlsx` | `python/ppi/` | Authorized HS codes, HSBC |
| `ppi_approval_q1_2026.xlsx` | `python/ppi/` | Gov-approved quantities, Q1 2026 |
| Invoice `.xlsx` | `python/invoices/incoming/` | Drop here before running |

> The Excel workbook is self-contained — embeds its own database and authorization lists. No external files needed for the Excel implementation.

## Key Rules

- **10-digit HS codes only** — normalization strips dots/spaces, pads to 10 digits
- **CVC vs HSBC split is critical** — wrong list = compliance failure
- **NEW ITEM = escalate** — never auto-approve a part not in the DB
- **No Power Query** — removed, do not add back
- **No FILTER formula injection** — inserting FILTER via XML corrupts the xlsx

## Version History

| Version | Date | What changed |
|---|---|---|
| Excel v2.1 | 2026-03 | Formula-driven 9-tab workbook, production use |
| Python v1.3.0 | 2026-03-31 | `run.py` + import quota dashboard + PyInstaller `.exe` |
| Python v1.4.0 | 2026-04-01 | Streamlit web UI + Qty column passthrough |

## Docs

| | |
|---|---|
| [`docs/architecture.md`](docs/architecture.md) | Data flow, check logic, verdict definitions |
| [`docs/data-sources.md`](docs/data-sources.md) | Database, authorization lists, approval file |
| [`excel/README.md`](excel/README.md) | Tab-by-tab Excel reference |
| [`python/README.md`](python/README.md) | Script reference, pipeline, critical notes |
