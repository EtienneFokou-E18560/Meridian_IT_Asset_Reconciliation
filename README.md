# SnorkelAI Learning — Yanou IT Asset Reconciliation

This repository contains the full working set of artifacts for the Yanou IT Asset Reconciliation task:

- `Yanou_IT_Asset_Reconciliation.xlsx` (primary deliverable workbook)
- `Yanou_IT_Asset_Reconciliation.zip` (submission bundle generated from the workbook)
- `Yanou_IT_Asset_Inputs.zip` (input bundle used to generate the reconciliation)
- `task.toml` (sector / occupation / difficulty metadata for platform upload)
- `tests/check_outputs.py` (fail-closed gate: requires `Yanou_IT_Asset_Reconciliation.xlsx` by name and OOXML type)
- Source inputs (CSV/XLSX/DOCX/PNG/PDF) used by the reconciliation workflow
- Python repair scripts (`fix_oracle_fidelity.py`, `fix_authorship_columns.py`, etc.)

## Sector and occupation

- Sector: Professional, Scientific, and Technical Services
- O*NET Occupation: 13-2011.00 Accountants and Auditors

O*NET Tasks (paste one per line):

```
Examine inventory to verify journal and ledger entries.
Review accounts for discrepancies and reconcile differences.
Prepare, examine, or analyze accounting records, financial statements, or other financial reports.
Report to management regarding the finances of establishment.
```

## Repo contents

The `.xlsx`/`.zip` artifacts are intentionally included for evaluation/verification runs.

Platform verifier note: input-integrity checks (writable inputs, snapshot failure) must fail closed (non-zero / reward 0). Do not warn-and-continue or fall back to a live mutable input directory.
