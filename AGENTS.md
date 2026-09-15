# Yanou IT Asset Reconciliation

This repository is a Snorkel/Harbor-style data task. The "application" is a set of
Python data-processing scripts (`fix_*.py`, `apply_hardcoded_values_and_synthetic_fixes.py`)
that read the source inputs (CSV/XLSX/DOCX/PNG/PDF) and produce/repair the graded
deliverable workbook `Yanou_IT_Asset_Reconciliation.xlsx`. `tests/check_outputs.py` is a
fail-closed gate that validates the deliverable by name and OOXML type. Task spec and
grading live in `task_prompt.txt`, `task.toml`, and `evaluation_rubric.md`.

## Cursor Cloud specific instructions

- Runtime: system Python 3 (`python3`, currently 3.12). Dependencies (`pandas`, `openpyxl`,
  `python-docx`, `pypdf`, `reportlab`, `Pillow`) are installed to the user site via the
  update script (`requirements.txt`). No virtualenv is used; run `python3` directly.
- Test / lint gate: `python3 tests/check_outputs.py .` — passes only when
  `Yanou_IT_Asset_Reconciliation.xlsx` exists and is a valid OOXML workbook. It intentionally
  fails closed and does NOT accept the legacy `Meridian_*` name as the graded deliverable.
- Deliverable naming: `Yanou_IT_Asset_Reconciliation.xlsx` is the graded name. The `Meridian_*`
  twins are retained only for backward compatibility with older graders; keep them in sync.
- GOTCHA: the `fix_*.py` / `apply_*.py` scripts are one-off dataset-repair scripts that mutate
  files in place using `ROOT = Path(__file__).resolve().parent` (the repo root). Running their
  `main()` rewrites the committed workbook, inputs, and zip bundles. To exercise the pipeline
  without dirtying tracked files, copy the repo to a temp dir (exclude `.git`) and run there.
- Safe read-only smoke checks (no mutation): `python3 -c "import fix_golden_fidelity as f; f.verify()"`
  (PO vs inventory vs ledger cross-check) and `f.rebuild_zips()` only repackages existing files
  into the `*.zip` bundles.
- The scripts import each other by module name (e.g. `import fix_golden_fidelity`), so run them
  with the repo root as the working directory.
