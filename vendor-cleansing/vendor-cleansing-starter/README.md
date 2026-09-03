# Vendor Cleansing - Stage 1

Stage 1 intentionally does NOT fill `Klasifikasi Final` yet.
It freezes source files, audits anomalies/duplicates, and extracts the legacy uppercase vocabulary.

## 1. Put the three immutable source files in `data/raw/`

- Data Cleansing.xlsx
- Data PO.xlsx
- Data Klasifikasi Kelompok.xlsx

Do not manually edit these files after starting the run.

## 2. Create virtualenv

Windows PowerShell:

    py -3 -m venv .venv
    .\.venv\Scripts\Activate.ps1
    python -m pip install --upgrade pip
    pip install -r requirements.txt

## 3. Run Stage 1

    python scripts/01_audit_and_vocabulary.py

Outputs:

- `output/01_audit/Data Cleansing - Audit Review.xlsx`
- `output/01_audit/audit_findings.csv`
- `output/01_audit/audit_row_summary.csv`
- `output/01_audit/vocabulary_v1.csv`

No source row is deleted. The script hard-fails unless the main sheet still contains exactly 2,735 data rows.

## Next stage

Stage 2 will build `classification_rules.csv` that maps PO descriptions/taxonomy evidence to the 34 allowed legacy uppercase labels.
Stage 3 will classify each PO item, aggregate by vendor, rank labels by distinct PO count, and fill only blank `Klasifikasi Final` cells.
