from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

DATA_DIR = PROJECT_ROOT / "data" / "raw"
OUTPUT_DIR = PROJECT_ROOT / "output"
CONFIG_DIR = PROJECT_ROOT / "config"

CLEANSING_FILE = DATA_DIR / "Data Cleansing.xlsx"
PO_FILE = DATA_DIR / "Data PO.xlsx"
TAXONOMY_FILE = DATA_DIR / "Data Klasifikasi Kelompok.xlsx"

CLEANSING_SHEET = "Sheet1"
PO_SHEET = "Data PO"
TAXONOMY_SHEET = "Sheet1"

EXPECTED_VENDOR_ROWS = 2735

PO_START_DATE = "2024-01-01"
PO_END_DATE = "2026-06-30"

# Data PO.xlsx:
# row 1 kosong
# row 2 kosong
# row 3 header
PO_HEADER = 2

REQUIRED_CLEANSING_COLUMNS = [
    "ID Vendor",
    "NO SAP",
    "Nama Rekanan",
    "NPWP",
    "Klasifikasi Circle",
    "Klasifikasi Final",
    "Item Pekerjaan Berdasarkan PO",
    "Kelompok Klasifikasi\nLevel 1 (Klasifikasi PO)",
    "Kelompok Klasifikasi\nLevel 2 (Klasifikasi PO)",
    "Kelompok Klasifikasi\nLevel 3 (Klasifikasi PO)",
]

L1_ALIASES = {
    "Engineering-Profesional Service":
        "Engineering & Professional Services",

    "Equipment and Plant":
        "Equipment & Plant",
}

NPWP_PLACEHOLDERS = {
    "",
    "0",
    "\\N",
    "N/A",
    "NA",
    "NULL",
    "NONE",
    "-",
    "1000000000000000",
}

VOCABULARY_OVERRIDES = {
    "AC": {
        "enabled": False,
        "mapped_to": "MEKANIKAL",
        "note": (
            "AC/HVAC diperlakukan sebagai evidence HVAC. "
            "Final classification menggunakan MEKANIKAL."
        ),
    },
}

SEVERITY = {
    "CRITICAL": 4,
    "HIGH": 3,
    "MEDIUM": 2,
    "LOW": 1,
    "INFO": 0,
}

ROW_FILL = {
    "CRITICAL": "F4CCCC",  # merah muda
    "HIGH": "FCE5CD",      # orange muda
    "MEDIUM": "FFF2CC",    # kuning
    "LOW": "D9EAF7",       # biru muda
    "INFO": "E2F0D9",      # hijau muda
}

# ============================================================
# STAGE 6 - FINAL CLASSIFICATION
# ============================================================

AUDIT_REVIEW_FILE = (
    OUTPUT_DIR
    / "01_audit"
    / "Data Cleansing - Audit Review.xlsx"
)

CLASSIFICATION_RULES_FILE = (
    CONFIG_DIR
    / "classification_rules.csv"
)

APPROVED_VOCABULARY_FILE = (
    CONFIG_DIR
    / "vocabulary_v1.csv"
)

FINAL_OUTPUT_DIR = (
    OUTPUT_DIR
    / "06_final"
)

PREVIEW_OUTPUT_FILE = (
    FINAL_OUTPUT_DIR
    / "Data Cleansing - Classification Preview.xlsx"
)

FINAL_OUTPUT_FILE = (
    FINAL_OUTPUT_DIR
    / "Data Cleansing - Final.xlsx"
)

# Data PO.xlsx:
# header berada pada Excel row 3.
PO_HEADER = 2

REQUIRED_PO_COLUMNS = [
    "Doc.Date",
    "PO",
    "Item.PO",
    "Vendor",
    "Nama Vendor",
    "Deskripsi",
]

WRITE_PREVIEW_ONLY = True

REQUIRE_ALL_CLASSIFIED_FOR_FINAL = True

# ============================================================
# CLASSIFICATION V2
# ============================================================

VOCABULARY_V2_FILE = (
    CONFIG_DIR
    / "vocabulary_v2.csv"
)

PO_RULES_V2_FILE = (
    CONFIG_DIR
    / "po_rules_v2.csv"
)

CONTEXT_RULES_V2_FILE = (
    CONFIG_DIR
    / "context_rules_v2.csv"
)

V2_OUTPUT_DIR = (
    OUTPUT_DIR
    / "06_final_v2"
)

V2_PREVIEW_OUTPUT_FILE = (
    V2_OUTPUT_DIR
    / "Data Cleansing - Classification Preview V2.xlsx"
)

V2_FINAL_OUTPUT_FILE = (
    V2_OUTPUT_DIR
    / "Data Cleansing - Final V2.xlsx"
)

WRITE_V2_PREVIEW_ONLY = True

# Hijau muda:
# Klasifikasi Final mengandung vocabulary baru.
NEW_VOCABULARY_FILL = "E2F0D9"

# Ungu muda:
# Belum berhasil diklasifikasikan.
UNRESOLVED_CLASSIFICATION_FILL = "D9D2E9"

# ============================================================
# CLASSIFICATION V2.1
# ============================================================

CLASSIFICATION_EXCLUSIONS_V2_FILE = (
    CONFIG_DIR
    / "classification_exclusions_v2.csv"
)

V2_1_OUTPUT_DIR = (
    OUTPUT_DIR
    / "06_final_v2_1"
)

V2_1_PREVIEW_OUTPUT_FILE = (
    V2_1_OUTPUT_DIR
    / "Data Cleansing - Classification Preview V2.1.xlsx"
)

V2_1_FINAL_OUTPUT_FILE = (
    V2_1_OUTPUT_DIR
    / "Data Cleansing - Final V2.1.xlsx"
)

V2_BASELINE_PREVIEW_FILE = (
    OUTPUT_DIR
    / "06_final_v2"
    / "Data Cleansing - Classification Preview V2.xlsx"
)

WRITE_V2_1_PREVIEW_ONLY = True