from __future__ import annotations

OUTPUT_COLUMNS = [
    "ID Vendor",
    "NO SAP",
    "Nama Rekanan",
    "NPWP",
    "PO",
    "Master Data Vendor SAP",
    "DRT",
    "DRT Lama",
    "Inject",
    "DCR",
    "DCM",
    "DBCR",
    "Kategori",
    "Perlakuan",
    "Kualifikasi",
    "Cakupan Wilayah",
    "Bidang Usaha",
    "Badan Usaha",
    "Klasifikasi Circle",
    "Klasifikasi Final",
    "Item Pekerjaan Berdasarkan PO",
    "Kelompok Klasifikasi\nLevel 1 (Klasifikasi PO)",
    "Kelompok Klasifikasi\nLevel 2 (Klasifikasi PO)",
    "Kelompok Klasifikasi\nLevel 3 (Klasifikasi PO)",
    "Saldo Hutang",
]

LEVEL_COLUMNS = OUTPUT_COLUMNS[21:24]

REQUIRED_INPUT_FILES = {
    "PO_HK": "PO HK.xlsx",
    "PO_JO": "PO JO.xlsx",
    "DBCR": "DBCR.xls",
    "DCR": "DCR.xls",
    "DRT": "DRT.xls",
    "DRT_LAMA": "DRT Lama.xls",
    "DCM": "DCM.xls",
    "DM": "DM.xls",
    "DM_LAMA": "DM Lama.xls",
}

CURRENT_MASTER_SOURCES = ("DRT", "DM")
LEGACY_MASTER_SOURCES = ("DRT_LAMA", "DM_LAMA")
CANDIDATE_SOURCES = ("DCR", "DCM", "DBCR")

SOURCE_PRECEDENCE = (
    "DRT",
    "DM",
    "DRT_LAMA",
    "DM_LAMA",
    "DCR",
    "DCM",
    "DBCR",
)

CHECKMARK = "✓"
EXCEL_CELL_TEXT_LIMIT = 32_000
