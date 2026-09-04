from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd
from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.revamp.constants import LEVEL_COLUMNS, OUTPUT_COLUMNS  # noqa: E402
from src.revamp.classification import classify_po, load_circle_rules  # noqa: E402
from src.revamp.matching import MatchResult, match_sources_to_po  # noqa: E402
from src.revamp.normalize import canonical_name, normalize_identifier, normalize_npwp  # noqa: E402
from src.revamp.pipeline import _category, build_output_rows, validate_output  # noqa: E402
from src.revamp.workbook import build_workbook  # noqa: E402


def vendor_record(source: str, **overrides):
    base = {
        "source": source,
        "source_file": f"{source}.xls",
        "source_row": 1,
        "id_vendor": "",
        "sap": "",
        "name": "PT Contoh Abadi",
        "npwp": "",
        "qualification": "",
        "coverage": "",
        "business_field": "",
        "circle": "",
        "status": "",
        "registration_date": "",
        "approval_date": "",
        "email": "",
        "entity_type": "INDIVIDUAL" if source in {"DM", "DM_LAMA", "DCM"} else "COMPANY",
    }
    base.update(overrides)
    return base


def empty_sources():
    return {source: [] for source in ("DRT", "DRT_LAMA", "DM", "DM_LAMA", "DCR", "DCM", "DBCR")}


class RevampPipelineTests(unittest.TestCase):
    def test_normalizers_preserve_identifiers_and_canonicalize_names(self):
        self.assertEqual(normalize_identifier("00123.0"), "00123")
        self.assertEqual(normalize_npwp("01.234.567.8-901.000"), "012345678901000")
        self.assertEqual(canonical_name("PT. Contoh Abadi, Tbk"), "CONTOH ABADI")

    def test_category_precedence_matches_legacy_business_flow(self):
        self.assertEqual(_category(True, True, True, True), "A")
        self.assertEqual(_category(False, True, True, True), "B")
        self.assertEqual(_category(False, False, True, True), "C")
        self.assertEqual(_category(False, False, False, True), "D")
        self.assertEqual(_category(False, False, False, False), "E")

    def test_candidate_id_links_to_master_sap_before_name(self):
        po = pd.DataFrame([{"sap": "2020000001", "name": "Vendor PO"}])
        sources = empty_sources()
        sources["DM"].append(
            vendor_record("DM", id_vendor="8000001", sap="2020000001", name="Nama Master")
        )
        sources["DCM"].append(
            vendor_record("DCM", id_vendor="8000001", name="Nama Calon Berbeda")
        )
        result = match_sources_to_po(po, sources)
        self.assertEqual(
            result.matched["2020000001"]["DCM"][0]["match_method"], "ID_LINK"
        )

    def test_ambiguous_name_is_not_auto_merged(self):
        po = pd.DataFrame(
            [
                {"sap": "1", "name": "PT Nama Sama"},
                {"sap": "2", "name": "PT Nama Sama"},
            ]
        )
        sources = empty_sources()
        sources["DBCR"].append(vendor_record("DBCR", name="PT Nama Sama"))
        result = match_sources_to_po(po, sources)
        self.assertNotIn("1", result.matched)
        self.assertNotIn("2", result.matched)
        self.assertEqual(result.review_rows[0]["Issue"], "AMBIGUOUS_EXACT_NAME")

    def test_output_keeps_level_columns_and_balance_blank(self):
        po = pd.DataFrame([{"sap": "1"}])
        classified = {
            "1": {
                "names": ["Vendor Satu"],
                "po_sources": {"HK"},
                "item_text": "Pekerjaan contoh",
                "item_text_truncated": False,
                "final_classification": "PEKERJAAN SIPIL STRUCTURE",
            }
        }
        matches = MatchResult(matched={}, review_rows=[], method_counts={}, outside_po_counts={})
        settings = {
            "po_only_entity_type": "Perusahaan",
            "categories": {
                "A": {"treatment": "DRP"},
                "B": {"treatment": "DRP + Daftar ulang"},
                "C": {"treatment": "DRP + Prioritas Approve"},
                "D": {"treatment": "DRP + Update data"},
                "E": {"treatment": "DRP + Daftar ulang"},
            },
        }
        rows, _ = build_output_rows(po, classified, matches, settings)
        validate_output(rows, 1)
        self.assertEqual(rows[0]["Kategori"], "E")
        self.assertEqual(rows[0]["Inject"], "✓")
        self.assertTrue(all(rows[0][column] == "" for column in LEVEL_COLUMNS))
        self.assertEqual(rows[0]["Saldo Hutang"], "")

    def test_output_reports_missing_identity_tax_and_classification(self):
        po = pd.DataFrame([{"sap": "1"}])
        classified = {
            "1": {
                "names": ["Vendor Belum Lengkap"],
                "po_sources": {"JO"},
                "item_text": "Uraian yang belum memiliki rule",
                "item_text_truncated": False,
                "final_classification": "",
            }
        }
        matches = MatchResult(matched={}, review_rows=[], method_counts={}, outside_po_counts={})
        settings = {
            "po_only_entity_type": "Perusahaan",
            "categories": {
                "A": {"treatment": "DRP"},
                "B": {"treatment": "DRP + Daftar ulang"},
                "C": {"treatment": "DRP + Prioritas Approve"},
                "D": {"treatment": "DRP + Update data"},
                "E": {"treatment": "DRP + Daftar ulang"},
            },
        }

        rows, reviews = build_output_rows(po, classified, matches, settings)

        issues = {review["Issue"] for review in reviews}
        self.assertEqual(rows[0]["NO SAP"], "1")
        self.assertIn("PO_VENDOR_NO_REGISTRY_MATCH", issues)
        self.assertIn("MISSING_ID_VENDOR", issues)
        self.assertIn("MISSING_NPWP", issues)
        self.assertIn("PO_RULE_GAP_CIRCLE_EMPTY", issues)

    def test_po_classification_is_specific_and_generic_text_stays_unresolved(self):
        po = pd.DataFrame(
            [
                {
                    "company": "HK", "po": "1", "item_po": "10", "sap": "100",
                    "name": "Vendor MEP", "description": "Pekerjaan MEP gedung",
                    "material": "", "division": "", "project": "P1",
                },
                {
                    "company": "HK", "po": "2", "item_po": "10", "sap": "200",
                    "name": "Vendor Pintu", "description": "Pengadaan fire rated steel door",
                    "material": "", "division": "", "project": "P1",
                },
                {
                    "company": "JO", "po": "3", "item_po": "10", "sap": "300",
                    "name": "Vendor Umum", "description": "Material alat bantu",
                    "material": "", "division": "", "project": "P2",
                },
                {
                    "company": "JO", "po": "4", "item_po": "10", "sap": "400",
                    "name": "Vendor Perancah", "description": "Sewa scafolding",
                    "material": "", "division": "", "project": "P2",
                },
                {
                    "company": "JO", "po": "5", "item_po": "10", "sap": "500",
                    "name": "Vendor Langsir", "description": "Upah langsir besi",
                    "material": "", "division": "", "project": "P2",
                },
                {
                    "company": "JO", "po": "6", "item_po": "10", "sap": "600",
                    "name": "Vendor Upah", "description": "Upah",
                    "material": "", "division": "", "project": "P2",
                },
            ]
        )

        classified, _, unresolved = classify_po(po, PROJECT_ROOT / "config")

        self.assertEqual(classified["100"]["ordered_labels"], ["MECHANICAL DAN ELECTRICAL"])
        self.assertEqual(classified["200"]["ordered_labels"], ["PINTU TAHAN API"])
        self.assertEqual(classified["300"]["ordered_labels"], [])
        self.assertEqual(classified["400"]["ordered_labels"], ["SCAFFOLDING / PERANCAH"])
        self.assertEqual(classified["500"]["ordered_labels"], ["LOGISTICS / TRANSPORT"])
        self.assertEqual(classified["600"]["ordered_labels"], [])
        self.assertEqual(len(unresolved), 2)

    def test_circle_validates_but_does_not_add_unrelated_final_classification(self):
        po = pd.DataFrame([{"sap": "1"}])
        classified = {
            "1": {
                "names": ["Vendor Satu"],
                "po_sources": {"HK"},
                "item_text": "Sewa vibro roller",
                "item_text_truncated": False,
                "ordered_labels": ["EQUIPMENT / RENTAL"],
                "final_classification": "EQUIPMENT / RENTAL",
            }
        }
        drt = vendor_record(
            "DRT", sap="1", name="Vendor Satu", circle="KONSTRUKSI BAJA"
        )
        matches = MatchResult(
            matched={"1": {"DRT": [drt]}},
            review_rows=[],
            method_counts={},
            outside_po_counts={},
        )
        rows, reviews = build_output_rows(
            po,
            classified,
            matches,
            {
                "po_only_entity_type": "Perusahaan",
                "categories": {
                    "A": {"treatment": "DRP"},
                    "B": {"treatment": "DRP + Daftar ulang"},
                    "C": {"treatment": "DRP + Prioritas Approve"},
                    "D": {"treatment": "DRP + Update data"},
                    "E": {"treatment": "DRP + Daftar ulang"},
                },
            },
            circle_rules=load_circle_rules(PROJECT_ROOT / "config"),
        )

        self.assertEqual(rows[0]["Klasifikasi Final"], "EQUIPMENT / RENTAL")
        self.assertNotIn("STEEL / FABRIKASI", rows[0]["Klasifikasi Final"])
        self.assertIn("PO_CIRCLE_NO_OVERLAP", {row["Issue"] for row in reviews})

    def test_output_guard_compares_exact_po_vendor_universe(self):
        row = {column: "" for column in OUTPUT_COLUMNS}
        row.update({"NO SAP": "1", "PO": "✓"})
        with self.assertRaises(RuntimeError):
            validate_output([row], {"2"})

    def test_workbook_builder_creates_two_sheet_auditable_output(self):
        data_row = {column: "" for column in OUTPUT_COLUMNS}
        data_row.update(
            {
                "ID Vendor": "00123",
                "NO SAP": "2020000001",
                "Nama Rekanan": "PT Contoh",
                "NPWP": "012345678901000",
                "PO": "✓",
                "Inject": "✓",
                "Kategori": "E",
                "Perlakuan": "DRP + Daftar ulang",
            }
        )
        bundle = {
            "output_columns": OUTPUT_COLUMNS,
            "data_rows": [data_row],
            "summary_rows": [
                {
                    "Metrik": "Vendor output unik",
                    "Nilai": 1,
                    "Keterangan": "Satu baris per NO SAP",
                }
            ],
            "review_rows": [
                {
                    "Severity": "HIGH",
                    "Issue": "PO_VENDOR_NO_REGISTRY_MATCH",
                    "Source": "PO HK",
                    "Source Row": "",
                    "ID Vendor": "00123",
                    "NO SAP": "2020000001",
                    "Nama Rekanan": "PT Contoh",
                    "Match Method": "PO_ONLY",
                    "Detail": "Vendor belum ditemukan pada registry.",
                }
            ],
            "evidence_rows": [
                {
                    "NO SAP": "2020000001",
                    "Nama Vendor PO": "PT Contoh",
                    "Rank": 1,
                    "Klasifikasi": "EQUIPMENT / RENTAL",
                    "Jumlah PO Berbeda": 1,
                    "Jumlah Item PO": 1,
                    "Rule ID": "PO-001",
                    "Confidence Rule": "HIGH",
                    "Dukungan Circle": "CIRCLE KOSONG",
                    "Sumber Final": "PO",
                    "Contoh Deskripsi": "Sewa alat",
                    "Rule Pattern": "SEWA ALAT",
                }
            ],
            "unresolved_rows": [
                {
                    "Company": "HK",
                    "NO SAP": "2020000001",
                    "Nama Vendor": "PT Contoh",
                    "Deskripsi Belum Terklasifikasi": "Material bantu",
                    "Jumlah Item": 1,
                    "Contoh PO": "450000001",
                    "Contoh Item PO": "10",
                    "Contoh Project": "P1",
                    "Tindakan": "Review",
                }
            ],
            "assumptions": ["PO HK dan PO JO menjadi universe output."],
        }

        with tempfile.TemporaryDirectory() as temporary:
            bundle_path = Path(temporary) / "bundle.json"
            output_path = Path(temporary) / "output.xlsx"
            bundle_path.write_text(json.dumps(bundle), encoding="utf-8")
            build_workbook(bundle_path, output_path)

            workbook = load_workbook(output_path, data_only=False)
            self.assertEqual(workbook.sheetnames, ["Data Cleansing", "Audit"])
            self.assertEqual(workbook["Data Cleansing"]["A2"].value, "00123")
            self.assertEqual(workbook["Data Cleansing"]["D2"].value, "012345678901000")
            self.assertEqual(workbook["Data Cleansing"]["A2"].number_format, "@")
            self.assertIn("DataCleansingTable", workbook["Data Cleansing"].tables)
            self.assertIn("AuditFindingsTable", workbook["Audit"].tables)
            self.assertIn("AuditClassificationEvidenceTable", workbook["Audit"].tables)
            self.assertIn("AuditUnresolvedPOTable", workbook["Audit"].tables)


if __name__ == "__main__":
    unittest.main()
