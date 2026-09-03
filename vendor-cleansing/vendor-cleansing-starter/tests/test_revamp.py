from __future__ import annotations

from pathlib import Path
import sys
import unittest

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.revamp.constants import LEVEL_COLUMNS  # noqa: E402
from src.revamp.matching import MatchResult, match_sources_to_po  # noqa: E402
from src.revamp.normalize import canonical_name, normalize_identifier, normalize_npwp  # noqa: E402
from src.revamp.pipeline import _category, build_output_rows, validate_output  # noqa: E402


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
        self.assertIn("UNRESOLVED_CLASSIFICATION", issues)


if __name__ == "__main__":
    unittest.main()
