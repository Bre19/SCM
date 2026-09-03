from __future__ import annotations

import csv
import json
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

from .classification import classify_po
from .constants import (
    CHECKMARK,
    CURRENT_MASTER_SOURCES,
    LEGACY_MASTER_SOURCES,
    LEVEL_COLUMNS,
    OUTPUT_COLUMNS,
    SOURCE_PRECEDENCE,
)
from .matching import MatchResult, choose_best, match_sources_to_po
from .normalize import clean_text, distinct_nonempty, normalize_name
from .readers import read_inputs


def load_settings(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        settings = json.load(handle)
    required_categories = {"A", "B", "C", "D", "E"}
    actual = set(settings.get("categories", {}))
    if actual != required_categories:
        raise ValueError(
            f"Konfigurasi categories harus tepat {sorted(required_categories)}, diperoleh {sorted(actual)}"
        )
    return settings


def _source_records(
    by_source: dict[str, list[dict[str, Any]]], source: str
) -> list[dict[str, Any]]:
    return by_source.get(source, [])


def _best_by_source(by_source: dict[str, list[dict[str, Any]]]) -> dict[str, dict[str, Any]]:
    return {
        source: best
        for source in SOURCE_PRECEDENCE
        if (best := choose_best(_source_records(by_source, source))) is not None
    }


def _preferred(best: dict[str, dict[str, Any]], field: str, po_value: str = "") -> str:
    values = [best[source].get(field, "") for source in SOURCE_PRECEDENCE if source in best]
    values.append(po_value)
    for value in values:
        if clean_text(value):
            return clean_text(value)
    return ""


def _entity_type(by_source: dict[str, list[dict[str, Any]]], default: str) -> str:
    for source in SOURCE_PRECEDENCE:
        records = _source_records(by_source, source)
        if records:
            return "Perorangan/Mandor" if records[0]["entity_type"] == "INDIVIDUAL" else "Perusahaan"
    return default


def _category(
    has_current: bool,
    has_legacy: bool,
    has_candidate: bool,
    has_dbcr: bool,
) -> str:
    if has_current:
        return "A"
    if has_legacy:
        return "B"
    if has_candidate:
        return "C"
    if has_dbcr:
        return "D"
    return "E"


def _most_common_name(names: list[str]) -> str:
    clean = [clean_text(name) for name in names if clean_text(name)]
    if not clean:
        return ""
    counts = Counter(normalize_name(name) for name in clean)
    winning_key = sorted(counts, key=lambda key: (-counts[key], key))[0]
    return next(name for name in clean if normalize_name(name) == winning_key)


def _conflict_review(
    sap: str,
    by_source: dict[str, list[dict[str, Any]]],
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for source, records in by_source.items():
        if len(records) > 1:
            rows.append(
                {
                    "Severity": "MEDIUM",
                    "Issue": "DUPLICATE_SOURCE_MATCH",
                    "Source": source,
                    "Source Row": ", ".join(str(record["source_row"]) for record in records[:20]),
                    "ID Vendor": " | ".join(distinct_nonempty(record["id_vendor"] for record in records)),
                    "NO SAP": sap,
                    "Nama Rekanan": " | ".join(distinct_nonempty(record["name"] for record in records)),
                    "Match Method": "MULTIPLE_RECORDS",
                    "Detail": f"{len(records)} record {source} terhubung ke SAP yang sama; record terlengkap dipilih.",
                }
            )

    all_records = [record for records in by_source.values() for record in records]
    field_specs = (
        ("id_vendor", "ID_VENDOR_CONFLICT", "HIGH"),
        ("npwp", "NPWP_CONFLICT", "HIGH"),
        ("name", "NAME_CONFLICT", "MEDIUM"),
    )
    for field, issue, severity in field_specs:
        values = distinct_nonempty(record.get(field, "") for record in all_records)
        normalized = {normalize_name(value) if field == "name" else value for value in values}
        if len(normalized) > 1:
            rows.append(
                {
                    "Severity": severity,
                    "Issue": issue,
                    "Source": "MULTI_SOURCE",
                    "Source Row": "",
                    "ID Vendor": " | ".join(
                        distinct_nonempty(record["id_vendor"] for record in all_records)
                    ),
                    "NO SAP": sap,
                    "Nama Rekanan": " | ".join(
                        distinct_nonempty(record["name"] for record in all_records)
                    ),
                    "Match Method": "SOURCE_RECONCILIATION",
                    "Detail": f"Nilai {field} berbeda: " + " | ".join(values[:10]),
                }
            )
    return rows


def build_output_rows(
    po: pd.DataFrame,
    classified: dict[str, dict[str, Any]],
    matches: MatchResult,
    settings: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    output: list[dict[str, Any]] = []
    reviews = list(matches.review_rows)
    categories = settings["categories"]
    default_entity_type = settings.get("po_only_entity_type", "Perusahaan")

    for sap in sorted(classified, key=lambda key: (normalize_name(_most_common_name(classified[key]["names"])), key)):
        po_info = classified[sap]
        by_source = matches.matched.get(sap, {})
        best = _best_by_source(by_source)
        has_current = any(_source_records(by_source, source) for source in CURRENT_MASTER_SOURCES)
        has_legacy_raw = any(_source_records(by_source, source) for source in LEGACY_MASTER_SOURCES)
        has_legacy_effective = has_legacy_raw and not has_current
        has_candidate = bool(_source_records(by_source, "DCR") or _source_records(by_source, "DCM"))
        has_dbcr = bool(_source_records(by_source, "DBCR"))
        category = _category(has_current, has_legacy_effective, has_candidate, has_dbcr)
        po_name = _most_common_name(po_info["names"])

        circle_values = distinct_nonempty(
            record.get("circle", "")
            for source in ("DRT", "DM", "DRT_LAMA", "DM_LAMA")
            for record in _source_records(by_source, source)
        )
        po_sources = sorted(po_info["po_sources"])

        row: dict[str, Any] = {
            "ID Vendor": _preferred(best, "id_vendor"),
            "NO SAP": sap,
            "Nama Rekanan": _preferred(best, "name", po_name),
            "NPWP": _preferred(best, "npwp"),
            "PO": CHECKMARK,
            "Master Data Vendor SAP": CHECKMARK if (has_current or has_legacy_raw) else "",
            "DRT": CHECKMARK if has_current else "",
            "DRT Lama": CHECKMARK if has_legacy_effective else "",
            "Inject": "" if (has_current or has_legacy_raw) else CHECKMARK,
            "DCR": CHECKMARK if _source_records(by_source, "DCR") else "",
            "DCM": CHECKMARK if _source_records(by_source, "DCM") else "",
            "DBCR": CHECKMARK if _source_records(by_source, "DBCR") else "",
            "Kategori": category,
            "Perlakuan": categories[category]["treatment"],
            "Kualifikasi": _preferred(
                {key: value for key, value in best.items() if key in {"DRT", "DRT_LAMA"}},
                "qualification",
            ),
            "Cakupan Wilayah": _preferred(
                {key: value for key, value in best.items() if key in {"DRT", "DRT_LAMA"}},
                "coverage",
            ),
            "Bidang Usaha": _preferred(
                {key: value for key, value in best.items() if key in {"DRT", "DRT_LAMA"}},
                "business_field",
            ),
            "Badan Usaha": _entity_type(by_source, default_entity_type),
            "Klasifikasi Circle": "\n".join(circle_values),
            "Klasifikasi Final": po_info["final_classification"],
            "Item Pekerjaan Berdasarkan PO": po_info["item_text"],
            LEVEL_COLUMNS[0]: "",
            LEVEL_COLUMNS[1]: "",
            LEVEL_COLUMNS[2]: "",
            "Saldo Hutang": "",
        }
        output.append({column: row[column] for column in OUTPUT_COLUMNS})

        if not by_source:
            reviews.append(
                {
                    "Severity": "HIGH",
                    "Issue": "PO_VENDOR_NO_REGISTRY_MATCH",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": "",
                    "NO SAP": sap,
                    "Nama Rekanan": po_name,
                    "Match Method": "PO_ONLY",
                    "Detail": "Vendor ada di PO tetapi tidak terhubung ke DRT/DM/DCR/DCM/DBCR.",
                }
            )
        if not row["ID Vendor"]:
            reviews.append(
                {
                    "Severity": "MEDIUM",
                    "Issue": "MISSING_ID_VENDOR",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": "",
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "OUTPUT_COMPLETENESS",
                    "Detail": "ID Vendor belum tersedia dari sumber yang berhasil dicocokkan.",
                }
            )
        if not row["NPWP"]:
            reviews.append(
                {
                    "Severity": "LOW",
                    "Issue": "MISSING_NPWP",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": row["ID Vendor"],
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "OUTPUT_COMPLETENESS",
                    "Detail": "NPWP belum tersedia dari sumber yang berhasil dicocokkan.",
                }
            )
        elif len(row["NPWP"]) not in {15, 16}:
            reviews.append(
                {
                    "Severity": "MEDIUM",
                    "Issue": "INVALID_NPWP_LENGTH",
                    "Source": "MASTER DATA",
                    "Source Row": "",
                    "ID Vendor": row["ID Vendor"],
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "OUTPUT_VALIDATION",
                    "Detail": f"Panjang NPWP {len(row['NPWP'])} digit; expected 15 atau 16 digit.",
                }
            )
        if row["Badan Usaha"] == "Perusahaan" and (has_current or has_legacy_raw):
            missing_attributes = [
                column
                for column in ("Kualifikasi", "Cakupan Wilayah", "Bidang Usaha")
                if not row[column]
            ]
            if missing_attributes:
                reviews.append(
                    {
                        "Severity": "LOW",
                        "Issue": "MISSING_MASTER_ATTRIBUTES",
                        "Source": "DRT/DRT LAMA",
                        "Source Row": "",
                        "ID Vendor": row["ID Vendor"],
                        "NO SAP": sap,
                        "Nama Rekanan": row["Nama Rekanan"],
                        "Match Method": "OUTPUT_COMPLETENESS",
                        "Detail": "Atribut master kosong: " + ", ".join(missing_attributes),
                    }
                )
        if len(po_info["names"]) > 1:
            reviews.append(
                {
                    "Severity": "LOW",
                    "Issue": "PO_NAME_VARIATION",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": row["ID Vendor"],
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "PO_RECONCILIATION",
                    "Detail": "Satu SAP memiliki variasi nama PO: " + " | ".join(po_info["names"][:10]),
                }
            )
        if not po_info["final_classification"]:
            reviews.append(
                {
                    "Severity": "MEDIUM",
                    "Issue": "UNRESOLVED_CLASSIFICATION",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": row["ID Vendor"],
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "NO_RULE_MATCH",
                    "Detail": "Tidak ada deskripsi PO yang cocok dengan rule klasifikasi aktif.",
                }
            )
        if po_info["item_text_truncated"]:
            reviews.append(
                {
                    "Severity": "LOW",
                    "Issue": "ITEM_TEXT_TRUNCATED",
                    "Source": "PO " + "+".join(po_sources),
                    "Source Row": "",
                    "ID Vendor": row["ID Vendor"],
                    "NO SAP": sap,
                    "Nama Rekanan": row["Nama Rekanan"],
                    "Match Method": "EXCEL_CELL_LIMIT",
                    "Detail": "Daftar item dipotong agar tidak melewati batas teks sel Excel.",
                }
            )
        reviews.extend(_conflict_review(sap, by_source))

    return output, reviews


def validate_output(rows: list[dict[str, Any]], expected_vendors: int) -> None:
    if len(rows) != expected_vendors:
        raise RuntimeError(
            f"Row guard gagal: expected {expected_vendors:,}, got {len(rows):,}"
        )
    saps = [clean_text(row["NO SAP"]) for row in rows]
    if any(not sap for sap in saps):
        raise RuntimeError("Output mengandung NO SAP kosong")
    if len(set(saps)) != len(saps):
        raise RuntimeError("Output mengandung duplikasi NO SAP")
    for column in LEVEL_COLUMNS:
        if any(clean_text(row[column]) for row in rows):
            raise RuntimeError(f"Kolom {column!r} wajib kosong pada versi ini")
    if any(clean_text(row["Saldo Hutang"]) for row in rows):
        raise RuntimeError("Saldo Hutang wajib kosong karena belum ada sumber")


def _write_csv(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _summary_rows(
    po: pd.DataFrame,
    sources: dict[str, list[dict[str, Any]]],
    rows: list[dict[str, Any]],
    reviews: list[dict[str, str]],
    evidence: list[dict[str, Any]],
    unresolved: list[dict[str, Any]],
    matches: MatchResult,
    duplicate_po_rows_removed: int,
) -> list[dict[str, Any]]:
    metrics: list[tuple[str, Any, str]] = [
        ("Vendor output unik", len(rows), "Satu baris per NO SAP vendor PO"),
        ("Item PO diproses", len(po), "Gabungan PO HK dan PO JO setelah deduplikasi persis"),
        ("Duplikasi baris PO dibuang", duplicate_po_rows_removed, "Duplikat identik pada sumber PO"),
        ("Vendor terklasifikasi", sum(bool(row["Klasifikasi Final"]) for row in rows), "Minimal satu rule PO cocok"),
        ("Vendor belum terklasifikasi", sum(not bool(row["Klasifikasi Final"]) for row in rows), "Perlu penambahan rule/review"),
        ("Item PO belum terklasifikasi", len(unresolved), "Deskripsi tidak cocok dengan rule aktif"),
        ("Baris evidence klasifikasi", len(evidence), "Agregat vendor dan klasifikasi"),
        ("Temuan review", len(reviews), "Seluruh tingkat severity"),
    ]
    for source, records in sources.items():
        metrics.append((f"Record sumber {source}", len(records), "Sebelum pencocokan ke PO"))
    for category in ("A", "B", "C", "D", "E"):
        metrics.append(
            (f"Kategori {category}", sum(row["Kategori"] == category for row in rows), "Hasil rule kategori")
        )
    for source in SOURCE_PRECEDENCE:
        for method, count in sorted(matches.method_counts.get(source, {}).items()):
            metrics.append((f"Match {source} - {method}", count, "Metode identity resolution"))
    return [
        {"Metrik": metric, "Nilai": value, "Keterangan": note}
        for metric, value, note in metrics
    ]


def run_pipeline(
    raw_dir: Path,
    config_dir: Path,
    output_dir: Path,
    settings_file: Path,
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    settings = load_settings(settings_file)
    po, sources = read_inputs(raw_dir)

    exact_duplicate_mask = po.duplicated(
        subset=["company", "po", "item_po", "sap", "description"], keep="first"
    )
    duplicate_po_rows_removed = int(exact_duplicate_mask.sum())
    po = po.loc[~exact_duplicate_mask].reset_index(drop=True)

    matches = match_sources_to_po(po, sources)
    classified, evidence, unresolved = classify_po(po, config_dir)
    rows, reviews = build_output_rows(po, classified, matches, settings)
    validate_output(rows, po["sap"].nunique())

    summary = _summary_rows(
        po,
        sources,
        rows,
        reviews,
        evidence,
        unresolved,
        matches,
        duplicate_po_rows_removed,
    )

    bundle_path = output_dir / "workbook_bundle.json"
    bundle_path.write_text(
        json.dumps(
            {
                "output_columns": OUTPUT_COLUMNS,
                "data_rows": rows,
                "summary_rows": summary,
                "review_rows": reviews,
                "evidence_rows": evidence,
                "assumptions": settings.get("assumptions", []),
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    _write_csv(output_dir / "data_cleansing.csv", rows, OUTPUT_COLUMNS)
    review_columns = [
        "Severity",
        "Issue",
        "Source",
        "Source Row",
        "ID Vendor",
        "NO SAP",
        "Nama Rekanan",
        "Match Method",
        "Detail",
    ]
    _write_csv(output_dir / "matching_review.csv", reviews, review_columns)
    evidence_columns = [
        "NO SAP",
        "Nama Vendor PO",
        "Rank",
        "Klasifikasi",
        "Jumlah PO Berbeda",
        "Jumlah Item PO",
        "Prioritas Rule",
        "Contoh Deskripsi",
    ]
    _write_csv(output_dir / "classification_evidence.csv", evidence, evidence_columns)
    unresolved_columns = [
        "Company",
        "PO",
        "Item PO",
        "NO SAP",
        "Nama Vendor",
        "Deskripsi",
        "Project",
    ]
    _write_csv(output_dir / "po_unresolved.csv", unresolved, unresolved_columns)
    (output_dir / "run_summary.json").write_text(
        json.dumps(
            {
                "metrics": summary,
                "match_method_counts": matches.method_counts,
                "outside_po_counts": matches.outside_po_counts,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return {
        "bundle": bundle_path,
        "data_csv": output_dir / "data_cleansing.csv",
        "review_csv": output_dir / "matching_review.csv",
        "evidence_csv": output_dir / "classification_evidence.csv",
        "unresolved_csv": output_dir / "po_unresolved.csv",
        "summary_json": output_dir / "run_summary.json",
    }
