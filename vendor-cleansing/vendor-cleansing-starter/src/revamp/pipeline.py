from __future__ import annotations

import csv
import json
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

from .classification import (
    classify_po,
    load_circle_rules,
    map_circle_classifications,
)
from .constants import (
    CHECKMARK,
    CURRENT_MASTER_SOURCES,
    LEGACY_MASTER_SOURCES,
    LEVEL_COLUMNS,
    OUTPUT_COLUMNS,
    SOURCE_PRECEDENCE,
)
from .matching import MatchResult, choose_best, match_sources_to_po
from .hierarchy import classify_po_hierarchy, format_vendor_hierarchy
from .normalize import clean_text, distinct_nonempty, normalize_name
from .readers import read_inputs
from .source_audit import audit_vendor_sources


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


def _po_labels(po_info: dict[str, Any]) -> list[str]:
    labels = po_info.get("ordered_labels")
    if labels is not None:
        return [clean_text(label) for label in labels if clean_text(label)]
    return [
        clean_text(label)
        for label in clean_text(po_info.get("final_classification", "")).split(",")
        if clean_text(label)
    ]


def _classification_review(
    sap: str,
    name: str,
    po_sources: list[str],
    po_labels: list[str],
    circle_values: list[str],
    circle_labels: list[str],
) -> dict[str, str]:
    po_set = set(po_labels)
    circle_set = set(circle_labels)
    overlap = po_set & circle_set
    source = "PO " + "+".join(po_sources)

    if not po_labels:
        if circle_values:
            mapped = ", ".join(circle_labels) or "tidak terpetakan"
            return {
                "Severity": "MEDIUM",
                "Issue": "PO_RULE_GAP_CIRCLE_PRESENT",
                "Source": source + " + HK CIRCLE",
                "Source Row": "",
                "ID Vendor": "",
                "NO SAP": sap,
                "Nama Rekanan": name,
                "Match Method": "NO_PO_RULE_MATCH",
                "Detail": (
                    "Item PO tersedia tetapi belum cocok dengan rule. "
                    f"Circle terbaca sebagai: {mapped}. Circle tidak disalin otomatis tanpa bukti PO yang relevan."
                ),
            }
        return {
            "Severity": "MEDIUM",
            "Issue": "PO_RULE_GAP_CIRCLE_EMPTY",
            "Source": source,
            "Source Row": "",
            "ID Vendor": "",
            "NO SAP": sap,
            "Nama Rekanan": name,
            "Match Method": "NO_PO_RULE_MATCH",
            "Detail": "Item PO tersedia, Circle kosong, dan belum ada deskripsi yang cocok dengan rule aktif.",
        }

    if not circle_values:
        return {
            "Severity": "LOW",
            "Issue": "PO_CLASSIFICATION_WITHOUT_CIRCLE",
            "Source": source,
            "Source Row": "",
            "ID Vendor": "",
            "NO SAP": sap,
            "Nama Rekanan": name,
            "Match Method": "PO_EVIDENCE_ONLY",
            "Detail": "Klasifikasi Final berasal dari bukti item PO; Circle tidak tersedia.",
        }

    if not circle_labels:
        return {
            "Severity": "MEDIUM",
            "Issue": "CIRCLE_UNMAPPED",
            "Source": source + " + HK CIRCLE",
            "Source Row": "",
            "ID Vendor": "",
            "NO SAP": sap,
            "Nama Rekanan": name,
            "Match Method": "PO_EVIDENCE_CIRCLE_UNMAPPED",
            "Detail": "Final tetap berasal dari PO; isi Circle belum dapat dipetakan ke vocabulary aktif.",
        }

    if not overlap:
        return {
            "Severity": "MEDIUM",
            "Issue": "PO_CIRCLE_NO_OVERLAP",
            "Source": source + " + HK CIRCLE",
            "Source Row": "",
            "ID Vendor": "",
            "NO SAP": sap,
            "Nama Rekanan": name,
            "Match Method": "PO_CIRCLE_RECONCILIATION",
            "Detail": (
                "Final tetap memakai bukti PO dan tidak menggabungkan Circle secara otomatis. "
                f"PO: {', '.join(po_labels)} | Circle: {', '.join(circle_labels)}"
            ),
        }

    issue = "PO_CIRCLE_AGREEMENT" if po_set == circle_set else "PO_CIRCLE_PARTIAL_SUPPORT"
    detail = (
        "Circle mendukung seluruh klasifikasi PO."
        if issue == "PO_CIRCLE_AGREEMENT"
        else "Circle mendukung sebagian klasifikasi PO; label tambahan tetap harus memiliki bukti PO."
    )
    return {
        "Severity": "LOW",
        "Issue": issue,
        "Source": source + " + HK CIRCLE",
        "Source Row": "",
        "ID Vendor": "",
        "NO SAP": sap,
        "Nama Rekanan": name,
        "Match Method": "PO_CIRCLE_RECONCILIATION",
        "Detail": detail,
    }


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
    circle_rules: list[Any] | None = None,
    hierarchy_by_vendor: dict[str, dict[str, Any]] | None = None,
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
        circle_labels = map_circle_classifications(circle_values, circle_rules or [])
        po_labels = _po_labels(po_info)
        po_sources = sorted(po_info["po_sources"])
        hierarchy_values = format_vendor_hierarchy(
            (hierarchy_by_vendor or {}).get(sap)
        )

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
            "Klasifikasi Final": ", ".join(po_labels),
            "Item Pekerjaan Berdasarkan PO": po_info["item_text"],
            LEVEL_COLUMNS[0]: hierarchy_values[0],
            LEVEL_COLUMNS[1]: hierarchy_values[1],
            LEVEL_COLUMNS[2]: hierarchy_values[2],
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
        classification_review = _classification_review(
            sap=sap,
            name=row["Nama Rekanan"],
            po_sources=po_sources,
            po_labels=po_labels,
            circle_values=circle_values,
            circle_labels=circle_labels,
        )
        classification_review["ID Vendor"] = row["ID Vendor"]
        reviews.append(classification_review)
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


def validate_output(
    rows: list[dict[str, Any]],
    expected_vendors: int | set[str],
) -> None:
    expected_count = (
        len(expected_vendors) if isinstance(expected_vendors, set) else expected_vendors
    )
    if len(rows) != expected_count:
        raise RuntimeError(
            f"Row guard gagal: expected {expected_count:,}, got {len(rows):,}"
        )
    saps = [clean_text(row["NO SAP"]) for row in rows]
    if any(not sap for sap in saps):
        raise RuntimeError("Output mengandung NO SAP kosong")
    if len(set(saps)) != len(saps):
        raise RuntimeError("Output mengandung duplikasi NO SAP")
    if isinstance(expected_vendors, set) and set(saps) != expected_vendors:
        missing = sorted(expected_vendors - set(saps))[:10]
        extra = sorted(set(saps) - expected_vendors)[:10]
        raise RuntimeError(
            "SAP guard gagal; output tidak sama dengan universe PO. "
            f"Missing={missing}, extra={extra}"
        )
    if any(row["PO"] != CHECKMARK for row in rows):
        raise RuntimeError("Seluruh vendor output wajib ditandai memiliki PO")
    valid_level1 = {"Supplier", "Subkontraktor", "Alat", "Jasa Konsultansi", "Jasa Lainnya"}
    for row in rows:
        level_values = [str(row[column] or "").strip() for column in LEVEL_COLUMNS]
        if any(level_values) and not all(level_values):
            raise RuntimeError(
                f"Hierarchy tidak lengkap untuk SAP {row['NO SAP']}: Level 1-3 harus terisi bersama."
            )
        if not all(level_values):
            continue
        level1_lines = level_values[0].splitlines()
        level2_lines = level_values[1].splitlines()
        level3_lines = level_values[2].splitlines()
        if not (len(level1_lines) == len(level2_lines) == len(level3_lines)):
            raise RuntimeError(
                f"Hierarchy tidak sejajar untuk SAP {row['NO SAP']}: jumlah jalur Level 1-3 berbeda."
            )
        for level1, level2, level3 in zip(level1_lines, level2_lines, level3_lines):
            if level1 not in valid_level1:
                raise RuntimeError(f"Level 1 tidak sah untuk SAP {row['NO SAP']}: {level1}")
            code2 = level2.split(" | ", 1)[0]
            code3 = level3.split(" | ", 1)[0]
            if not code2 or code2 != code3:
                raise RuntimeError(
                    f"Relasi Level 2-3 tidak konsisten untuk SAP {row['NO SAP']}: {level2!r} / {level3!r}"
                )
    if any(clean_text(row["Saldo Hutang"]) for row in rows):
        raise RuntimeError("Saldo Hutang wajib kosong karena belum ada sumber")


def _write_csv(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _po_input_reviews(
    po_stats: dict[str, dict[str, Any]],
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for company, stats in po_stats.items():
        for rejected in stats.get("rejected_rows", []):
            rows.append(
                {
                    "Severity": "HIGH",
                    "Issue": "PO_ROW_WITHOUT_VENDOR_SAP",
                    "Source": f"PO {company}",
                    "Source Row": clean_text(rejected.get("source_row")),
                    "ID Vendor": "",
                    "NO SAP": "",
                    "Nama Rekanan": clean_text(rejected.get("name")),
                    "Match Method": "PO_INPUT_VALIDATION",
                    "Detail": (
                        f"PO {clean_text(rejected.get('po'))}, item "
                        f"{clean_text(rejected.get('item_po'))} tidak mempunyai Vendor/SAP; "
                        "baris tidak dapat dibentuk menjadi vendor output. Deskripsi: "
                        f"{clean_text(rejected.get('description'))}"
                    ),
                }
            )
    return rows


def _group_unresolved_items(unresolved: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str], dict[str, Any]] = defaultdict(
        lambda: {
            "companies": set(),
            "name": "",
            "description": "",
            "count": 0,
            "po_examples": [],
            "item_examples": [],
            "projects": [],
        }
    )
    for row in unresolved:
        sap = clean_text(row.get("NO SAP"))
        description = clean_text(row.get("Deskripsi"))
        key = (sap, description.upper())
        group = groups[key]
        group["companies"].add(clean_text(row.get("Company")))
        group["name"] = group["name"] or clean_text(row.get("Nama Vendor"))
        group["description"] = group["description"] or description
        group["count"] += 1
        for field, target in (
            ("PO", "po_examples"),
            ("Item PO", "item_examples"),
            ("Project", "projects"),
        ):
            value = clean_text(row.get(field))
            if value and value not in group[target] and len(group[target]) < 5:
                group[target].append(value)

    result = [
        {
            "Company": "+".join(sorted(value["companies"])),
            "NO SAP": sap,
            "Nama Vendor": value["name"],
            "Deskripsi Belum Terklasifikasi": value["description"],
            "Jumlah Item": value["count"],
            "Contoh PO": " | ".join(value["po_examples"]),
            "Contoh Item PO": " | ".join(value["item_examples"]),
            "Contoh Project": " | ".join(value["projects"]),
            "Tindakan": "Tambahkan rule hanya setelah klasifikasi bisnis terverifikasi.",
        }
        for (sap, _), value in groups.items()
    ]
    return sorted(
        result,
        key=lambda row: (-int(row["Jumlah Item"]), row["NO SAP"], row["Deskripsi Belum Terklasifikasi"]),
    )


def _group_hierarchy_unresolved(
    unresolved: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], dict[str, Any]] = defaultdict(
        lambda: {
            "companies": set(), "name": "", "description": "", "reason": "",
            "detail": "", "count": 0, "po_examples": [], "item_examples": [],
            "source_rows": [], "projects": [],
        }
    )
    for row in unresolved:
        sap = clean_text(row.get("NO SAP"))
        description = clean_text(row.get("Deskripsi"))
        reason = clean_text(row.get("Reason")) or "NO_MATCH"
        group = groups[(sap, description.upper(), reason)]
        group["companies"].add(clean_text(row.get("Company")))
        group["name"] = group["name"] or clean_text(row.get("Nama Vendor"))
        group["description"] = group["description"] or description
        group["reason"] = reason
        group["detail"] = group["detail"] or clean_text(row.get("Detail"))
        group["count"] += 1
        source_row = f"{clean_text(row.get('Company'))}:{clean_text(row.get('Source Row'))}"
        if source_row not in group["source_rows"] and len(group["source_rows"]) < 10:
            group["source_rows"].append(source_row)
        for field, target in (("PO", "po_examples"), ("Item PO", "item_examples"), ("Project", "projects")):
            value = clean_text(row.get(field))
            if value and value not in group[target] and len(group[target]) < 5:
                group[target].append(value)
    result = [
        {
            "Company": "+".join(sorted(value["companies"])), "NO SAP": sap,
            "Nama Vendor": value["name"], "Deskripsi Belum Memiliki Level": value["description"],
            "Alasan": value["reason"], "Detail": value["detail"], "Jumlah Item": value["count"],
            "Contoh PO": " | ".join(value["po_examples"]),
            "Contoh Item PO": " | ".join(value["item_examples"]),
            "Baris Sumber PO": " | ".join(value["source_rows"]),
            "Contoh Project": " | ".join(value["projects"]),
            "Tindakan": "Tambahkan istilah/rule hanya setelah jalur Level 1-3 diverifikasi terhadap master.",
        }
        for (sap, _, _), value in groups.items()
    ]
    return sorted(result, key=lambda row: (-int(row["Jumlah Item"]), row["NO SAP"], row["Deskripsi Belum Memiliki Level"]))


def _enrich_evidence_with_circle(
    evidence: list[dict[str, Any]],
    rows: list[dict[str, Any]],
    circle_rules: list[Any],
) -> None:
    output_by_sap = {clean_text(row["NO SAP"]): row for row in rows}
    for item in evidence:
        output = output_by_sap.get(clean_text(item.get("NO SAP")), {})
        circle_values = [
            clean_text(value)
            for value in str(output.get("Klasifikasi Circle", "")).splitlines()
            if clean_text(value)
        ]
        circle_labels = set(map_circle_classifications(circle_values, circle_rules))
        classification = clean_text(item.get("Klasifikasi"))
        if not circle_values:
            item["Dukungan Circle"] = "CIRCLE KOSONG"
            item["Sumber Final"] = "PO"
        elif not circle_labels:
            item["Dukungan Circle"] = "CIRCLE TIDAK TERPETAKAN"
            item["Sumber Final"] = "PO"
        elif classification in circle_labels:
            item["Dukungan Circle"] = "YA"
            item["Sumber Final"] = "PO + CIRCLE"
        else:
            item["Dukungan Circle"] = "TIDAK"
            item["Sumber Final"] = "PO"


def _summary_rows(
    po: pd.DataFrame,
    sources: dict[str, list[dict[str, Any]]],
    rows: list[dict[str, Any]],
    reviews: list[dict[str, str]],
    evidence: list[dict[str, Any]],
    unresolved: list[dict[str, Any]],
    matches: MatchResult,
    duplicate_po_rows_removed: int,
    po_stats: dict[str, dict[str, Any]],
    source_stats: dict[str, dict[str, Any]],
    unresolved_group_count: int,
    hierarchy_evidence: list[dict[str, Any]],
    hierarchy_unresolved: list[dict[str, Any]],
    hierarchy_unresolved_group_count: int,
) -> list[dict[str, Any]]:
    metrics: list[tuple[str, Any, str]] = [
        ("Vendor output unik", len(rows), "Satu baris per NO SAP vendor PO"),
        ("Item PO diproses", len(po), "Gabungan PO HK dan PO JO setelah deduplikasi persis"),
        ("Dokumen PO unik", po[["company", "po"]].drop_duplicates().shape[0], "Nomor PO unik per perusahaan"),
        ("Duplikasi baris PO dibuang", duplicate_po_rows_removed, "Duplikat identik pada sumber PO"),
        ("Vendor terklasifikasi", sum(bool(row["Klasifikasi Final"]) for row in rows), "Minimal satu rule PO cocok"),
        ("Vendor belum terklasifikasi", sum(not bool(row["Klasifikasi Final"]) for row in rows), "Perlu penambahan rule/review"),
        ("Item PO terklasifikasi", len(po) - len(unresolved), "Minimal satu rule PO cocok"),
        ("Item PO belum terklasifikasi", len(unresolved), "Deskripsi tidak cocok dengan rule aktif"),
        ("Kelompok item PO belum terklasifikasi", unresolved_group_count, "Dikelompokkan per SAP dan deskripsi pada Audit"),
        ("Baris evidence klasifikasi", len(evidence), "Agregat vendor dan klasifikasi"),
        ("Vendor memiliki Level 1-3", len({clean_text(row.get("NO SAP")) for row in hierarchy_evidence}), "Minimal satu item PO cocok ke master hierarchy"),
        ("Vendor belum memiliki Level 1-3", sum(not bool(row[LEVEL_COLUMNS[0]]) for row in rows), "Tidak ada bukti item PO yang cukup atau masih ambigu"),
        ("Jalur Level 1-3 terbukti", len(hierarchy_evidence), "Agregat jalur per vendor dari item PO asli"),
        ("Item PO memiliki Level 1-3", len(po) - len(hierarchy_unresolved), "Minimal satu hierarchy path terbukti"),
        ("Item PO belum memiliki Level 1-3", len(hierarchy_unresolved), "Tidak cocok atau ambigu terhadap hierarchy master"),
        ("Kelompok item belum memiliki Level 1-3", hierarchy_unresolved_group_count, "Dikelompokkan per SAP, deskripsi, dan alasan pada Audit"),
        ("Temuan review", len(reviews), "Seluruh tingkat severity"),
    ]
    for company in ("HK", "JO"):
        company_stats = po_stats[company]
        metrics.extend(
            [
                (f"Baris worksheet PO {company}", company_stats["raw_rows"], "Sebelum validasi Vendor"),
                (f"Baris PO {company} valid", company_stats["valid_vendor_rows"], "Vendor/SAP terisi"),
                (f"Baris total/footer PO {company}", company_stats.get("footer_rows", 0), "Direkonsiliasi terpisah; bukan item PO"),
                (f"Item PO {company} tanpa Vendor", company_stats["blank_vendor_rows"], "Item nyata tanpa SAP; masuk Audit HIGH"),
                (f"Vendor SAP unik PO {company}", po.loc[po["company"] == company, "sap"].nunique(), "Setelah deduplikasi persis"),
            ]
        )
    for source, records in sources.items():
        stats = source_stats[source]
        metrics.extend(
            [
                (f"Baris sumber {source}", stats["raw_rows"], "Seluruh baris data yang dibaca"),
                (f"Record sumber {source}", len(records), "Record bernama yang diperiksa dan dicocokkan"),
                (f"Baris {source} tanpa nama", stats["blank_name_rows"], "Tetap dicatat sebagai temuan Audit"),
            ]
        )
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
    po, sources, po_stats, source_stats = read_inputs(raw_dir)

    exact_duplicate_mask = po.duplicated(
        subset=["company", "po", "item_po", "sap", "description"], keep="first"
    )
    duplicate_po_rows_removed = int(exact_duplicate_mask.sum())
    po = po.loc[~exact_duplicate_mask].reset_index(drop=True)
    valid_input_rows = sum(
        int(stats["valid_vendor_rows"]) for stats in po_stats.values()
    )
    if len(po) + duplicate_po_rows_removed != valid_input_rows:
        raise RuntimeError(
            "PO row guard gagal; baris valid tidak seluruhnya terhitung setelah deduplikasi."
        )

    matches = match_sources_to_po(po, sources)
    classified, evidence, unresolved = classify_po(po, config_dir)
    hierarchy_by_vendor, hierarchy_evidence, hierarchy_unresolved, hierarchy_reviews = (
        classify_po_hierarchy(po, config_dir)
    )
    circle_rules = load_circle_rules(config_dir)
    rows, reviews = build_output_rows(
        po,
        classified,
        matches,
        settings,
        circle_rules=circle_rules,
        hierarchy_by_vendor=hierarchy_by_vendor,
    )
    reviews.extend(hierarchy_reviews)
    reviews.extend(_po_input_reviews(po_stats))
    reviews.extend(audit_vendor_sources(sources, source_stats))
    validate_output(rows, set(po["sap"].astype(str)))
    _enrich_evidence_with_circle(evidence, rows, circle_rules)
    unresolved_groups = _group_unresolved_items(unresolved)
    hierarchy_unresolved_groups = _group_hierarchy_unresolved(hierarchy_unresolved)

    summary = _summary_rows(
        po,
        sources,
        rows,
        reviews,
        evidence,
        unresolved,
        matches,
        duplicate_po_rows_removed,
        po_stats,
        source_stats,
        len(unresolved_groups),
        hierarchy_evidence,
        hierarchy_unresolved,
        len(hierarchy_unresolved_groups),
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
                "unresolved_rows": unresolved_groups,
                "hierarchy_evidence_rows": hierarchy_evidence,
                "hierarchy_unresolved_rows": hierarchy_unresolved_groups,
                "po_footer_rows": [
                    footer
                    for stats in po_stats.values()
                    for footer in stats.get("footer_reconciliations", [])
                ],
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
        "Rule ID",
        "Confidence Rule",
        "Dukungan Circle",
        "Sumber Final",
        "Contoh Deskripsi",
        "Rule Pattern",
    ]
    _write_csv(output_dir / "classification_evidence.csv", evidence, evidence_columns)
    unresolved_columns = [
        "Company",
        "PO",
        "Item PO",
        "NO SAP",
        "Nama Vendor",
        "Deskripsi",
        "Material",
        "Divisi",
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
