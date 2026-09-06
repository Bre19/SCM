from __future__ import annotations

from collections import defaultdict
from typing import Any, Callable

from .constants import CURRENT_MASTER_SOURCES, LEGACY_MASTER_SOURCES
from .normalize import canonical_name, clean_text, normalize_name


MASTER_SOURCES = set(CURRENT_MASTER_SOURCES + LEGACY_MASTER_SOURCES)


def _review(
    issue: str,
    severity: str,
    record: dict[str, Any],
    detail: str,
    method: str = "SOURCE_VALIDATION",
) -> dict[str, str]:
    return {
        "Severity": severity,
        "Issue": issue,
        "Source": clean_text(record.get("source")),
        "Source Row": str(record.get("source_row", "")),
        "ID Vendor": clean_text(record.get("id_vendor")),
        "NO SAP": clean_text(record.get("sap")),
        "Nama Rekanan": clean_text(record.get("name")),
        "Match Method": method,
        "Detail": detail,
    }


def _duplicate_reviews(
    source: str,
    records: list[dict[str, Any]],
    field: str,
    issue: str,
    severity: str,
    key_fn: Callable[[Any], str] = clean_text,
) -> list[dict[str, str]]:
    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        key = key_fn(record.get(field, ""))
        if key:
            groups[key].append(record)

    reviews: list[dict[str, str]] = []
    for key, duplicates in sorted(groups.items()):
        if len(duplicates) < 2:
            continue
        example = duplicates[0]
        line_numbers = [str(record.get("source_row", "")) for record in duplicates]
        names = sorted(
            {clean_text(record.get("name")) for record in duplicates if clean_text(record.get("name"))}
        )
        review = _review(
            issue,
            severity,
            example,
            (
                f"{len(duplicates)} record {source} memiliki {field} yang sama ({key}). "
                f"Baris sumber: {', '.join(line_numbers)}"
                + (f". Nama: {' | '.join(names[:10])}" if names else "")
            ),
            "WITHIN_SOURCE_DUPLICATE",
        )
        review["Source Row"] = ", ".join(line_numbers)
        reviews.append(review)
    return reviews


def audit_vendor_sources(
    sources: dict[str, list[dict[str, Any]]],
    source_stats: dict[str, dict[str, Any]],
) -> list[dict[str, str]]:
    """Audit every parsed vendor-source row, including rows not connected to a PO."""
    reviews: list[dict[str, str]] = []

    for source, records in sources.items():
        for rejected in source_stats.get(source, {}).get("rejected_rows", []):
            reviews.append(
                _review(
                    "SOURCE_RECORD_MISSING_NAME",
                    "HIGH",
                    rejected,
                    "Baris sumber berisi data tetapi Nama Rekanan kosong; baris tidak dapat dicocokkan secara aman.",
                )
            )

        reviews.extend(
            _duplicate_reviews(
                source, records, "sap", "SOURCE_DUPLICATE_SAP", "HIGH"
            )
        )
        reviews.extend(
            _duplicate_reviews(
                source, records, "id_vendor", "SOURCE_DUPLICATE_ID", "HIGH"
            )
        )
        reviews.extend(
            _duplicate_reviews(
                source,
                records,
                "name",
                "SOURCE_DUPLICATE_NAME",
                "MEDIUM",
                key_fn=canonical_name,
            )
        )

        for record in records:
            if not clean_text(record.get("sap")):
                severity = "HIGH" if source in MASTER_SOURCES else "LOW"
                kind = "master" if source in MASTER_SOURCES else "calon/vendor belum terverifikasi"
                reviews.append(
                    _review(
                        "SOURCE_RECORD_MISSING_SAP",
                        severity,
                        record,
                        f"Record {kind} tidak mempunyai NO SAP/Ext Number.",
                    )
                )
            if not clean_text(record.get("id_vendor")):
                reviews.append(
                    _review(
                        "SOURCE_RECORD_MISSING_ID",
                        "MEDIUM" if source in MASTER_SOURCES else "LOW",
                        record,
                        "ID Vendor/Kode Identitas kosong pada file sumber.",
                    )
                )
            npwp = clean_text(record.get("npwp"))
            if not npwp:
                reviews.append(
                    _review(
                        "SOURCE_RECORD_MISSING_NPWP",
                        "LOW",
                        record,
                        "NPWP kosong pada file sumber.",
                    )
                )
            elif len(npwp) not in {15, 16}:
                reviews.append(
                    _review(
                        "SOURCE_RECORD_INVALID_NPWP",
                        "MEDIUM",
                        record,
                        f"Panjang NPWP {len(npwp)} digit; expected 15 atau 16 digit.",
                    )
                )

    by_sap = defaultdict(list)
    for records in sources.values():
        for record in records:
            if record['sap']:
                by_sap[record['sap']].append(record)
    for sap, records in by_sap.items():
        for field, issue in [('npwp', 'NPWP_CONFLICT'), ('id_vendor', 'ID_VENDOR_CONFLICT'), ('name', 'NAME_CONFLICT')]:
            values = {normalize_name(r[field]) if field == 'name' else clean_text(r[field]) for r in records if clean_text(r[field])}
            if len(values) > 1:
                for record in records:
                    reviews.append(_review(issue, 'HIGH', record,
                        f'SAP {sap} memiliki {field} berbeda: ' + ' | '.join(sorted(values)) + '. Record asli dipertahankan, tidak dipilih salah satu untuk mengganti record lain.'))
    return reviews
