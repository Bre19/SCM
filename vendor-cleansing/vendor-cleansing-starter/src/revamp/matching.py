from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any

import pandas as pd

from .constants import SOURCE_PRECEDENCE
from .normalize import canonical_name, clean_text, normalize_name


@dataclass
class MatchResult:
    matched: dict[str, dict[str, list[dict[str, Any]]]]
    review_rows: list[dict[str, str]]
    method_counts: dict[str, dict[str, int]]
    outside_po_counts: dict[str, int]


def _set_map(items: list[tuple[str, str]]) -> dict[str, set[str]]:
    result: dict[str, set[str]] = defaultdict(set)
    for key, sap in items:
        if key and sap:
            result[key].add(sap)
    return dict(result)


def build_po_indexes(po: pd.DataFrame) -> tuple[set[str], dict[str, set[str]], dict[str, set[str]]]:
    po_saps = set(po["sap"].astype(str))
    exact = _set_map(
        [(normalize_name(row.name), row.sap) for row in po[["name", "sap"]].itertuples(index=False)]
    )
    canonical = _set_map(
        [(canonical_name(row.name), row.sap) for row in po[["name", "sap"]].itertuples(index=False)]
    )
    return po_saps, exact, canonical


def build_id_index(sources: dict[str, list[dict[str, Any]]]) -> dict[str, set[str]]:
    pairs: list[tuple[str, str]] = []
    for source in ("DRT", "DRT_LAMA", "DM", "DM_LAMA"):
        for record in sources[source]:
            if record["id_vendor"] and record["sap"]:
                pairs.append((record["id_vendor"], record["sap"]))
    return _set_map(pairs)


def _review(
    issue: str,
    severity: str,
    record: dict[str, Any],
    method: str,
    detail: str,
) -> dict[str, str]:
    return {
        "Severity": severity,
        "Issue": issue,
        "Source": record.get("source", ""),
        "Source Row": str(record.get("source_row", "")),
        "ID Vendor": record.get("id_vendor", ""),
        "NO SAP": record.get("sap", ""),
        "Nama Rekanan": record.get("name", ""),
        "Match Method": method,
        "Detail": detail,
    }


def match_sources_to_po(
    po: pd.DataFrame,
    sources: dict[str, list[dict[str, Any]]],
) -> MatchResult:
    po_saps, exact_names, canonical_names = build_po_indexes(po)
    id_index = build_id_index(sources)
    matched: dict[str, dict[str, list[dict[str, Any]]]] = defaultdict(
        lambda: defaultdict(list)
    )
    review_rows: list[dict[str, str]] = []
    method_counts: dict[str, Counter[str]] = defaultdict(Counter)
    outside_po_counts: Counter[str] = Counter()

    for source in SOURCE_PRECEDENCE:
        for record in sources[source]:
            target_sap = ""
            method = ""
            sap = record["sap"]
            internal_id = record["id_vendor"]

            if sap:
                if sap in po_saps:
                    target_sap, method = sap, "DIRECT_SAP"
                else:
                    outside_po_counts[source] += 1
                    method_counts[source]["OUTSIDE_PO"] += 1
                    review_rows.append(
                        _review(
                            "SOURCE_SAP_NOT_IN_PO",
                            "LOW",
                            record,
                            "OUTSIDE_PO",
                            "NO SAP pada file sumber tidak ditemukan pada PO HK maupun PO JO; record tetap dicatat di Audit.",
                        )
                    )
                    continue
            elif internal_id and internal_id in id_index:
                candidates = id_index[internal_id]
                if len(candidates) > 1:
                    review_rows.append(
                        _review(
                            "ID_TO_MULTIPLE_SAP",
                            "HIGH",
                            record,
                            "ID_AMBIGUOUS",
                            "ID terhubung ke beberapa SAP: " + ", ".join(sorted(candidates)),
                        )
                    )
                    method_counts[source]["ID_AMBIGUOUS"] += 1
                    continue
                linked_sap = next(iter(candidates))
                if linked_sap in po_saps:
                    target_sap, method = linked_sap, "ID_LINK"
                else:
                    outside_po_counts[source] += 1
                    method_counts[source]["ID_OUTSIDE_PO"] += 1
                    review_rows.append(
                        _review(
                            "SOURCE_ID_LINK_OUTSIDE_PO",
                            "LOW",
                            record,
                            "ID_OUTSIDE_PO",
                            f"ID Vendor terhubung ke SAP {linked_sap}, tetapi SAP tersebut tidak ditemukan pada PO HK maupun PO JO.",
                        )
                    )
                    continue

            if not target_sap:
                exact_key = normalize_name(record["name"])
                exact_candidates = exact_names.get(exact_key, set())
                if len(exact_candidates) == 1:
                    target_sap, method = next(iter(exact_candidates)), "EXACT_NAME"
                elif len(exact_candidates) > 1:
                    review_rows.append(
                        _review(
                            "AMBIGUOUS_EXACT_NAME",
                            "HIGH",
                            record,
                            "EXACT_NAME_AMBIGUOUS",
                            "Nama cocok ke beberapa SAP PO: "
                            + ", ".join(sorted(exact_candidates)),
                        )
                    )
                    method_counts[source]["EXACT_NAME_AMBIGUOUS"] += 1
                    continue
                else:
                    canonical_key = canonical_name(record["name"])
                    canonical_candidates = canonical_names.get(canonical_key, set())
                    if canonical_key and len(canonical_candidates) == 1:
                        target_sap, method = next(iter(canonical_candidates)), "CANONICAL_NAME"
                    elif len(canonical_candidates) > 1:
                        review_rows.append(
                            _review(
                                "AMBIGUOUS_CANONICAL_NAME",
                                "MEDIUM",
                                record,
                                "CANONICAL_NAME_AMBIGUOUS",
                                "Nama kanonik cocok ke beberapa SAP PO: "
                                + ", ".join(sorted(canonical_candidates)),
                            )
                        )
                        method_counts[source]["CANONICAL_NAME_AMBIGUOUS"] += 1
                        continue

            if target_sap:
                enriched = dict(record)
                enriched["matched_sap"] = target_sap
                enriched["match_method"] = method
                matched[target_sap][source].append(enriched)
                method_counts[source][method] += 1
            else:
                method_counts[source]["NO_PO_MATCH"] += 1
                review_rows.append(
                    _review(
                        "SOURCE_NO_PO_MATCH",
                        "LOW",
                        record,
                        "NO_PO_MATCH",
                        "Record sumber tidak dapat dicocokkan secara unik ke vendor PO melalui SAP, ID, nama exact, maupun nama kanonik.",
                    )
                )

    return MatchResult(
        matched={sap: dict(by_source) for sap, by_source in matched.items()},
        review_rows=review_rows,
        method_counts={source: dict(counts) for source, counts in method_counts.items()},
        outside_po_counts=dict(outside_po_counts),
    )


def choose_best(records: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not records:
        return None

    def score(record: dict[str, Any]) -> tuple[int, int, str, int]:
        useful_fields = (
            "id_vendor",
            "sap",
            "name",
            "npwp",
            "qualification",
            "coverage",
            "business_field",
            "circle",
        )
        completeness = sum(bool(clean_text(record.get(field))) for field in useful_fields)
        return (
            int(bool(clean_text(record.get("approval_date")))),
            completeness,
            clean_text(record.get("approval_date")) or clean_text(record.get("registration_date")),
            -int(record.get("source_row", 0)),
        )

    return max(records, key=score)
