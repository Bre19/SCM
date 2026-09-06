from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

from src.classifier_v2_1 import (
    ContextRule,
    classify_context,
    classify_primary_text,
    load_exclusion_rules,
    load_context_rules,
    load_primary_rules,
    load_vocabulary,
)

from .constants import EXCEL_CELL_TEXT_LIMIT
from .normalize import clean_text


def _bounded_lines(values: list[str], limit: int = EXCEL_CELL_TEXT_LIMIT) -> tuple[str, bool]:
    output: list[str] = []
    length = 0
    truncated = False
    for value in values:
        text = clean_text(value)
        if not text or text in output:
            continue
        addition = len(text) + (1 if output else 0)
        if length + addition > limit:
            truncated = True
            break
        output.append(text)
        length += addition
    return "\n".join(output), truncated


def load_circle_rules(config_dir: Path) -> list[ContextRule]:
    vocabulary = load_vocabulary(config_dir / "vocabulary_v2.csv")
    return load_context_rules(config_dir / "context_rules_v2.csv", vocabulary)


def map_circle_classifications(
    values: list[str],
    rules: list[ContextRule],
) -> list[str]:
    text = " | ".join(clean_text(value) for value in values if clean_text(value))
    return list(classify_context(text, "CIRCLE", rules)) if text else []


def classify_po(
    po: pd.DataFrame,
    config_dir: Path,
) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    vocabulary = load_vocabulary(config_dir / "vocabulary_v2.csv")
    rules = load_primary_rules(config_dir / "po_rules_v2.csv", vocabulary)
    exclusions = load_exclusion_rules(
        config_dir / "classification_exclusions_v2.csv", vocabulary
    )

    vendor_data: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "names": [],
            "po_sources": set(),
            "descriptions": [],
            "classifications": defaultdict(
                lambda: {
                    "po_keys": set(),
                    "item_keys": set(),
                    "examples": [],
                    "max_priority": 0,
                    "rule_ids": set(),
                    "rule_patterns": set(),
                    "confidence_levels": set(),
                }
            ),
        }
    )
    unresolved: list[dict[str, Any]] = []
    description_cache = {}

    for row in po.itertuples(index=False):
        target = vendor_data[row.sap]
        if row.name and row.name not in target["names"]:
            target["names"].append(row.name)
        target["po_sources"].add(row.company)
        if row.description and row.description not in target["descriptions"]:
            target["descriptions"].append(row.description)

        if row.description not in description_cache:
            description_cache[row.description] = classify_primary_text(row.description, rules, exclusions)
        matches = description_cache[row.description]
        po_key = f"{row.company}|{row.po}"
        item_key = f"{row.company}|{row.po}|{row.item_po}"
        if not matches:
            unresolved.append(
                {
                    "Company": row.company,
                    "PO": row.po,
                    "Item PO": row.item_po,
                    "NO SAP": row.sap,
                    "Nama Vendor": row.name,
                    "Deskripsi": row.description,
                    "Material": row.material,
                    "Divisi": row.division,
                    "Project": row.project,
                }
            )
            continue

        for label, rule in matches.items():
            stats = target["classifications"][label]
            stats["po_keys"].add(po_key)
            stats["item_keys"].add(item_key)
            stats["max_priority"] = max(stats["max_priority"], rule.priority)
            stats["rule_ids"].add(rule.rule_id)
            stats["rule_patterns"].add(rule.pattern_text)
            stats["confidence_levels"].add(rule.confidence)
            if row.description and row.description not in stats["examples"]:
                stats["examples"].append(row.description)

    evidence: list[dict[str, Any]] = []
    for sap, target in vendor_data.items():
        descriptions, truncated = _bounded_lines(target["descriptions"])
        target["item_text"] = descriptions
        target["item_text_truncated"] = truncated
        ranked: list[tuple[str, dict[str, Any]]] = sorted(
            target["classifications"].items(),
            key=lambda item: (
                -len(item[1]["po_keys"]),
                -len(item[1]["item_keys"]),
                -item[1]["max_priority"],
                item[0],
            ),
        )
        target["ordered_labels"] = [label for label, _ in ranked]
        target["final_classification"] = ", ".join(target["ordered_labels"])
        target["classification_count"] = len(ranked)
        for rank, (label, stats) in enumerate(ranked, start=1):
            evidence.append(
                {
                    "NO SAP": sap,
                    "Nama Vendor PO": target["names"][0] if target["names"] else "",
                    "Rank": rank,
                    "Klasifikasi": label,
                    "Jumlah PO Berbeda": len(stats["po_keys"]),
                    "Jumlah Item PO": len(stats["item_keys"]),
                    "Prioritas Rule": stats["max_priority"],
                    "Rule ID": ", ".join(sorted(stats["rule_ids"])),
                    "Confidence Rule": ", ".join(
                        level
                        for level in ("HIGH", "MEDIUM", "LOW")
                        if level in stats["confidence_levels"]
                    ),
                    "Rule Pattern": " | ".join(sorted(stats["rule_patterns"])),
                    "Contoh Deskripsi": " | ".join(stats["examples"][:5]),
                }
            )

    return dict(vendor_data), evidence, unresolved
