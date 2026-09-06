"""Shared workbook layout and audit grouping; export is streamed to bound memory."""
from pathlib import Path
from typing import Any

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

DATA_WIDTHS = [
    14, 14, 30, 18, 8, 15, 9, 11, 9, 9, 9, 9, 10, 23, 15, 18, 20, 20,
    30, 34, 55, 28, 32, 38, 18,
]
AUDIT_WIDTHS = [31, 34, 45, 18, 18, 18, 32, 25, 65, 24, 55, 55, 24]


def _safe_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _partition_review_rows(rows):
    partitions = {key: [] for key in
                  ("duplicates", "completeness", "matching", "classification", "other")}
    matching_issues = {
        "SOURCE_SAP_NOT_IN_PO", "SOURCE_ID_LINK_OUTSIDE_PO", "SOURCE_NO_PO_MATCH",
        "PO_VENDOR_NO_REGISTRY_MATCH", "CIRCLE_WITHOUT_PO", "NAME_LINK_REQUIRES_CONFIRMATION",
    }
    for row in rows:
        issue = str(row.get("Issue", ""))
        if issue.startswith("HIERARCHY_"):
            key = "classification"
        elif any(token in issue for token in ("DUPLICATE", "CONFLICT", "AMBIGUOUS")) or issue in {
            "ID_TO_MULTIPLE_SAP", "NAME_MULTIPLE_SAP", "PO_REPEATED_ITEM_KEY"
        }:
            key = "duplicates"
        elif issue in matching_issues:
            key = "matching"
        elif any(token in issue for token in ("MISSING", "INVALID", "WITHOUT_VENDOR")):
            key = "completeness"
        elif issue.startswith("PO_") or issue.startswith("CIRCLE_") or issue == "ITEM_TEXT_TRUNCATED":
            key = "classification"
        else:
            key = "other"
        partitions[key].append(row)
    return partitions


def build_workbook(bundle_path: Path, output_path: Path) -> Path:
    from .stream_workbook import export_streaming
    return export_streaming(bundle_path, output_path)
