from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

import pandas as pd

from .config import (
    L1_ALIASES,
    NPWP_PLACEHOLDERS,
    SEVERITY,
)

from .utils import (
    canonical_vendor_name,
    join_flags,
    normalize_npwp,
    normalize_text,
)


@dataclass(frozen=True)
class Finding:
    excel_row: int
    code: str
    severity: str
    detail: str


def _duplicate_rows(
    series: pd.Series,
    ignore_blank: bool = True,
) -> set[int]:

    groups: dict[str, list[int]] = defaultdict(list)

    for excel_row, value in enumerate(
        series.tolist(),
        start=2,
    ):
        key = normalize_text(value)

        if ignore_blank and not key:
            continue

        groups[key].append(excel_row)

    return {
        row
        for rows in groups.values()
        if len(rows) > 1
        for row in rows
    }


def audit_cleansing(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:

    work = df.copy()

    work["__excel_row"] = range(
        2,
        len(work) + 2,
    )

    work["__sap_key"] = (
        work["NO SAP"]
        .map(normalize_text)
    )

    work["__id_vendor_key"] = (
        work["ID Vendor"]
        .map(normalize_text)
    )

    # Exact name normalization:
    # uppercase + whitespace normalization,
    # punctuation/legal entity belum dihapus.
    work["__name_key"] = (
        work["Nama Rekanan"]
        .map(normalize_text)
    )

    # Secondary duplicate candidate:
    # PT. ABC / PT ABC / ABC, dll.
    work["__canonical_name_key"] = (
        work["Nama Rekanan"]
        .map(canonical_vendor_name)
    )

    work["__npwp_key"] = (
        work["NPWP"]
        .map(normalize_npwp)
    )

    findings: list[Finding] = []

    def add(
        row: int,
        code: str,
        severity: str,
        detail: str,
    ) -> None:

        findings.append(
            Finding(
                excel_row=row,
                code=code,
                severity=severity,
                detail=detail,
            )
        )

    # =========================================================
    # 1. HARD DATA INTEGRITY
    # =========================================================

    source_columns = list(df.columns)

    exact_duplicate_mask = (
        work[source_columns]
        .duplicated(keep=False)
    )

    for _, row in work.loc[
        exact_duplicate_mask
    ].iterrows():

        add(
            int(row["__excel_row"]),
            "EXACT_DUPLICATE_ROW",
            "CRITICAL",
            (
                "Seluruh isi row sama persis "
                "dengan row lain di Data Cleansing"
            ),
        )

    # NO SAP kosong
    for _, row in work.loc[
        work["__sap_key"].eq("")
    ].iterrows():

        add(
            int(row["__excel_row"]),
            "MISSING_NO_SAP",
            "CRITICAL",
            "NO SAP kosong",
        )

    # NO SAP duplicate
    for excel_row in _duplicate_rows(
        work["NO SAP"]
    ):
        add(
            excel_row,
            "DUPLICATE_NO_SAP",
            "CRITICAL",
            "NO SAP muncul lebih dari sekali",
        )

    # ID Vendor duplicate
    for excel_row in _duplicate_rows(
        work["ID Vendor"]
    ):
        add(
            excel_row,
            "DUPLICATE_ID_VENDOR",
            "CRITICAL",
            "ID Vendor muncul lebih dari sekali",
        )

    # =========================================================
    # 2. VENDOR DUPLICATE CANDIDATES
    # =========================================================

    name_counts = (
        work.loc[
            work["__name_key"].ne(""),
            "__name_key",
        ]
        .value_counts()
    )

    duplicate_names = set(
        name_counts[
            name_counts.gt(1)
        ].index
    )

    for _, row in work.loc[
        work["__name_key"].isin(
            duplicate_names
        )
    ].iterrows():

        count = int(
            name_counts[
                row["__name_key"]
            ]
        )

        add(
            int(row["__excel_row"]),
            "DUPLICATE_VENDOR_NAME",
            "MEDIUM",
            (
                f"Nama Rekanan muncul "
                f"{count} kali: "
                f"{row['Nama Rekanan']}"
            ),
        )

    # =========================================================
    # 3. NPWP
    # =========================================================

    valid_npwp_mask = (
        work["__npwp_key"].ne("")
        &
        ~work["__npwp_key"].isin(
            NPWP_PLACEHOLDERS
        )
    )

    # Strong duplicate candidate:
    # nama sama DAN NPWP sama.
    same_identity_counts = (
        work.loc[
            valid_npwp_mask
            & work["__name_key"].ne("")
        ]
        .groupby([
            "__name_key",
            "__npwp_key",
        ])
        .size()
    )

    repeated_identity = set(
        same_identity_counts[
            same_identity_counts.gt(1)
        ].index
    )

    for _, row in work.loc[
        valid_npwp_mask
    ].iterrows():

        key = (
            row["__name_key"],
            row["__npwp_key"],
        )

        if key in repeated_identity:
            add(
                int(row["__excel_row"]),
                "SAME_NAME_AND_NPWP",
                "HIGH",
                (
                    "Nama Rekanan dan NPWP "
                    "sama dengan row lain"
                ),
            )

    # NPWP reused oleh beberapa row.
    npwp_counts = (
        work.loc[
            valid_npwp_mask,
            "__npwp_key",
        ]
        .value_counts()
    )

    repeated_npwp = set(
        npwp_counts[
            npwp_counts.gt(1)
        ].index
    )

    for _, row in work.loc[
        work["__npwp_key"].isin(
            repeated_npwp
        )
    ].iterrows():

        count = int(
            npwp_counts[
                row["__npwp_key"]
            ]
        )

        add(
            int(row["__excel_row"]),
            "REUSED_NPWP",
            "HIGH",
            (
                f"NPWP yang sama dipakai "
                f"pada {count} row: "
                f"{row['NPWP']}"
            ),
        )

    # =========================================================
    # 4. CANONICAL NAME CANDIDATES
    # =========================================================

    canonical_groups = (
        work.loc[
            work[
                "__canonical_name_key"
            ].ne("")
        ]
        .groupby(
            "__canonical_name_key",
            sort=False,
        )
    )

    for canonical_key, group in canonical_groups:

        if len(group) <= 1:
            continue

        displayed_names = set(
            group["__name_key"]
        )

        # Exact duplicate sudah ditangkap sebelumnya.
        if len(displayed_names) <= 1:
            continue

        for _, row in group.iterrows():

            add(
                int(row["__excel_row"]),
                "POSSIBLE_DUPLICATE_CANONICAL_NAME",
                "MEDIUM",
                (
                    "Nama canonical sama tetapi "
                    "penulisan berbeda: "
                    f"{canonical_key}"
                ),
            )

    # =========================================================
    # 5. CLASSIFICATION SOURCE ANOMALIES
    # =========================================================

    item_col = (
        "Item Pekerjaan Berdasarkan PO"
    )

    l1_col = (
        "Kelompok Klasifikasi\n"
        "Level 1 (Klasifikasi PO)"
    )

    l2_col = (
        "Kelompok Klasifikasi\n"
        "Level 2 (Klasifikasi PO)"
    )

    l3_col = (
        "Kelompok Klasifikasi\n"
        "Level 3 (Klasifikasi PO)"
    )

    for _, row in work.iterrows():

        excel_row = int(
            row["__excel_row"]
        )

        l1 = str(
            row[l1_col] or ""
        )

        for old, new in L1_ALIASES.items():

            if old in l1:

                add(
                    excel_row,
                    "NONCANONICAL_L1_NAME",
                    "LOW",
                    f"{old} -> {new}",
                )

        has_po_items = bool(
            str(
                row[item_col] or ""
            ).strip()
        )

        if (
            has_po_items
            and not str(
                row[l2_col] or ""
            ).strip()
        ):
            add(
                excel_row,
                "MISSING_PO_LEVEL_2",
                "LOW",
                (
                    "Item pekerjaan PO ada "
                    "tetapi Level 2 kosong"
                ),
            )

        if (
            has_po_items
            and not str(
                row[l3_col] or ""
            ).strip()
        ):
            add(
                excel_row,
                "MISSING_PO_LEVEL_3",
                "LOW",
                (
                    "Item pekerjaan PO ada "
                    "tetapi Level 3 kosong"
                ),
            )

    # =========================================================
    # RESULT
    # =========================================================

    finding_df = pd.DataFrame([
        finding.__dict__
        for finding in findings
    ])

    if finding_df.empty:

        finding_df = pd.DataFrame(
            columns=[
                "excel_row",
                "code",
                "severity",
                "detail",
            ]
        )

        summary = pd.DataFrame(
            columns=[
                "excel_row",
                "highest_severity",
                "flags",
                "details",
            ]
        )

        return finding_df, summary

    finding_df["severity_rank"] = (
        finding_df["severity"]
        .map(SEVERITY)
    )

    summary = (
        finding_df
        .sort_values(
            [
                "excel_row",
                "severity_rank",
            ],
            ascending=[
                True,
                False,
            ],
        )
        .groupby(
            "excel_row",
            as_index=False,
        )
        .agg(
            highest_severity=(
                "severity",
                "first",
            ),
            flags=(
                "code",
                join_flags,
            ),
            details=(
                "detail",
                lambda values:
                    " || ".join(
                        dict.fromkeys(values)
                    ),
            ),
        )
    )

    return (
        finding_df.drop(
            columns="severity_rank"
        ),
        summary,
    )