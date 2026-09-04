from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


TRUE_VALUES = {
    "TRUE",
    "1",
    "YES",
}


@dataclass(frozen=True)
class PrimaryRule:
    rule_id: str
    priority: int
    confidence: str
    classification: str
    pattern_text: str
    exclude_pattern_text: str
    pattern: re.Pattern[str]
    exclude_pattern: re.Pattern[str] | None


@dataclass(frozen=True)
class ContextRule:
    source: str
    classification: str
    pattern_text: str
    pattern: re.Pattern[str]
    fallback_allowed: bool


@dataclass(frozen=True)
class ExclusionRule:
    classification: str
    pattern_text: str
    reason: str
    pattern: re.Pattern[str]


def normalize_text(
    value: object,
) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = text.upper().strip()

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text


def normalize_id(
    value: object,
) -> str:
    text = normalize_text(
        value
    )

    return re.sub(
        r"\.0$",
        "",
        text,
    )


def split_work_items(
    value: object,
) -> list[str]:
    if value is None:
        return []

    raw = str(value)

    if not raw.strip():
        return []

    raw = re.sub(
        r"<br\s*/?>",
        "\n",
        raw,
        flags=re.IGNORECASE,
    )

    parts = re.split(
        r"[\r\n]+",
        raw,
    )

    cleaned: list[str] = []

    for part in parts:
        text = part.strip()

        text = re.sub(
            r"^[\-\u2022•]+\s*",
            "",
            text,
        )

        if text:
            cleaned.append(
                text
            )

    return cleaned


def load_vocabulary(
    path: Path,
) -> dict[str, str]:
    df = pd.read_csv(
        path,
        dtype=str,
        keep_default_na=False,
    )

    required = {
        "classification",
        "status",
        "enabled",
    }

    missing = (
        required
        - set(df.columns)
    )

    if missing:
        raise ValueError(
            "Vocabulary V2 missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .str.upper()
        .isin(TRUE_VALUES)
    )

    df = df.loc[
        enabled_mask
    ].copy()

    result: dict[str, str] = {}

    for _, row in df.iterrows():
        classification = (
            row["classification"]
            .strip()
        )

        status = (
            row["status"]
            .strip()
            .upper()
        )

        if classification:
            result[
                classification
            ] = status

    return result


def load_primary_rules(
    path: Path,
    vocabulary: dict[str, str],
) -> list[PrimaryRule]:
    df = pd.read_csv(
        path,
        dtype=str,
        keep_default_na=False,
    )

    required = {
        "priority",
        "classification",
        "pattern",
        "exclude_pattern",
        "enabled",
    }

    missing = (
        required
        - set(df.columns)
    )

    if missing:
        raise ValueError(
            "PO rules V2 missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .str.upper()
        .isin(TRUE_VALUES)
    )

    df = df.loc[
        enabled_mask
    ].copy()

    rules: list[PrimaryRule] = []

    for position, (_, row) in enumerate(df.iterrows(), start=1):
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in vocabulary
        ):
            raise ValueError(
                "Classification in po_rules_v2.csv "
                "does not exist or is disabled in "
                "vocabulary_v2.csv: "
                f"{classification}"
            )

        pattern_text = (
            row["pattern"]
            .strip()
        )

        exclude_pattern_text = (
            row["exclude_pattern"]
            .strip()
        )

        rule_id = (
            row.get("rule_id", "").strip()
            or f"PO-{position:03d}"
        )
        confidence = (
            row.get("confidence", "").strip().upper()
            or ("HIGH" if int(row["priority"]) >= 120 else "MEDIUM")
        )
        if confidence not in {"HIGH", "MEDIUM", "LOW"}:
            raise ValueError(
                f"Confidence rule {rule_id} harus HIGH, MEDIUM, atau LOW: {confidence}"
            )

        rules.append(
            PrimaryRule(
                rule_id=rule_id,
                priority=int(
                    row["priority"]
                ),
                confidence=confidence,
                classification=(
                    classification
                ),
                pattern_text=(
                    pattern_text
                ),
                exclude_pattern_text=(
                    exclude_pattern_text
                ),
                pattern=re.compile(
                    pattern_text,
                    flags=re.IGNORECASE,
                ),
                exclude_pattern=(
                    re.compile(
                        exclude_pattern_text,
                        flags=re.IGNORECASE,
                    )
                    if exclude_pattern_text
                    else None
                ),
            )
        )

    rules.sort(
        key=lambda rule:
            rule.priority,
        reverse=True,
    )

    return rules


def load_context_rules(
    path: Path,
    vocabulary: dict[str, str],
) -> list[ContextRule]:
    df = pd.read_csv(
        path,
        dtype=str,
        keep_default_na=False,
    )

    required = {
        "source",
        "classification",
        "pattern",
        "fallback_allowed",
        "enabled",
    }

    missing = (
        required
        - set(df.columns)
    )

    if missing:
        raise ValueError(
            "Context rules V2 missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .str.upper()
        .isin(TRUE_VALUES)
    )

    df = df.loc[
        enabled_mask
    ].copy()

    rules: list[ContextRule] = []

    for _, row in df.iterrows():
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in vocabulary
        ):
            raise ValueError(
                "Classification in context_rules_v2.csv "
                "does not exist or is disabled in "
                "vocabulary_v2.csv: "
                f"{classification}"
            )

        pattern_text = (
            row["pattern"]
            .strip()
        )

        fallback_allowed = (
            str(
                row["fallback_allowed"]
            )
            .strip()
            .upper()
            in TRUE_VALUES
        )

        rules.append(
            ContextRule(
                source=(
                    row["source"]
                    .strip()
                    .upper()
                ),
                classification=(
                    classification
                ),
                pattern_text=(
                    pattern_text
                ),
                pattern=re.compile(
                    pattern_text,
                    flags=re.IGNORECASE,
                ),
                fallback_allowed=(
                    fallback_allowed
                ),
            )
        )

    return rules


def load_exclusion_rules(
    path: Path,
    vocabulary: dict[str, str],
) -> dict[
    str,
    list[ExclusionRule],
]:
    df = pd.read_csv(
        path,
        dtype=str,
        keep_default_na=False,
    )

    required = {
        "classification",
        "pattern",
        "reason",
        "enabled",
    }

    missing = (
        required
        - set(df.columns)
    )

    if missing:
        raise ValueError(
            "classification_exclusions_v2.csv "
            "missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .str.upper()
        .isin(TRUE_VALUES)
    )

    df = df.loc[
        enabled_mask
    ].copy()

    result: dict[
        str,
        list[ExclusionRule],
    ] = {}

    for _, row in df.iterrows():
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in vocabulary
        ):
            raise ValueError(
                "Classification in exclusions "
                "does not exist in vocabulary V2: "
                f"{classification}"
            )

        rule = ExclusionRule(
            classification=classification,
            pattern_text=(
                row["pattern"]
                .strip()
            ),
            reason=(
                row["reason"]
                .strip()
            ),
            pattern=re.compile(
                row["pattern"].strip(),
                flags=re.IGNORECASE,
            ),
        )

        result.setdefault(
            classification,
            [],
        ).append(
            rule
        )

    return result


def get_exclusion_reason(
    classification: str,
    text: str,
    exclusions: dict[
        str,
        list[ExclusionRule],
    ],
) -> str:
    for rule in exclusions.get(
        classification,
        [],
    ):
        if rule.pattern.search(
            text
        ):
            return rule.reason

    return ""


def classify_primary_text(
    value: object,
    rules: list[PrimaryRule],
    exclusions: dict[
        str,
        list[ExclusionRule],
    ],
) -> dict[
    str,
    PrimaryRule,
]:
    text = normalize_text(
        value
    )

    if not text:
        return {}

    matches: dict[
        str,
        PrimaryRule,
    ] = {}

    for rule in rules:
        if not rule.pattern.search(
            text
        ):
            continue

        if (
            rule.exclude_pattern
            is not None
            and rule.exclude_pattern.search(
                text
            )
        ):
            continue

        exclusion_reason = (
            get_exclusion_reason(
                rule.classification,
                text,
                exclusions,
            )
        )

        if exclusion_reason:
            continue

        current = matches.get(
            rule.classification
        )

        if (
            current is None
            or rule.priority
            > current.priority
        ):
            matches[
                rule.classification
            ] = rule

    return matches


def classify_context(
    value: object,
    source: str,
    rules: list[ContextRule],
) -> dict[
    str,
    ContextRule,
]:
    text = normalize_text(
        value
    )

    source = (
        source
        .strip()
        .upper()
    )

    if not text:
        return {}

    matches: dict[
        str,
        ContextRule,
    ] = {}

    for rule in rules:
        if (
            rule.source
            != source
        ):
            continue

        if rule.pattern.search(
            text
        ):
            matches[
                rule.classification
            ] = rule

    return matches
