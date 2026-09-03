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
class RegexRule:
    priority: int
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

    enabled = (
        df["enabled"]
        .str.upper()
        .isin(
            TRUE_VALUES
        )
    )

    df = df.loc[
        enabled
    ].copy()

    result: dict[
        str,
        str
    ] = {}

    for _, row in (
        df.iterrows()
    ):
        classification = (
            row["classification"]
            .strip()
        )

        status = (
            row["status"]
            .strip()
            .upper()
        )

        result[
            classification
        ] = status

    return result


def load_po_rules(
    path: Path,
    vocabulary: dict[str, str],
) -> list[RegexRule]:
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

    enabled = (
        df["enabled"]
        .str.upper()
        .isin(
            TRUE_VALUES
        )
    )

    df = df.loc[
        enabled
    ].copy()

    rules: list[
        RegexRule
    ] = []

    for _, row in (
        df.iterrows()
    ):
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in vocabulary
        ):
            raise ValueError(
                "PO rule classification "
                "not found in vocabulary V2: "
                f"{classification}"
            )

        pattern_text = (
            row["pattern"]
            .strip()
        )

        exclude_text = (
            row["exclude_pattern"]
            .strip()
        )

        rules.append(
            RegexRule(
                priority=int(
                    row["priority"]
                ),
                classification=(
                    classification
                ),
                pattern_text=(
                    pattern_text
                ),
                exclude_pattern_text=(
                    exclude_text
                ),
                pattern=re.compile(
                    pattern_text,
                    flags=re.IGNORECASE,
                ),
                exclude_pattern=(
                    re.compile(
                        exclude_text,
                        flags=re.IGNORECASE,
                    )
                    if exclude_text
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

    enabled = (
        df["enabled"]
        .str.upper()
        .isin(
            TRUE_VALUES
        )
    )

    df = df.loc[
        enabled
    ].copy()

    rules: list[
        ContextRule
    ] = []

    for _, row in (
        df.iterrows()
    ):
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in vocabulary
        ):
            raise ValueError(
                "Context classification "
                "not found in vocabulary V2: "
                f"{classification}"
            )

        pattern_text = (
            row["pattern"]
            .strip()
        )

        fallback_allowed = (
            str(
                row[
                    "fallback_allowed"
                ]
            )
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


def classify_po_description(
    value: object,
    rules: list[RegexRule],
) -> list[RegexRule]:
    text = normalize_text(
        value
    )

    if not text:
        return []

    matches: dict[
        str,
        RegexRule,
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

    return list(
        matches.values()
    )


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

    source = source.upper()

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