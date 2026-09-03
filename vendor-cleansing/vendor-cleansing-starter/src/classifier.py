from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


@dataclass(frozen=True)
class ClassificationRule:
    priority: int
    classification: str
    pattern_text: str
    exclude_pattern_text: str
    pattern: re.Pattern[str]
    exclude_pattern: re.Pattern[str] | None


def normalize_text(value: object) -> str:
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


def normalize_id(value: object) -> str:
    text = normalize_text(
        value
    )

    # Defensive handling jika Excel numeric ID
    # berubah menjadi "2020001234.0".
    text = re.sub(
        r"\.0$",
        "",
        text,
    )

    return text


def load_enabled_vocabulary(
    vocabulary_file: Path,
) -> set[str]:
    df = pd.read_csv(
        vocabulary_file,
        dtype=str,
        keep_default_na=False,
    )

    required = {
        "classification",
        "enabled",
    }

    missing = (
        required
        - set(df.columns)
    )

    if missing:
        raise ValueError(
            "Vocabulary missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .astype(str)
        .str.upper()
        .isin(
            {
                "TRUE",
                "1",
                "YES",
            }
        )
    )

    return set(
        df.loc[
            enabled_mask,
            "classification",
        ]
        .astype(str)
        .str.strip()
    )


def load_classification_rules(
    rules_file: Path,
    enabled_vocabulary: set[str],
) -> list[ClassificationRule]:
    df = pd.read_csv(
        rules_file,
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
            "Classification rules missing columns: "
            f"{sorted(missing)}"
        )

    enabled_mask = (
        df["enabled"]
        .astype(str)
        .str.upper()
        .isin(
            {
                "TRUE",
                "1",
                "YES",
            }
        )
    )

    df = df.loc[
        enabled_mask
    ].copy()

    rules: list[
        ClassificationRule
    ] = []

    for _, row in df.iterrows():
        classification = (
            row["classification"]
            .strip()
        )

        if (
            classification
            not in enabled_vocabulary
        ):
            raise ValueError(
                "Rule uses classification "
                "not enabled in vocabulary: "
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

        compiled_pattern = (
            re.compile(
                pattern_text,
                flags=re.IGNORECASE,
            )
        )

        compiled_exclude = (
            re.compile(
                exclude_text,
                flags=re.IGNORECASE,
            )
            if exclude_text
            else None
        )

        rules.append(
            ClassificationRule(
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
                pattern=(
                    compiled_pattern
                ),
                exclude_pattern=(
                    compiled_exclude
                ),
            )
        )

    rules.sort(
        key=lambda rule:
            rule.priority,
        reverse=True,
    )

    return rules


def classify_text(
    value: object,
    rules: list[
        ClassificationRule
    ],
) -> list[
    tuple[
        str,
        ClassificationRule,
    ]
]:
    text = normalize_text(
        value
    )

    if not text:
        return []

    matches: dict[
        str,
        ClassificationRule,
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

        # Satu klasifikasi cukup disimpan sekali
        # untuk satu description.
        if (
            rule.classification
            not in matches
        ):
            matches[
                rule.classification
            ] = rule

    return [
        (
            classification,
            rule,
        )
        for (
            classification,
            rule,
        ) in matches.items()
    ]