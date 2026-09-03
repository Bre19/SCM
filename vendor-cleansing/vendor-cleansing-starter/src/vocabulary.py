from __future__ import annotations

from collections import Counter

import pandas as pd

from .config import VOCABULARY_OVERRIDES
from .utils import (
    is_all_upper_label,
    split_top_level_commas,
)


def extract_uppercase_legacy_vocabulary(
    cleansing_df: pd.DataFrame,
    final_column: str = "Klasifikasi Final",
) -> pd.DataFrame:
    """
    Extract legacy final vocabulary.

    Rules:
    - hanya membaca first line Klasifikasi Final;
    - baris-baris evidence PO setelah first line diabaikan;
    - hanya vocabulary full-uppercase;
    - multiple classification dipisahkan dengan koma,
      kecuali koma di dalam tanda kurung;
    - explicit override tetap dicatat untuk audit.
    """

    counter: Counter[str] = Counter()

    for raw_value in cleansing_df[final_column].fillna("").astype(str):
        value = raw_value.strip()

        if not value:
            continue

        first_line = value.splitlines()[0].strip()

        if not is_all_upper_label(first_line):
            continue

        labels = split_top_level_commas(first_line)

        for label in labels:
            label = label.strip()

            if label and is_all_upper_label(label):
                counter[label] += 1

    rows: list[dict[str, object]] = []

    for label, count in sorted(counter.items()):
        override = VOCABULARY_OVERRIDES.get(label, {})

        rows.append({
            "classification": label,
            "existing_count": count,
            "enabled": override.get("enabled", True),
            "mapped_to": override.get("mapped_to", ""),
            "note": override.get("note", ""),
        })

    return pd.DataFrame(rows)