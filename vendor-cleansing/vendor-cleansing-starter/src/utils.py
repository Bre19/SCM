from __future__ import annotations

import re
from typing import Iterable


def clean_string(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_text(value: object) -> str:
    return clean_string(value).upper()


def normalize_vendor_name(value: object) -> str:
    """Conservative normalized key; does not remove legal-entity words."""
    text = normalize_text(value)
    text = re.sub(r"[.,]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def canonical_vendor_name(value: object) -> str:
    """Secondary key for possible duplicates such as PT ABC vs PT. ABC."""
    text = normalize_vendor_name(value)
    text = re.sub(r"^(PT|CV|PD|UD)\s+", "", text)
    text = re.sub(r"\s+(TBK)$", "", text)
    return text.strip()


def normalize_npwp(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    # Keep letters if source contains non-standard placeholders; remove common separators only.
    return re.sub(r"[.\-\s]", "", text)


def is_all_upper_label(value: str) -> bool:
    letters = [ch for ch in value if ch.isalpha()]
    return bool(letters) and not any(ch.islower() for ch in letters)


def split_top_level_commas(value: str) -> list[str]:
    """Split category lists on commas, but keep commas inside parentheses."""
    output: list[str] = []
    buffer: list[str] = []
    depth = 0

    for ch in value:
        if ch == "(":
            depth += 1
            buffer.append(ch)
        elif ch == ")":
            depth = max(0, depth - 1)
            buffer.append(ch)
        elif ch == "," and depth == 0:
            token = "".join(buffer).strip(" ,")
            if token:
                output.append(token)
            buffer = []
        else:
            buffer.append(ch)

    token = "".join(buffer).strip(" ,")
    if token:
        output.append(token)

    return output


def join_flags(flags: Iterable[str]) -> str:
    return " | ".join(sorted(set(f for f in flags if f)))
