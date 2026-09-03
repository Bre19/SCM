from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Iterable


EMPTY_VALUES = {"", "-", "N/A", "NA", "NAN", "NONE", "NULL", "\\N"}
LEGAL_PREFIXES = {
    "PT",
    "CV",
    "UD",
    "PD",
    "KSO",
    "JO",
    "KOPERASI",
    "YAYASAN",
}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return "" if text.upper() in EMPTY_VALUES else text


def normalize_identifier(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"\s+", "", text)


def normalize_npwp(value: object) -> str:
    text = clean_text(value).upper()
    if not text:
        return ""
    digits = re.sub(r"\D", "", text)
    if not digits or set(digits) == {"0"}:
        return ""
    return digits


def normalize_name(value: object) -> str:
    text = clean_text(value).upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def canonical_name(value: object) -> str:
    tokens = normalize_name(value).split()
    while tokens and tokens[0] in LEGAL_PREFIXES:
        tokens.pop(0)
    while tokens and tokens[-1] in {"TBK", "PERSERO"}:
        tokens.pop()
    return " ".join(tokens)


def first_nonempty(values: Iterable[object]) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""


def distinct_nonempty(values: Iterable[object]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = clean_text(value)
        key = text.upper()
        if text and key not in seen:
            seen.add(key)
            result.append(text)
    return result
