from __future__ import annotations

from pathlib import Path
from decimal import Decimal, InvalidOperation
from typing import Any

import pandas as pd

from .constants import REQUIRED_INPUT_FILES
from .normalize import clean_text, normalize_identifier, normalize_npwp


def validate_input_files(raw_dir: Path) -> dict[str, Path]:
    paths = {key: raw_dir / filename for key, filename in REQUIRED_INPUT_FILES.items()}
    missing = [str(path) for path in paths.values() if not path.is_file()]
    if missing:
        raise FileNotFoundError("Input wajib tidak ditemukan:\n- " + "\n- ".join(missing))
    return paths


def _flatten_column(column: object) -> str:
    if not isinstance(column, tuple):
        return clean_text(column)
    parts: list[str] = []
    for value in column:
        text = clean_text(value)
        if not text or text.lower().startswith("unnamed:"):
            continue
        if not parts or text.lower() != parts[-1].lower():
            parts.append(text)
    return " | ".join(parts)


def _read_html_or_excel(path: Path, multi_header: bool) -> pd.DataFrame:
    with path.open("rb") as handle:
        signature = handle.read(512).lstrip().lower()
    if signature.startswith(b"pk") or signature.startswith(bytes.fromhex("d0cf11e0")):
        frame = pd.read_excel(
            path,
            header=[0, 1] if multi_header else 0,
            dtype=str,
            keep_default_na=False,
        )
    elif b"<table" in signature or b"<style" in signature or b"<html" in signature:
        tables = pd.read_html(path, keep_default_na=False)
        if not tables:
            raise ValueError(f"Tidak ada tabel pada {path.name}")
        frame = tables[0]
    else:
        raise ValueError(f"Format file tidak dikenali: {path}")
    frame.columns = [_flatten_column(column) for column in frame.columns]
    return frame.fillna("")


def _column(frame: pd.DataFrame, *aliases: str, required: bool = False) -> str | None:
    normalized = {clean_text(column).upper(): column for column in frame.columns}
    for alias in aliases:
        match = normalized.get(alias.upper())
        if match is not None:
            return match
    if required:
        raise ValueError(
            f"Kolom wajib {aliases!r} tidak ditemukan. Kolom tersedia: {list(frame.columns)!r}"
        )
    return None


def _value(row: pd.Series, column: str | None) -> str:
    return clean_text(row[column]) if column is not None else ""


def _decimal(value: object) -> Decimal:
    text = clean_text(value).replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def read_vendor_source(
    path: Path, source: str
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    multi_header = source in {"DRT", "DCR", "DCM", "DM"}
    frame = _read_html_or_excel(path, multi_header=multi_header)
    mapping: dict[str, tuple[str, ...]] = {
        "id_vendor": ("ID VENDOR", "ID", "KODE IDENTITAS"),
        "sap": ("EXT NUMBER", "NO SAP"),
        "name": ("NAMA REKANAN", "NAMA MANDOR/PEORANGAN"),
        "npwp": ("NPWP",),
        "qualification": ("KUALIFIKASI",),
        "coverage": ("CAKUPAN WILAYAH",),
        "business_field": ("BIDANG USAHA",),
        "circle": ("KLASIFIKASI",),
        "status": ("STATUS",),
        "registration_date": ("TGL REGISTRASI", "TANGGAL REGISTRASI"),
        "approval_date": ("TGL APPROVE", "TANGGAL APPROVE"),
        "email": ("EMAIL",),
    }
    columns = {
        field: _column(frame, *aliases, required=(field == "name"))
        for field, aliases in mapping.items()
    }

    records: list[dict[str, Any]] = []
    rejected_rows: list[dict[str, Any]] = []
    first_data_row = 3 if multi_header else 2
    for ordinal, (_, row) in enumerate(frame.iterrows(), start=1):
        name = _value(row, columns["name"])
        source_row = first_data_row + ordinal - 1
        if not name:
            values = [clean_text(value) for value in row.tolist()]
            if any(values):
                rejected_rows.append(
                    {
                        "source": source,
                        "source_file": path.name,
                        "source_row": source_row,
                        "id_vendor": normalize_identifier(_value(row, columns["id_vendor"])),
                        "sap": normalize_identifier(_value(row, columns["sap"])),
                        "npwp": normalize_npwp(_value(row, columns["npwp"])),
                    }
                )
            continue
        record = {
            "source": source,
            "source_file": path.name,
            "source_row": source_row,
            "id_vendor": normalize_identifier(_value(row, columns["id_vendor"])),
            "sap": normalize_identifier(_value(row, columns["sap"])),
            "name": name,
            "npwp": normalize_npwp(_value(row, columns["npwp"])),
            "qualification": _value(row, columns["qualification"]),
            "coverage": _value(row, columns["coverage"]),
            "business_field": _value(row, columns["business_field"]),
            "circle": _value(row, columns["circle"]),
            "status": _value(row, columns["status"]),
            "registration_date": _value(row, columns["registration_date"]),
            "approval_date": _value(row, columns["approval_date"]),
            "email": _value(row, columns["email"]),
            "entity_type": "INDIVIDUAL" if source in {"DM", "DM_LAMA", "DCM"} else "COMPANY",
        }
        records.append(record)
    return records, {
        "raw_rows": len(frame),
        "parsed_records": len(records),
        "blank_name_rows": len(rejected_rows),
        "rejected_rows": rejected_rows,
    }


def read_po(path: Path, company: str) -> tuple[pd.DataFrame, dict[str, Any]]:
    frame = pd.read_excel(
        path,
        sheet_name="Data",
        header=2,
        dtype=str,
        keep_default_na=False,
    ).fillna("")
    required = ["PO", "Item.PO", "Vendor", "Nama Vendor", "Deskripsi"]
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError(f"{path.name} tidak memiliki kolom: {missing}")

    rows: list[dict[str, str]] = []
    rejected_rows: list[dict[str, str]] = []
    footer_rows: list[dict[str, str]] = []
    for ordinal, row in enumerate(frame.to_dict("records"), start=4):
        sap = normalize_identifier(row.get("Vendor"))
        if not sap:
            identity_fields_blank = all(
                not clean_text(row.get(column))
                for column in ("PO", "Item.PO", "Vendor", "Nama Vendor", "Material", "Deskripsi")
            )
            has_totals = bool(
                clean_text(row.get("Currency"))
                and (clean_text(row.get("Nilai PO")) or clean_text(row.get("Harga Satuan")))
            )
            if identity_fields_blank and has_totals:
                footer_rows.append(
                    {
                        "Company": company,
                        "Source File": path.name,
                        "Source Row": str(ordinal),
                        "Currency": clean_text(row.get("Currency")),
                        "Nilai PO Footer": clean_text(row.get("Nilai PO")),
                        "Harga Satuan Footer": clean_text(row.get("Harga Satuan")),
                    }
                )
                continue
            rejected_rows.append(
                {
                    "company": company,
                    "source_file": path.name,
                    "source_row": str(ordinal),
                    "po": normalize_identifier(row.get("PO")),
                    "item_po": normalize_identifier(row.get("Item.PO")),
                    "name": clean_text(row.get("Nama Vendor")),
                    "description": clean_text(row.get("Deskripsi"))
                    or clean_text(row.get("Material")),
                }
            )
            continue
        description = clean_text(row.get("Deskripsi")) or clean_text(row.get("Material"))
        rows.append(
            {
                "company": company,
                "source_file": path.name,
                "source_row": str(ordinal),
                "doc_date": clean_text(row.get("Doc.Date")),
                "po": normalize_identifier(row.get("PO")),
                "item_po": normalize_identifier(row.get("Item.PO")),
                "sap": sap,
                "name": clean_text(row.get("Nama Vendor")),
                "material": clean_text(row.get("Material")),
                "description": description,
                "division": clean_text(row.get("Nama Divisi")) or clean_text(row.get("Divisi")),
                "project": clean_text(row.get("Project/KP")),
                "po_value": clean_text(row.get("Nilai PO")),
                "unit_price": clean_text(row.get("Harga Satuan")),
                "currency": clean_text(row.get("Currency")),
            }
        )
    footer_reconciliations: list[dict[str, str]] = []
    for footer in footer_rows:
        currency = footer["Currency"]
        currency_rows = [row for row in rows if row["currency"] == currency]
        po_value_total = sum((_decimal(row["po_value"]) for row in currency_rows), Decimal("0"))
        unit_price_total = sum((_decimal(row["unit_price"]) for row in currency_rows), Decimal("0"))
        footer_po_value = _decimal(footer["Nilai PO Footer"])
        footer_unit_price = _decimal(footer["Harga Satuan Footer"])
        footer_reconciliations.append(
            {
                **footer,
                "Nilai PO Hitung Ulang": str(po_value_total),
                "Status Nilai PO": "SESUAI" if abs(po_value_total - footer_po_value) <= Decimal("0.01") else "TIDAK SESUAI",
                "Harga Satuan Hitung Ulang": str(unit_price_total),
                "Status Harga Satuan": "SESUAI" if abs(unit_price_total - footer_unit_price) <= Decimal("0.01") else "TIDAK SESUAI",
                "Keterangan": "Baris total/footer, bukan item PO dan bukan anomali Vendor/SAP.",
            }
        )
    return pd.DataFrame(rows), {
        "raw_rows": len(frame),
        "valid_vendor_rows": len(rows),
        "blank_vendor_rows": len(rejected_rows),
        "footer_rows": len(footer_rows),
        "footer_reconciliations": footer_reconciliations,
        "rejected_rows": rejected_rows,
    }


def read_inputs(
    raw_dir: Path,
) -> tuple[
    pd.DataFrame,
    dict[str, list[dict[str, Any]]],
    dict[str, dict[str, Any]],
    dict[str, dict[str, Any]],
]:
    paths = validate_input_files(raw_dir)
    po_hk, hk_stats = read_po(paths["PO_HK"], "HK")
    po_jo, jo_stats = read_po(paths["PO_JO"], "JO")
    po = pd.concat(
        [po_hk, po_jo],
        ignore_index=True,
    )
    sources: dict[str, list[dict[str, Any]]] = {}
    source_stats: dict[str, dict[str, Any]] = {}
    for source in ("DRT", "DRT_LAMA", "DM", "DM_LAMA", "DCR", "DCM", "DBCR"):
        records, stats = read_vendor_source(paths[source], source)
        sources[source] = records
        source_stats[source] = stats
    return po, sources, {"HK": hk_stats, "JO": jo_stats}, source_stats
