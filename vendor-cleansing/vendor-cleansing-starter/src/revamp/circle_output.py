"""Record-preserving HK Circle output and SAP-level PO accounting.

Output identity is a source record, not a unique SAP. Amount ownership is a
separate, deterministic accounting anchor; it never merges vendor identities.
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any

from .constants import CHECKMARK, CURRENT_MASTER_SOURCES, LEGACY_MASTER_SOURCES, LEVEL_COLUMNS, OUTPUT_COLUMNS, SOURCE_PRECEDENCE
from .hierarchy import format_vendor_hierarchy
from .matching import build_id_index
from .normalize import clean_text, normalize_name


def po_amount(value: Any) -> Decimal:
    """Invalid/missing financial input must never silently become zero."""
    text = clean_text(value)
    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Nilai PO bukan angka valid: {text!r}") from exc
    if not amount.is_finite():
        raise ValueError(f"Nilai PO bukan angka finite: {text!r}")
    return amount


def assemble_circle_output(po, sources, matches, classified, hierarchy, settings):
    # Import here to keep existing public helpers available without an import cycle.
    from .pipeline import _category, _classification_review, _po_labels
    from .classification import map_circle_classifications

    reviews = []
    totals = defaultdict(lambda: Decimal(0))
    subtotals = defaultdict(lambda: Decimal(0))
    po_counts = defaultdict(int)
    currencies = defaultdict(set)
    names_to_saps = defaultdict(set)
    for item in po.to_dict("records"):
        sap = item["sap"]
        try:
            amount = po_amount(item["po_value"])
        except ValueError as exc:
            raise ValueError(f"{item['source_file']} baris {item['source_row']}: {exc}") from exc
        totals[sap] += amount
        subtotals[(sap, item["company"])] += amount
        po_counts[sap] += 1
        currencies[sap].add(item.get("currency", "") or "KOSONG")
        name = normalize_name(item["name"])
        if name:
            names_to_saps[name].add(sap)
    records = [record for source in SOURCE_PRECEDENCE for record in sources[source]]
    id_index = build_id_index(sources)
    linked = {
        (record["source"], record["source_row"]): (sap, record["match_method"])
        for sap, by_source in matches.matched.items()
        for group in by_source.values() for record in group
    }
    for record in records:
        if record["sap"] and normalize_name(record["name"]):
            names_to_saps[normalize_name(record["name"])].add(record["sap"])
    conflicts_by_sap = defaultdict(list)
    for name, sap_group in names_to_saps.items():
        if len(sap_group) > 1:
            detail = f"{name}: " + ", ".join(sorted(sap_group))
            for sap in sap_group:
                conflicts_by_sap[sap].append(detail)

    rows, provenance = [], []
    membership = defaultdict(set)
    for record in records:
        sap = record["sap"]
        if sap:
            membership[sap].add(record["source"])
    for sap, by_source in matches.matched.items():
        for source, group in by_source.items():
            if any(r["match_method"] in {"DIRECT_SAP", "ID_LINK"} for r in group):
                membership[sap].add(source)

    circle_rules = settings.get("_circle_rules", [])

    def append_record(record, po_only=False):
        source, source_row = record["source"], record["source_row"]
        sap = record["sap"]
        match_sap, method = linked.get((source, source_row), (sap if sap in totals else "", "DIRECT_SAP" if sap in totals else "NO_PO_MATCH"))
        # A verified unique internal-ID link may fill SAP, but name-only matching
        # provides provisional classification evidence, never balance ownership.
        ids = id_index.get(record["id_vendor"], set())
        if not sap and len(ids) == 1:
            sap = next(iter(ids))
            match_sap = sap if sap in totals else ""
            method = "ID_LINK"
        flags = membership.get(sap, set()) | ({source} if not po_only else set())
        current = bool(flags & set(CURRENT_MASTER_SOURCES))
        legacy = bool(flags & set(LEGACY_MASTER_SOURCES))
        category = _category(current, legacy and not current, bool(flags & {"DCR", "DCM"}), "DBCR" in flags)
        info = classified.get(match_sap, {})
        levels = format_vendor_hierarchy(hierarchy.get(match_sap))
        row = {column: "" for column in OUTPUT_COLUMNS}
        row.update({
            "ID Vendor": record["id_vendor"], "NO SAP": sap,
            "Nama Rekanan": record["name"], "NPWP": record.get("npwp", ""),
            "PO": CHECKMARK if info else "",
            "Master Data Vendor SAP": CHECKMARK if current or legacy else "",
            "DRT": CHECKMARK if current else "", "DRT Lama": CHECKMARK if legacy and not current else "",
            "Inject": CHECKMARK if info and not (current or legacy) else "",
            "DCR": CHECKMARK if "DCR" in flags else "", "DCM": CHECKMARK if "DCM" in flags else "",
            "DBCR": CHECKMARK if "DBCR" in flags else "", "Kategori": category,
            "Perlakuan": settings["categories"][category]["treatment"],
            "Kualifikasi": record.get("qualification", ""), "Cakupan Wilayah": record.get("coverage", ""),
            "Bidang Usaha": record.get("business_field", ""),
            "Badan Usaha": "" if po_only else ("Perorangan/Mandor" if record.get("entity_type") == "INDIVIDUAL" else "Perusahaan"),
            "Klasifikasi Circle": record.get("circle", ""),
            "Klasifikasi Final": ", ".join(_po_labels(info)),
            "Item Pekerjaan Berdasarkan PO": info.get("item_text", ""),
            **dict(zip(LEVEL_COLUMNS, levels)),
        })
        rows.append(row)
        excel_row = len(rows) + 1
        provenance.append({
            "Baris Data Cleansing": excel_row, "Source": source, "Source File": record.get("source_file", source),
            "Source Row": source_row, "ID Vendor": record["id_vendor"], "NO SAP": sap,
            "SAP Sumber": record.get("sap_raw", record["sap"]), "SAP Bukti PO": match_sap, "Nama Rekanan": record["name"],
            "Match Method": method, "Baris Pemilik Saldo": "", "Status Saldo": "Tidak ada PO terhubung",
        })

        def issue(code, severity, detail):
            reviews.append({"Severity": severity, "Issue": code, "Source": source,
                "Source Row": str(source_row), "ID Vendor": record["id_vendor"], "NO SAP": sap,
                "Nama Rekanan": record["name"], "Match Method": method, "Detail": detail,
                "Baris Data Cleansing": str(excel_row)})

        if not record["sap"]:
            issue("SOURCE_RECORD_MISSING_SAP", "MEDIUM", "SAP belum tersedia pada sumber" + (f" (nilai asli: {record['sap_raw']})" if record.get('sap_raw') else '') + ". " + (f"Diisi dari ID Vendor unik ke SAP {sap}." if sap else "Tidak diisi dari kemiripan nama."))
        if match_sap and method in {"EXACT_NAME", "CANONICAL_NAME"}:
            issue("NAME_LINK_REQUIRES_CONFIRMATION", "MEDIUM", f"Bukti klasifikasi sementara dari nama unik ke SAP {match_sap}; verifikasi identitas. Saldo tidak dialokasikan lewat nama.")
        if not info:
            issue("CIRCLE_WITHOUT_PO", "LOW", "Record HK Circle tetap ditampilkan; tidak ada bukti PO terhubung. Klasifikasi Final tidak ditebak dari deklarasi Circle.")
        else:
            review = _classification_review(match_sap, record["name"], sorted(info["po_sources"]), _po_labels(info),
                [record["circle"]] if record.get("circle") else [], map_circle_classifications([record.get("circle", "")], circle_rules))
            # Agreements and PO-only evidence belong in evidence, not anomaly totals.
            if review["Issue"] not in {"PO_CIRCLE_AGREEMENT", "PO_CLASSIFICATION_WITHOUT_CIRCLE"}:
                review.update({"Source": source, "Source Row": str(source_row), "Baris Data Cleansing": str(excel_row), "NO SAP": sap})
                reviews.append(review)
            if info.get('item_text_truncated'):
                issue('ITEM_TEXT_TRUNCATED', 'MEDIUM', 'Gabungan deskripsi melebihi kapasitas satu sel Excel. Semua item tetap dihitung untuk saldo dan klasifikasi; lihat file PO asli berdasarkan SAP untuk daftar lengkap.')
        name_conflicts = conflicts_by_sap.get(sap, [])
        if name_conflicts:
            issue("NAME_MULTIPLE_SAP", "HIGH", "Nama sama ditemukan pada SAP berbeda. " + " | ".join(name_conflicts) + ". Record dan saldo setiap SAP tetap terpisah; belum tentu vendor yang sama.")
        if po_only:
            issue("PO_VENDOR_NO_REGISTRY_MATCH", "MEDIUM", "SAP PO belum terhubung melalui SAP/ID ke HK Circle. Tetap ditampilkan agar pekerjaan dan nilai PO tidak hilang.")

    for record in records:
        append_record(record)
    represented = {row["NO SAP"] for row in rows if row["NO SAP"]}
    for sap in sorted(set(totals) - represented):
        append_record({"source": "PO_ONLY", "source_row": "", "source_file": "PO HK / PO JO",
            "sap": sap, "id_vendor": "", "name": classified[sap]["names"][0] if classified[sap]["names"] else ""}, po_only=True)

    anchors = {}
    ledger = []
    for row, origin in zip(rows, provenance):
        sap = row["NO SAP"]
        if sap not in totals:
            continue
        if sap not in anchors:
            anchors[sap] = origin["Baris Data Cleansing"]
            row["Saldo Hutang"] = float(totals[sap])
            ledger.append({"NO SAP": sap, "Nama Rekanan": row["Nama Rekanan"],
                "Baris Data Cleansing": anchors[sap], "Jumlah Item PO": po_counts[sap],
                "Nilai PO HK (IDR)": float(subtotals[(sap, "HK")]), "Nilai PO JO (IDR)": float(subtotals[(sap, "JO")]),
                "Saldo Hutang (IDR)": float(totals[sap]), "Currency Sumber": "; ".join(sorted(currencies[sap])),
                "Dasar": "Jumlah Nilai PO setiap baris; IDR sesuai konfirmasi pengguna. Tanpa konversi kurs."})
            origin["Status Saldo"] = "Dicatat sekali pada baris ini"
        else:
            origin["Status Saldo"] = "SAP berulang; saldo tidak diulang"
        origin["Baris Pemilik Saldo"] = anchors[sap]
    for sap, labels in currencies.items():
        if labels != {"IDR"}:
            reviews.append({"Severity": "MEDIUM", "Issue": "CURRENCY_LABEL_OVERRIDE", "Source": "PO HK / PO JO",
                "Source Row": "", "NO SAP": sap, "Nama Rekanan": "", "ID Vendor": "",
                "Match Method": "USER_CONFIRMED_IDR", "Baris Data Cleansing": str(anchors[sap]),
                "Detail": f"Label sumber: {', '.join(sorted(labels))}. Nilai PO diperlakukan sebagai IDR sesuai konfirmasi pengguna, tanpa konversi."})
    if len(rows) != len(records) + len(set(totals) - represented):
        raise RuntimeError("Record HK Circle tidak seluruhnya terwakili")
    if set(anchors) != set(totals):
        raise RuntimeError("Ada SAP PO tanpa pemilik saldo")
    exported_total = sum((Decimal(str(row["Saldo Hutang"])) for row in rows if row["Saldo Hutang"] != ""), Decimal(0))
    if abs(exported_total - sum(totals.values(), Decimal(0))) > Decimal("0.01"):
        raise RuntimeError("Total saldo output tidak sama dengan total Nilai PO")
    return rows, reviews, provenance, ledger


def attach_output_references(reviews, provenance):
    """Resolve each finding by source + row first, including missing SAP records."""
    source_index = defaultdict(list)
    sap_index = defaultdict(list)
    for row in provenance:
        source_index[(row["Source"], str(row["Source Row"]))].append(row["Baris Data Cleansing"])
        if row["NO SAP"]:
            sap_index[row["NO SAP"]].append(row["Baris Data Cleansing"])
    for review in reviews:
        if review.get("Baris Data Cleansing"):
            continue
        locations = []
        for source_row in str(review.get("Source Row", "")).split(","):
            locations.extend(source_index.get((review.get("Source"), source_row.strip()), []))
        if not locations and review.get("Source") not in SOURCE_PRECEDENCE:
            locations = sap_index.get(review.get("NO SAP", ""), [])
        review["Baris Data Cleansing"] = ", ".join(map(str, sorted(set(locations))))
    unique = {}
    for review in reviews:
        key = tuple(str(review.get(field, "")) for field in ("Issue", "Source", "Source Row", "NO SAP", "Detail"))
        unique.setdefault(key, review)
    return list(unique.values())
