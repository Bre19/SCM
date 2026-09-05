from __future__ import annotations

import json
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .normalize import clean_text, normalize_name


@dataclass(frozen=True)
class HierarchyPath:
    level1: str
    code: str
    level2: str
    level3: str


@dataclass(frozen=True)
class TermRule:
    path: HierarchyPath
    term: str
    normalized_term: str
    tokens: tuple[str, ...]


INSTALL_CUES = {
    "PEKERJAAN", "PEK", "PEMASANGAN", "PASANG", "INSTALASI", "INSTALLATION",
    "INSTALL", "ERECTION", "EREKSI", "APLIKASI", "PELAKSANAAN",
    "FABRIKASI", "FABRICATION", "CONSTRUCTION", "KONSTRUKSI",
}
SUPPLY_CUES = {
    "PENGADAAN", "SUPPLY", "PEMBELIAN", "BELANJA", "MATERIAL", "BARANG",
    "PRODUK", "KOMPONEN",
}
RENT_CUES = {"SEWA", "RENTAL", "RENT", "PENYEWAAN"}
REPAIR_CUES = {
    "REPAIR", "PERBAIKAN", "PEMELIHARAAN", "MAINTENANCE", "SERVICE",
    "SERVIS", "OVERHAUL",
}
CONSULTING_CUES = {
    "KONSULTAN", "KONSULTANSI", "CONSULTING", "ADVISORY", "KAJIAN",
    "STUDI", "STUDY", "DESIGN", "DESAIN", "PERENCANAAN", "SUPERVISI",
    "SUPERVISION", "ASSESSMENT", "DUE DILIGENCE",
}
SPARE_CUES = {"SPARE PART", "SUKU CADANG", "KOMPONEN PENGGANTI", "PARTS"}
UNSAFE_TERMS = {
    "ALAT", "PLANT", "MATERIAL", "BARANG", "KOMPONEN", "PART", "PARTS",
    "SERVICE", "SERVICES", "SYSTEM", "SISTEM", "INSTALLATION", "WORKS",
    "EQUIPMENT", "PROJECT", "VEHICLE", "SUPPORT", "OPERATION", "GENERAL",
    "ACCESS", "HANDLING", "MANAGEMENT", "ORGANIZATION", "COST", "RISK",
    "CUT", "FILL", "DATA", "ASSET", "ERECTION", "CLEANING", "VIDEO",
    "STORAGE", "DISTRIBUTION", "FILTER", "TANK", "WORKSHOP", "ASSESSMENT",
}


# Aliases supplement the exact Level-3 wording from the approved master. They
# remain evidence terms applied directly to PO descriptions, never vendor names
# or Circle values.
ALIASES: dict[tuple[str, str], tuple[str, ...]] = {
    ("SUP-01", "Beton Siap Pakai (Ready-Mix Concrete)"): ("READY MIX", "READYMIX", "BETON READY MIX"),
    ("SUP-01", "Mortar & Dry Mix"): ("DRY MIX", "MORTAR"),
    ("SUP-02", "Batu Pecah/Agregat"): ("BATU SPLIT", "ABU BATU", "AGREGAT"),
    ("SUP-02", "Base Course/Subbase"): ("BASE COURSE", "SUB BASE", "SUBBASE", "LPA", "LPB"),
    ("SUP-03", "Besi Beton Polos"): ("BESI BETON", "REBAR", "REINFORCEMENT BAR"),
    ("SUP-03", "Welded Wire Mesh"): ("WIREMESH", "WIRE MESH"),
    ("SUP-03", "Tie Wire"): ("KAWAT BENDRAT", "BENDRAT", "KAWAT BETON"),
    ("SUP-04", "Profil Baja"): ("BAJA PROFIL", "H BEAM", "I BEAM", "WF BEAM"),
    ("SUP-04", "Plate/Sheet"): ("STEEL PLATE", "PLAT BAJA", "PLATE"),
    ("SUP-04", "Structural Pipe/Tube"): ("PIPA BAJA", "STEEL PIPE", "STEEL TUBE"),
    ("SUP-05", "Produk Pracetak/Prategang Lainnya"): ("PRECAST", "PRACETAK", "PRESTRESS"),
    ("SUP-06", "Aspal/Bitumen"): ("ASPAL", "BITUMEN"),
    ("SUP-06", "Hotmix"): ("HOT MIX", "AC WC", "AC BC", "LASTON"),
    ("SUP-07", "Bata Ringan/AAC"): ("BATA RINGAN", "AAC"),
    ("SUP-08", "Keramik/Tiles"): ("KERAMIK", "TILE", "TILES"),
    ("SUP-08", "Flooring"): ("LANTAI",),
    ("SUP-08", "Pintu & Jendela"): ("PINTU", "JENDELA", "DOOR", "WINDOW"),
    ("SUP-09", "Epoxy/Resin"): ("EPOXY", "RESIN"),
    ("SUP-10", "Pump Permanen"): ("POMPA PERMANEN", "PUMP PERMANEN"),
    ("SUP-11", "Panel/Switchgear"): ("PANEL LISTRIK", "ELECTRICAL PANEL", "SWITCHGEAR"),
    ("SUP-11", "Transformer"): ("TRAFO", "TRANSFORMER"),
    ("SUP-11", "Instrument/Sensor"): ("INSTRUMENT", "SENSOR"),
    ("SUP-11", "Aksesori Elektrikal"): ("ELECTRICAL ANCILLARIES", "AKSESORI ELEKTRIKAL"),
    ("SUP-12", "Sanitary Ware"): ("SANITAIR", "SANITARY"),
    ("SUP-13", "PVD"): ("PREFABRICATED VERTICAL DRAIN", "VERTICAL DRAIN", "PVD"),
    ("SUP-14", "Material Bekisting"): ("BEKISTING", "FORMWORK"),
    ("SUP-15", "Material Marka"): ("CAT MARKA", "THERMOPLASTIC MARKING", "MATERIAL MARKA"),
    ("SUP-16", "APD"): ("ALAT PELINDUNG DIRI", "APD", "PPE"),
    ("SUP-17", "Spare Part Plant/Alat"): ("SPARE PART", "SUKU CADANG"),
    ("SUP-17", "Tire"): ("BAN", "TIRE", "TYRE"),
    ("SUP-17", "Battery"): ("ACCU", "AKI", "BATTERY"),
    ("SUP-18", "Cutting/Grinding Consumables"): ("CUTTING WHEEL", "GRINDING WHEEL", "BATU GERINDA"),
    ("SUP-18", "Workshop Consumables"): ("CONSUMABLE MATERIAL", "CONSUMABLE"),
    ("SUP-19", "BBM"): ("BAHAN BAKAR", "SOLAR", "DIESEL", "BBM"),
    ("SUP-19", "Oli"): ("PELUMAS", "LUBRICANT", "OLI"),
    ("SUP-20", "Laptop/Desktop"): ("LAPTOP", "DESKTOP", "KOMPUTER", "PC"),
    ("SUP-20", "Network Device"): ("ROUTER", "SWITCH NETWORK", "ACCESS POINT"),
    ("SUP-20", "Storage"): ("STORAGE DEVICE", "DATA STORAGE HARDWARE"),
    ("SUP-20", "CCTV/Access Control Hardware"): ("CCTV", "ACCESS CONTROL"),
    ("SUP-21", "ATK"): ("ALAT TULIS KANTOR", "ATK", "STATIONERY"),
    ("SUB-01", "Land Clearing"): ("PEMBERSIHAN LAHAN", "LAND CLEARING"),
    ("SUB-01", "Site Preparation"): ("PEKERJAAN PERSIAPAN", "JASA PERSIAPAN"),
    ("SUB-01", "Demolition"): ("PEMBONGKARAN", "DEMOLITION"),
    ("SUB-01", "Excavation"): ("GALIAN", "PENGGALIAN", "EXCAVATION"),
    ("SUB-01", "Cut & Fill"): ("CUT AND FILL", "CUT FILL", "TIMBUNAN DAN GALIAN"),
    ("SUB-02", "Driven Pile"): ("PEMANCANGAN", "DRIVEN PILE"),
    ("SUB-02", "Bored Pile"): ("BORED PILE", "BOR PILE", "PONDASI BOR"),
    ("SUB-02", "Micro Pile"): ("MICROPILE", "MICRO PILE"),
    ("SUB-02", "Ground Improvement"): ("PERBAIKAN TANAH", "GROUND IMPROVEMENT"),
    ("SUB-03", "Pembesian"): ("PEMBESIAN", "REBAR INSTALLATION"),
    ("SUB-03", "Beton Cast-in-Situ"): ("PENGECORAN", "CAST IN SITU"),
    ("SUB-03", "Struktur Beton"): ("STRUKTUR BETON", "BETON STRUKTUR", "CONCRETE STRUCTURE"),
    ("SUB-04", "Erection Precast"): ("ERECTION PRECAST", "EREKSI PRECAST"),
    ("SUB-04", "Post-Tensioning"): ("POST TENSION", "POST TENSIONING"),
    ("SUB-05", "Fabrication & Erection"): ("FABRIKASI DAN EREKSI", "FABRICATION AND ERECTION"),
    ("SUB-05", "Steel Structure"): ("STRUKTUR BAJA", "STEEL STRUCTURE"),
    ("SUB-05", "Welding"): ("PENGELASAN", "WELDING"),
    ("SUB-06", "Asphalt Paving"): ("PENGASPALAN", "ASPHALT PAVING"),
    ("SUB-06", "Rigid Pavement"): ("PERKERASAN KAKU", "RIGID PAVEMENT"),
    ("SUB-06", "Road Marking Installation"): ("PEMASANGAN MARKA", "MARKA JALAN"),
    ("SUB-07", "Drainase"): ("DRAINASE", "DRAINAGE"),
    ("SUB-07", "Gorong-Gorong"): ("GORONG GORONG", "CULVERT"),
    ("SUB-08", "Mechanical"): ("MEKANIKAL", "MECHANICAL"),
    ("SUB-08", "Electrical"): ("ELEKTRIKAL", "ELECTRICAL", "LISTRIK"),
    ("SUB-08", "Plumbing"): ("PLUMBING", "PERPIPAAN"),
    ("SUB-08", "Fire Fighting"): ("FIRE FIGHTING", "FIRE PROTECTION", "HYDRANT", "SPRINKLER"),
    ("SUB-09", "Masonry"): ("PASANGAN DINDING", "PASANG DINDING", "PEKERJAAN DINDING", "MASONRY"),
    ("SUB-09", "Flooring"): ("LANTAI", "PASANG PLINT", "PEKERJAAN PLINT"),
    ("SUB-09", "Plastering"): ("PLESTER", "PLESTERAN", "ACIAN", "ACI", "PLASTERING"),
    ("SUB-09", "Painting"): ("PENGECATAN", "PAINTING"),
    ("SUB-09", "Landscape Construction"): ("PEKERJAAN LANDSCAPE", "PEKERJAAN LANSEKAP"),
    ("SUB-10", "Waterproofing Application"): ("APLIKASI WATERPROOFING", "PEKERJAAN WATERPROOFING"),
    ("SUB-10", "Structural Repair"): ("PERBAIKAN STRUKTUR", "STRUCTURAL REPAIR"),
    ("SUB-10", "Injection/Grouting"): ("INJEKSI GROUT", "INJECTION GROUTING"),
    ("SUB-10", "Scaffolding Erection/Dismantling"): ("PEMASANGAN SCAFFOLDING", "PEMASANGAN PERANCAH", "BONGKAR PASANG SCAFFOLDING"),
    ("ALT-01", "Excavator"): ("EXCAVATOR", "EXCAVATOR"),
    ("ALT-01", "Bulldozer"): ("BULLDOZER", "DOZER"),
    ("ALT-01", "Wheel Loader"): ("WHEEL LOADER", "LOADER"),
    ("ALT-01", "Motor Grader"): ("MOTOR GRADER", "GRADER"),
    ("ALT-02", "Dump Truck"): ("DUMP TRUCK", "DUMPTRUCK"),
    ("ALT-02", "Lowbed"): ("LOWBED", "LOW BED"),
    ("ALT-02", "Water Tanker"): ("WATER TANKER", "WATER TANK"),
    ("ALT-03", "Mobile Crane"): ("MOBILE CRANE", "CRANE"),
    ("ALT-04", "Vibro Roller"): ("VIBRO ROLLER", "VIBRATORY ROLLER"),
    ("ALT-04", "Plate Compactor"): ("PLATE COMPACTOR", "STAMPER"),
    ("ALT-06", "Bored Pile Rig"): ("BORED PILE RIG", "BORING RIG"),
    ("ALT-07", "Concrete Pump"): ("CONCRETE PUMP", "POMPA BETON"),
    ("ALT-07", "Truck Mixer"): ("TRUCK MIXER", "MIXER TRUCK"),
    ("ALT-08", "Batching Plant"): ("BATCHING PLANT", "BP PLANT"),
    ("ALT-08", "Asphalt Mixing Plant"): ("ASPHALT MIXING PLANT", "AMP"),
    ("ALT-09", "Genset"): ("GENSET", "GENERATOR SET"),
    ("ALT-10", "Scaffolding System"): ("SCAFFOLDING", "PERANCAH"),
    ("ALT-10", "Reusable Formwork System"): ("FORMWORK SYSTEM", "BEKISTING SISTEM"),
    ("ALT-12", "Pickup"): ("PICKUP", "PICK UP"),
    ("KON-01", "Detailed Engineering"): ("DETAIL ENGINEERING DESIGN", "DED"),
    ("KON-01", "BIM/Engineering Design"): ("BIM", "ENGINEERING DESIGN", "DESAIN ENGINEERING"),
    ("KON-02", "Project Management"): ("MANAJEMEN PROYEK", "PROJECT MANAGEMENT"),
    ("KON-02", "Project Control"): ("PROJECT CONTROL", "PENGENDALIAN PROYEK"),
    ("KON-03", "Construction Supervision"): ("PENGAWASAN KONSTRUKSI", "SUPERVISI KONSTRUKSI"),
    ("KON-04", "Topographic Study"): ("SURVEY TOPOGRAFI", "SURVEI TOPOGRAFI", "PEMETAAN TOPOGRAFI"),
    ("KON-04", "Geotechnical Study"): ("SOIL INVESTIGATION", "PENYELIDIKAN TANAH", "KAJIAN GEOTEKNIK"),
    ("KON-05", "Business Process"): ("PROSES BISNIS", "BUSINESS PROCESS"),
    ("KON-06", "Digital Transformation"): ("TRANSFORMASI DIGITAL", "DIGITAL TRANSFORMATION"),
    ("KON-06", "Cybersecurity Advisory"): ("CYBERSECURITY", "KEAMANAN INFORMASI"),
    ("KON-07", "Tax Advisory"): ("KONSULTASI PAJAK", "TAX ADVISORY"),
    ("KON-08", "Legal Advisory"): ("KONSULTASI HUKUM", "LEGAL ADVISORY"),
    ("KON-09", "Environmental Study"): ("KAJIAN LINGKUNGAN", "STUDI LINGKUNGAN", "AMDAL"),
    ("KON-10", "Financial Audit"): ("AUDIT KEUANGAN", "FINANCIAL AUDIT"),
    ("JAS-01", "Hauling Service"): ("JASA ANGKUT", "JASA ANGKUTAN", "HAULING", "LANGSIR"),
    ("JAS-01", "Courier"): ("KURIR", "COURIER"),
    ("JAS-01", "Warehousing"): ("PERGUDANGAN", "WAREHOUSE", "WAREHOUSING"),
    ("JAS-02", "Preventive Maintenance"): ("PREVENTIVE MAINTENANCE", "PEMELIHARAAN PREVENTIF", "PEMELIHARAAN"),
    ("JAS-02", "Corrective Maintenance"): ("CORRECTIVE MAINTENANCE", "PEMELIHARAAN KOREKTIF"),
    ("JAS-02", "Repair"): ("PERBAIKAN", "REPAIR"),
    ("JAS-03", "Laboratory Testing"): ("PENGUJIAN LABORATORIUM", "LABORATORY TESTING", "UJI LAB"),
    ("JAS-03", "Material Testing"): ("PENGUJIAN MATERIAL", "MATERIAL TESTING", "PENGUJIAN PIT", "PIT TEST"),
    ("JAS-03", "Calibration"): ("KALIBRASI", "CALIBRATION"),
    ("JAS-04", "Cleaning"): ("CLEANING SERVICE", "JASA KEBERSIHAN"),
    ("JAS-04", "Catering"): ("CATERING", "KATERING", "JASA BOGA"),
    ("JAS-05", "Security Guard"): ("SATPAM", "SECURITY GUARD", "JASA PENGAMANAN"),
    ("JAS-06", "Operator/Technician Manpower"): ("TENAGA KERJA", "MANPOWER", "TENAGA TEKNISI", "TENAGA OPERATOR"),
    ("JAS-06", "Recruitment Service"): ("REKRUTMEN", "RECRUITMENT"),
    ("JAS-07", "Software/SaaS Subscription"): ("SOFTWARE", "SAAS", "LISENSI APLIKASI", "SUBSCRIPTION"),
    ("JAS-07", "Cloud Service"): ("CLOUD SERVICE", "CLOUD COMPUTING"),
    ("JAS-07", "IT Support"): ("IT SUPPORT", "DUKUNGAN IT"),
    ("JAS-08", "Training"): ("PELATIHAN", "TRAINING"),
    ("JAS-08", "Professional Certification"): ("SERTIFIKASI PROFESI", "PROFESSIONAL CERTIFICATION"),
    ("JAS-09", "Printing"): ("PERCETAKAN", "CETAK", "PRINTING"),
    ("JAS-09", "Documentation"): ("DOKUMENTASI", "DOCUMENTATION"),
    ("JAS-09", "Event Organizer"): ("EVENT ORGANIZER", "PENYELENGGARAAN EVENT"),
    ("JAS-10", "Ticketing"): ("TIKET", "TICKETING"),
    ("JAS-10", "Accommodation"): ("AKOMODASI", "ACCOMMODATION"),
    ("JAS-11", "Waste Transportation"): ("PENGANGKUTAN LIMBAH", "WASTE TRANSPORTATION"),
    ("JAS-11", "Waste Treatment"): ("PENGOLAHAN LIMBAH", "WASTE TREATMENT"),
    ("JAS-12", "Insurance"): ("ASURANSI", "INSURANCE"),
}


def load_hierarchy(
    config_dir: Path,
) -> tuple[list[HierarchyPath], dict[str, list[TermRule]]]:
    path = config_dir / "classification_hierarchy.json"
    payload = json.loads(path.read_text(encoding="utf-8"))
    paths: list[HierarchyPath] = []
    rules_by_first_token: dict[str, list[TermRule]] = defaultdict(list)
    seen_codes: set[str] = set()
    for entry in payload.get("levels", []):
        code = clean_text(entry.get("code"))
        if not code or code in seen_codes:
            raise ValueError(f"Kode Level 2 kosong atau duplikat pada {path.name}: {code!r}")
        seen_codes.add(code)
        for level3 in entry.get("level3", []):
            hierarchy_path = HierarchyPath(
                clean_text(entry.get("level1")),
                code,
                clean_text(entry.get("level2")),
                clean_text(level3),
            )
            paths.append(hierarchy_path)
            terms = {hierarchy_path.level3}
            terms.update(ALIASES.get((code, hierarchy_path.level3), ()))
            expanded: set[str] = set()
            for term in terms:
                expanded.update(_term_variants(term))
            for term in expanded:
                normalized = normalize_name(term)
                if normalized and normalized not in UNSAFE_TERMS:
                    tokens = tuple(normalized.split())
                    rules_by_first_token[tokens[0]].append(
                        TermRule(hierarchy_path, term, normalized, tokens)
                    )
    if seen_codes and len(seen_codes) != 65:
        raise ValueError(f"Master hierarchy harus memiliki 65 kode Level 2; diperoleh {len(seen_codes)}")
    if len(paths) != 389:
        raise ValueError(
            f"Master hierarchy harus memiliki 389 cakupan Level 3; diperoleh {len(paths)}"
        )
    return paths, dict(rules_by_first_token)


def _term_variants(term: str) -> set[str]:
    values = {clean_text(term)}
    parenthetical = re.search(r"\(([^)]+)\)", term)
    if parenthetical:
        values.add(parenthetical.group(1))
        values.add(clean_text(term[: parenthetical.start()]))
    for separator in ("/", " & "):
        for value in list(values):
            if separator in value:
                values.update(clean_text(part) for part in value.split(separator))
    return {value for value in values if value}


def _contains(text: str, term: str) -> bool:
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(term)}(?![A-Z0-9])", text))


def _has_any(text: str, cues: set[str]) -> bool:
    return any(_contains(text, normalize_name(cue)) for cue in cues)


def _lookup(
    paths: list[HierarchyPath], code: str, level3: str
) -> HierarchyPath:
    for path in paths:
        if path.code == code and path.level3 == level3:
            return path
    raise KeyError(f"Hierarchy path tidak ditemukan: {code} / {level3}")


def _replace_for_boundary(
    text: str,
    candidates: list[tuple[HierarchyPath, str]],
    paths: list[HierarchyPath],
) -> tuple[list[tuple[HierarchyPath, str]], list[str]]:
    notes: list[str] = []
    install = _has_any(text, INSTALL_CUES)
    supply = _has_any(text, SUPPLY_CUES)
    rent = _has_any(text, RENT_CUES)
    repair = _has_any(text, REPAIR_CUES)
    consulting = _has_any(text, CONSULTING_CUES)
    spare = _has_any(text, SPARE_CUES)

    def forced(code: str, level3: str, evidence: str, note: str):
        notes.append(note)
        return [(_lookup(paths, code, level3), evidence)], notes

    if _has_any(text, {"SCAFFOLDING", "PERANCAH", "FORMWORK", "BEKISTING"}):
        if install:
            return forced(
                "SUB-10", "Scaffolding Erection/Dismantling", "SCAFFOLDING/FORMWORK + INSTALASI",
                "Boundary Scaffolding/Formwork: pemasangan atau bongkar-pasang diklasifikasikan sebagai paket pekerjaan Subkontraktor.",
            )
        if rent or _has_any(text, {"SYSTEM", "SISTEM", "REUSABLE"}):
            return forced(
                "ALT-10", "Scaffolding System", "SCAFFOLDING/FORMWORK + SEWA/SISTEM",
                "Boundary Scaffolding/Formwork: sistem reusable atau sewa diklasifikasikan sebagai Alat.",
            )
        if supply:
            return forced(
                "SUP-14", "Material Bekisting", "SCAFFOLDING/FORMWORK + SUPPLY/MATERIAL",
                "Boundary Scaffolding/Formwork: material yang dibeli diklasifikasikan sebagai Supplier.",
            )
        return [], ["AMBIGUOUS: scaffolding/formwork tidak menyebut material, sistem/sewa, atau pemasangan."]

    if _has_any(text, {"SOFTWARE", "SAAS", "CLOUD", "HOSTING", "APLIKASI", "IT SUPPORT", "MANAGED SERVICE"}):
        if consulting:
            return forced(
                "KON-06", "System Design", "IT + ADVISORY/DESIGN",
                "Boundary IT: output desain/advisory diklasifikasikan sebagai Jasa Konsultansi.",
            )
        return forced(
            "JAS-07", "System Implementation" if install else "IT Support",
            "IT SOFTWARE/SERVICE",
            "Boundary IT: software, implementasi, managed service, atau support diklasifikasikan sebagai Jasa Lainnya.",
        )

    if _has_any(text, {"LAPTOP", "DESKTOP", "SERVER", "ROUTER", "PRINTER", "SCANNER", "CCTV"}) and not repair:
        hardware_candidates = [item for item in candidates if item[0].code == "SUP-20"]
        if hardware_candidates:
            candidates = hardware_candidates
            notes.append("Boundary IT: perangkat keras diklasifikasikan sebagai Supplier.")

    mep_candidates = [item for item in candidates if item[0].code == "SUB-08"]
    if mep_candidates and not install:
        material_candidates = [
            item for item in candidates if item[0].code in {"SUP-10", "SUP-11", "SUP-12"}
        ]
        candidates = [item for item in candidates if item[0].code != "SUB-08"]
        if material_candidates:
            notes.append(
                "Boundary Supplier vs Subkontraktor: MEP tanpa bukti instalasi dipertahankan sebagai material/komponen Supplier."
            )
        else:
            notes.append(
                "AMBIGUOUS: istilah MEP tidak menyebut instalasi maupun material/komponen yang cukup spesifik."
            )

    if _has_any(text, {"RAMBU", "GUARDRAIL", "ROAD STUD", "DELINEATOR", "BARRIER", "MARKA"}):
        if install:
            level3 = "Road Marking Installation" if _has_any(text, {"MARKA"}) else "Road Furniture Installation"
            return forced(
                "SUB-06", level3, "ROAD SAFETY + INSTALASI",
                "Boundary Traffic/Road Safety: pemasangan diklasifikasikan sebagai Subkontraktor.",
            )
        if _has_any(text, {"TEMPORARY", "SEMENTARA", "SAFETY", "KESELAMATAN"}):
            return forced(
                "SUP-16", "Temporary Barrier" if _has_any(text, {"BARRIER"}) else "Rambu Keselamatan Proyek",
                "ROAD SAFETY + TEMPORARY",
                "Boundary Traffic/Road Safety: perlengkapan sementara proyek diklasifikasikan sebagai Supplier keselamatan.",
            )

    equipment_candidates = [item for item in candidates if item[0].level1 == "Alat"]
    if equipment_candidates and spare:
        return forced(
            "SUP-17", "Spare Part Plant/Alat", "EQUIPMENT + SPARE PART",
            "Boundary Supplier vs Alat: suku cadang/komponen diklasifikasikan sebagai Supplier.",
        )
    if equipment_candidates and repair:
        return forced(
            "JAS-02", "Repair", "EQUIPMENT + REPAIR/MAINTENANCE",
            "Boundary Alat vs Jasa Lainnya: pemeliharaan/perbaikan alat diklasifikasikan sebagai Jasa Lainnya.",
        )

    supplier_install_map = {
        "SUP-03": ("SUB-03", "Pembesian"),
        "SUP-04": ("SUB-05", "Steel Structure"),
        "SUP-05": ("SUB-04", "Instalasi Elemen Pracetak"),
        "SUP-06": ("SUB-06", "Asphalt Paving"),
        "SUP-07": ("SUB-09", "Masonry"),
        "SUP-08": ("SUB-09", "Architectural, Landscape & Finishing Works"),
        "SUP-09": ("SUB-10", "Waterproofing Application"),
        "SUP-10": ("SUB-08", "Mechanical"),
        "SUP-11": ("SUB-08", "Electrical"),
        "SUP-12": ("SUB-08", "Plumbing"),
        "SUP-14": ("SUB-03", "Bekisting"),
        "SUP-15": ("SUB-06", "Road Furniture Installation"),
        "SUP-21": ("SUB-09", "Interior"),
    }
    if install:
        replaced: list[tuple[HierarchyPath, str]] = []
        for candidate, evidence in candidates:
            mapping = supplier_install_map.get(candidate.code)
            if not mapping:
                replaced.append((candidate, evidence))
                continue
            target_code, target_level3 = mapping
            if target_level3 == "Architectural, Landscape & Finishing Works":
                target_level3 = {
                    "Flooring": "Flooring",
                    "Plafon": "Ceiling",
                    "Cat": "Painting",
                    "Aluminium Frame/Facade": "Facade",
                }.get(candidate.level3, "Interior")
            replaced.append((_lookup(paths, target_code, target_level3), evidence))
            notes.append("Boundary Supplier vs Subkontraktor: supply dengan instalasi diperlakukan sebagai paket pekerjaan.")
        candidates = replaced
    elif supply:
        candidates = [item for item in candidates if item[0].level1 != "Subkontraktor"] or candidates

    if consulting:
        consultancy = [item for item in candidates if item[0].level1 == "Jasa Konsultansi"]
        if consultancy:
            candidates = consultancy
            notes.append("Boundary Jasa Konsultansi: terdapat bukti output kajian/desain/supervisi/advisory.")
    elif repair:
        services = [item for item in candidates if item[0].level1 == "Jasa Lainnya"]
        if services:
            candidates = services

    if rent and equipment_candidates:
        candidates = equipment_candidates
        notes.append("Boundary Supplier vs Alat: penyewaan equipment utama diklasifikasikan sebagai Alat.")
    return candidates, notes


def classify_hierarchy_text(
    description: str,
    paths: list[HierarchyPath],
    rules_by_first_token: dict[str, list[TermRule]],
) -> tuple[list[tuple[HierarchyPath, str]], list[str]]:
    text = normalize_name(description)
    if not text:
        return [], []
    raw_candidates: list[tuple[HierarchyPath, str]] = []
    text_tokens = text.split()
    for index, token in enumerate(text_tokens):
        for rule in rules_by_first_token.get(token, []):
            end = index + len(rule.tokens)
            if tuple(text_tokens[index:end]) == rule.tokens:
                raw_candidates.append((rule.path, rule.term))
    if not raw_candidates:
        return [], []

    # Prefer the most specific phrase when one match is wholly contained in a
    # longer match (e.g. BORED PILE versus BORED PILE RIG).
    candidate_by_path: dict[HierarchyPath, str] = {}
    for path, term in sorted(raw_candidates, key=lambda item: -len(normalize_name(item[1]))):
        current = candidate_by_path.get(path)
        if current is None or len(normalize_name(term)) > len(normalize_name(current)):
            candidate_by_path[path] = term
    candidates = list(candidate_by_path.items())
    longest_terms = [normalize_name(term) for _, term in candidates]
    candidates = [
        (path, term)
        for path, term in candidates
        if not any(
            normalize_name(term) != other and _contains(other, normalize_name(term))
            for other in longest_terms
        )
    ]
    candidates, notes = _replace_for_boundary(text, candidates, paths)

    grouped: dict[str, list[tuple[HierarchyPath, str]]] = defaultdict(list)
    for path, term in candidates:
        grouped[normalize_name(term)].append((path, term))
    resolved: list[tuple[HierarchyPath, str]] = []
    ambiguities: list[str] = []
    for term, choices in grouped.items():
        unique = {path for path, _ in choices}
        if len(unique) == 1:
            path = next(iter(unique))
            resolved.append((path, choices[0][1]))
        else:
            ambiguities.append(
                f"Istilah {term!r} mengarah ke beberapa jalur: "
                + " | ".join(f"{p.level1}/{p.code}/{p.level3}" for p in sorted(unique, key=lambda x: x.code))
            )
    deduplicated = sorted(set(resolved), key=lambda item: (item[0].level1, item[0].code, item[0].level3))
    return deduplicated, list(dict.fromkeys(notes + ambiguities))


def classify_po_hierarchy(
    po: pd.DataFrame,
    config_dir: Path,
) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, str]]]:
    paths, rules_by_first_token = load_hierarchy(config_dir)
    vendor_stats: dict[str, dict[HierarchyPath, dict[str, Any]]] = defaultdict(
        lambda: defaultdict(
            lambda: {
                "po_keys": set(), "item_keys": set(), "source_rows": set(),
                "examples": [], "evidence_terms": set(), "boundary_notes": set(),
            }
        )
    )
    unresolved: list[dict[str, Any]] = []
    reviews: list[dict[str, str]] = []

    for row in po.itertuples(index=False):
        matches, notes = classify_hierarchy_text(
            row.description, paths, rules_by_first_token
        )
        ambiguity_notes = [note for note in notes if note.startswith("AMBIGUOUS:") or "mengarah ke beberapa jalur" in note]
        if ambiguity_notes:
            reviews.append(
                {
                    "Severity": "MEDIUM", "Issue": "HIERARCHY_AMBIGUOUS_ITEM",
                    "Source": f"PO {row.company}", "Source Row": str(row.source_row),
                    "ID Vendor": "", "NO SAP": row.sap, "Nama Rekanan": row.name,
                    "Match Method": "PO_ITEM_BOUNDARY_REVIEW",
                    "Detail": f"PO {row.po} item {row.item_po}: " + " | ".join(dict.fromkeys(ambiguity_notes)),
                }
            )
        if not matches:
            unresolved.append(
                {
                    "Company": row.company, "Source Row": row.source_row,
                    "PO": row.po, "Item PO": row.item_po, "NO SAP": row.sap,
                    "Nama Vendor": row.name, "Deskripsi": row.description,
                    "Material": row.material, "Divisi": row.division, "Project": row.project,
                    "Reason": "AMBIGUOUS" if any("AMBIGUOUS:" in note for note in notes) else "NO_MATCH",
                    "Detail": " | ".join(dict.fromkeys(notes)),
                }
            )
            continue

        po_key = f"{row.company}|{row.po}"
        item_key = f"{po_key}|{row.item_po}"
        boundary_notes = [
            note for note in notes
            if not note.startswith("AMBIGUOUS:") and "mengarah ke beberapa jalur" not in note
        ]
        for path, term in matches:
            stats = vendor_stats[row.sap][path]
            stats["po_keys"].add(po_key)
            stats["item_keys"].add(item_key)
            stats["source_rows"].add(f"{row.company}:{row.source_row}")
            stats["evidence_terms"].add(clean_text(term))
            stats["boundary_notes"].update(boundary_notes)
            if row.description and row.description not in stats["examples"] and len(stats["examples"]) < 5:
                stats["examples"].append(row.description)

    vendors: dict[str, dict[str, Any]] = {}
    evidence: list[dict[str, Any]] = []
    for sap, by_path in vendor_stats.items():
        ranked = sorted(
            by_path.items(),
            key=lambda item: (-len(item[1]["item_keys"]), -len(item[1]["po_keys"]), item[0].code, item[0].level3),
        )
        vendors[sap] = {"ranked_paths": ranked}
        for rank, (path, stats) in enumerate(ranked, start=1):
            evidence.append(
                {
                    "NO SAP": sap, "Rank": rank, "Level 1": path.level1,
                    "Kode Level 2": path.code, "Level 2": path.level2, "Level 3": path.level3,
                    "Jumlah PO Berbeda": len(stats["po_keys"]), "Jumlah Item PO": len(stats["item_keys"]),
                    "Baris Sumber PO": " | ".join(sorted(stats["source_rows"])[:25]),
                    "Bukti Istilah": " | ".join(sorted(stats["evidence_terms"])),
                    "Contoh Deskripsi": " | ".join(stats["examples"]),
                    "Boundary Diterapkan": " | ".join(sorted(stats["boundary_notes"])),
                }
            )
    return vendors, evidence, unresolved, reviews


def format_vendor_hierarchy(info: dict[str, Any] | None) -> tuple[str, str, str]:
    if not info:
        return "", "", ""
    ranked = info.get("ranked_paths", [])
    level1 = "\n".join(path.level1 for path, _ in ranked)
    level2 = "\n".join(f"{path.code} | {path.level2}" for path, _ in ranked)
    level3 = "\n".join(f"{path.code} | {path.level3}" for path, _ in ranked)
    return level1, level2, level3
