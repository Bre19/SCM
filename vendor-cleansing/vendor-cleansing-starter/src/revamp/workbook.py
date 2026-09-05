from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


DATA_WIDTHS = [
    14, 14, 30, 18, 8, 15, 9, 11, 9, 9, 9, 9, 10, 23, 15, 18, 20, 20,
    30, 34, 55, 28, 32, 38, 18,
]
AUDIT_WIDTHS = [31, 31, 45, 18, 18, 18, 32, 25, 30, 24, 55, 55]


def _safe_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _row_values(columns: Iterable[str], row: dict[str, Any]) -> list[Any]:
    return [_safe_value(row.get(column, "")) for column in columns]


def _fill(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=color)


def _set_widths(sheet, widths: list[float]) -> None:
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def _add_table(sheet, reference: str, name: str, style: str) -> None:
    table = Table(displayName=name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name=style,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def _build_data_sheet(workbook: Workbook, bundle: dict[str, Any]) -> None:
    columns = bundle["output_columns"]
    rows = bundle["data_rows"]
    if not rows:
        raise ValueError("Data Cleansing tidak memiliki baris vendor")

    sheet = workbook.create_sheet("Data Cleansing")
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 80
    sheet.freeze_panes = "D2"
    sheet.append(columns)
    for row in rows:
        sheet.append(_row_values(columns, row))

    last_row = sheet.max_row
    last_column = sheet.max_column
    last_letter = get_column_letter(last_column)
    _add_table(sheet, f"A1:{last_letter}{last_row}", "DataCleansingTable", "TableStyleMedium2")
    _set_widths(sheet, DATA_WIDTHS)

    for cell in sheet[1]:
        cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for start, end, color in ((1, 4, "5B9BD5"), (5, 12, "4472C4"), (13, 20, "BF9000"), (21, 25, "17365D")):
        for column in range(start, end + 1):
            sheet.cell(1, column).fill = _fill(color)
    sheet.row_dimensions[1].height = 72

    for row_number in range(2, last_row + 1):
        sheet.row_dimensions[row_number].height = 18
        for column in range(1, last_column + 1):
            cell = sheet.cell(row_number, column)
            cell.font = Font(name="Aptos", size=9)
            cell.alignment = Alignment(vertical="top")
        for column in (1, 2, 4):
            cell = sheet.cell(row_number, column)
            cell.value = "" if cell.value is None else str(cell.value)
            cell.number_format = "@"
        for column in range(5, 14):
            sheet.cell(row_number, column).alignment = Alignment(horizontal="center", vertical="top")
        for column in range(19, 26):
            sheet.cell(row_number, column).alignment = Alignment(vertical="top", wrap_text=True)
        sheet.cell(row_number, 25).number_format = "#,##0"

    red_fill = _fill("F4CCCC")
    red_font = Font(color="9C0006", bold=True)
    yellow_fill = _fill("FFF2CC")
    yellow_font = Font(color="9C6500")
    orange_fill = _fill("FCE5CD")
    orange_font = Font(color="783F04", bold=True)
    purple_fill = _fill("E4DFEC")
    purple_font = Font(color="3F3151", bold=True)

    sheet.conditional_formatting.add(
        f"A2:A{last_row}",
        FormulaRule(formula=["LEN(TRIM(A2))=0"], fill=yellow_fill, font=yellow_font),
    )
    sheet.conditional_formatting.add(
        f"B2:B{last_row}",
        FormulaRule(
            formula=[f"COUNTIF($B$2:$B${last_row},B2)>1"],
            fill=red_fill,
            font=red_font,
        ),
    )
    sheet.conditional_formatting.add(
        f"D2:D{last_row}",
        FormulaRule(formula=["LEN(TRIM(D2))=0"], fill=yellow_fill, font=yellow_font),
    )
    sheet.conditional_formatting.add(
        f"I2:I{last_row}",
        FormulaRule(formula=['I2="✓"'], fill=orange_fill, font=orange_font),
    )
    sheet.conditional_formatting.add(
        f"T2:T{last_row}",
        FormulaRule(formula=["LEN(TRIM(T2))=0"], fill=purple_fill, font=purple_font),
    )
    sheet.conditional_formatting.add(
        f"V2:X{last_row}",
        FormulaRule(formula=["LEN(TRIM(V2))=0"], fill=purple_fill, font=purple_font),
    )

    for row_number in range(2, last_row + 1):
        for column, blank_fill, blank_font in (
            (1, yellow_fill, yellow_font),
            (4, yellow_fill, yellow_font),
            (20, purple_fill, purple_font),
        ):
            cell = sheet.cell(row_number, column)
            if not str(cell.value or "").strip():
                cell.fill, cell.font = blank_fill, blank_font
        inject_cell = sheet.cell(row_number, 9)
        if inject_cell.value == "✓":
            inject_cell.fill, inject_cell.font = orange_fill, orange_font
        if not str(sheet.cell(row_number, 22).value or "").strip():
            for column in range(22, 25):
                sheet.cell(row_number, column).fill = purple_fill
                sheet.cell(row_number, column).font = purple_font

    # Direct fills make audit findings visible immediately, even before Excel
    # recalculates conditional formatting. A red HIGH finding takes precedence.
    row_by_sap = {
        str(sheet.cell(row_number, 2).value): row_number
        for row_number in range(2, last_row + 1)
    }
    issue_columns = {
        "ID_VENDOR_CONFLICT": (1, red_fill, red_font),
        "MISSING_ID_VENDOR": (1, yellow_fill, yellow_font),
        "NAME_CONFLICT": (3, yellow_fill, yellow_font),
        "PO_NAME_VARIATION": (3, yellow_fill, yellow_font),
        "NPWP_CONFLICT": (4, red_fill, red_font),
        "MISSING_NPWP": (4, yellow_fill, yellow_font),
        "INVALID_NPWP_LENGTH": (4, yellow_fill, yellow_font),
        "MISSING_MASTER_ATTRIBUTES": (15, yellow_fill, yellow_font),
        "CIRCLE_UNMAPPED": (19, yellow_fill, yellow_font),
        "PO_CIRCLE_NO_OVERLAP": (19, yellow_fill, yellow_font),
        "PO_RULE_GAP_CIRCLE_PRESENT": (20, purple_fill, purple_font),
        "PO_RULE_GAP_CIRCLE_EMPTY": (20, purple_fill, purple_font),
    }
    source_columns = {
        "DRT": 7,
        "DM": 7,
        "DRT_LAMA": 8,
        "DM_LAMA": 8,
        "DCR": 10,
        "DCM": 11,
        "DBCR": 12,
    }
    for review in bundle.get("review_rows", []):
        sap = str(review.get("NO SAP", ""))
        row_number = row_by_sap.get(sap)
        if not row_number:
            continue
        issue = str(review.get("Issue", ""))
        if review.get("Severity") == "HIGH":
            cell = sheet.cell(row_number, 2)
            cell.fill, cell.font = red_fill, red_font
        if issue in issue_columns:
            column, fill, font = issue_columns[issue]
            target_columns = range(15, 18) if issue == "MISSING_MASTER_ATTRIBUTES" else (column,)
            for target_column in target_columns:
                cell = sheet.cell(row_number, target_column)
                if cell.fill.fgColor.rgb not in {"00F4CCCC", "FFF4CCCC"}:
                    cell.fill, cell.font = fill, font
        if issue in {"DUPLICATE_SOURCE_MATCH", "SOURCE_DUPLICATE_SAP"}:
            column = source_columns.get(str(review.get("Source", "")))
            if column:
                cell = sheet.cell(row_number, column)
                if issue == "SOURCE_DUPLICATE_SAP":
                    cell.fill, cell.font = red_fill, red_font
                else:
                    cell.fill, cell.font = yellow_fill, yellow_font


def _style_header(sheet, row: int, start_column: int, end_column: int, color: str) -> None:
    for column in range(start_column, end_column + 1):
        cell = sheet.cell(row, column)
        cell.fill = _fill(color)
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_audit_detail_table(
    sheet,
    start_row: int,
    title: str,
    columns: list[str],
    rows: list[dict[str, Any]],
    table_name: str,
) -> int:
    last_column = len(columns)
    sheet.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=last_column,
    )
    title_cell = sheet.cell(start_row, 1, title)
    title_cell.fill = _fill("D9EAF7")
    title_cell.font = Font(name="Aptos", bold=True, color="17365D")
    for column in range(1, last_column + 1):
        sheet.cell(start_row, column).fill = _fill("D9EAF7")

    header_row = start_row + 1
    for column, value in enumerate(columns, start=1):
        sheet.cell(header_row, column, value)
    _style_header(sheet, header_row, 1, last_column, "17365D")
    sheet.row_dimensions[header_row].height = 42

    for offset, row in enumerate(rows, start=1):
        excel_row = header_row + offset
        for column, value in enumerate(_row_values(columns, row), start=1):
            cell = sheet.cell(excel_row, column, value)
            cell.font = Font(name="Aptos", size=9)
            cell.alignment = Alignment(vertical="top", wrap_text=column >= 4)
            if columns[column - 1] in {"ID Vendor", "NO SAP", "Contoh PO", "Contoh Item PO"}:
                cell.value = "" if cell.value is None else str(cell.value)
                cell.number_format = "@"
        if "Severity" in columns:
            severity_cell = sheet.cell(excel_row, columns.index("Severity") + 1)
            severity_styles = {
                "HIGH": ("F4CCCC", "9C0006"),
                "MEDIUM": ("FFF2CC", "9C6500"),
                "LOW": ("D9EAF7", "1F4E78"),
            }
            colors = severity_styles.get(str(severity_cell.value))
            if colors:
                severity_cell.fill = _fill(colors[0])
                severity_cell.font = Font(name="Aptos", size=9, bold=True, color=colors[1])

    end_row = header_row + len(rows)
    if rows:
        _add_table(
            sheet,
            f"A{header_row}:{get_column_letter(last_column)}{end_row}",
            table_name,
            "TableStyleMedium2",
        )
    return end_row


def _partition_review_rows(
    rows: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    partitions: dict[str, list[dict[str, Any]]] = {
        "duplicates": [],
        "completeness": [],
        "matching": [],
        "classification": [],
        "other": [],
    }
    matching_issues = {
        "SOURCE_SAP_NOT_IN_PO",
        "SOURCE_ID_LINK_OUTSIDE_PO",
        "SOURCE_NO_PO_MATCH",
        "PO_VENDOR_NO_REGISTRY_MATCH",
    }
    for row in rows:
        issue = str(row.get("Issue", ""))
        if issue.startswith("HIERARCHY_"):
            partitions["classification"].append(row)
        elif any(token in issue for token in ("DUPLICATE", "CONFLICT", "AMBIGUOUS")) or issue == "ID_TO_MULTIPLE_SAP":
            partitions["duplicates"].append(row)
        elif issue in matching_issues:
            partitions["matching"].append(row)
        elif any(token in issue for token in ("MISSING", "INVALID", "WITHOUT_VENDOR")):
            partitions["completeness"].append(row)
        elif issue.startswith("PO_") or issue.startswith("CIRCLE_") or issue == "ITEM_TEXT_TRUNCATED":
            partitions["classification"].append(row)
        else:
            partitions["other"].append(row)
    return partitions


def _build_audit_sheet(workbook: Workbook, bundle: dict[str, Any]) -> None:
    sheet = workbook.create_sheet("Audit")
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 90
    _set_widths(sheet, AUDIT_WIDTHS)

    sheet.merge_cells("A1:K1")
    title = sheet["A1"]
    title.value = "Laporan Audit Vendor Cleansing"
    title.fill = _fill("17365D")
    title.font = Font(name="Aptos Display", size=16, bold=True, color="FFFFFF")
    title.alignment = Alignment(horizontal="left", vertical="center")
    sheet.row_dimensions[1].height = 32
    for column in range(1, 12):
        sheet.cell(1, column).fill = _fill("17365D")

    summary_columns = ["Metrik", "Nilai", "Keterangan"]
    summary_start = 3
    sheet.append([])
    for column, value in enumerate(summary_columns, start=1):
        sheet.cell(summary_start, column, value)
    for offset, row in enumerate(bundle["summary_rows"], start=1):
        for column, value in enumerate(_row_values(summary_columns, row), start=1):
            sheet.cell(summary_start + offset, column, value)
    summary_end = summary_start + len(bundle["summary_rows"])
    _add_table(sheet, f"A{summary_start}:C{summary_end}", "AuditSummaryTable", "TableStyleMedium2")
    _style_header(sheet, summary_start, 1, 3, "4472C4")
    for row in range(summary_start + 1, summary_end + 1):
        sheet.cell(row, 2).number_format = "#,##0"

    severity_order = ["HIGH", "MEDIUM", "LOW"]
    severity_counts = Counter(row.get("Severity", "") for row in bundle["review_rows"])
    severity_values = [
        ["Severity", "Jumlah"],
        *[[severity, severity_counts[severity]] for severity in severity_order],
        ["TOTAL", len(bundle["review_rows"])],
    ]
    for row_offset, values in enumerate(severity_values, start=3):
        for column_offset, value in enumerate(values, start=5):
            sheet.cell(row_offset, column_offset, value)
    _add_table(sheet, "E3:F7", "AuditSeverityTable", "TableStyleMedium4")
    _style_header(sheet, 3, 5, 6, "4472C4")
    for row in range(4, 8):
        sheet.cell(row, 6).number_format = "#,##0"

    sheet.merge_cells("H3:I3")
    sheet["H3"] = "Legenda Highlight"
    _style_header(sheet, 3, 8, 9, "4472C4")
    legend = [
        ("Merah", "Anomali HIGH / duplikasi", "F4CCCC"),
        ("Kuning", "Data wajib belum lengkap", "FFF2CC"),
        ("Ungu", "Klasifikasi Final atau Level 1-3 belum terdeteksi dari item PO", "E4DFEC"),
        ("Oranye", "Status Inject: vendor PO belum ditemukan di master aktif/lama; perlu registrasi/daftar ulang", "FCE5CD"),
    ]
    for row, (label, detail, color) in enumerate(legend, start=4):
        sheet.cell(row, 8, label)
        sheet.cell(row, 8).fill = _fill(color)
        sheet.cell(row, 9, detail)

    sheet.merge_cells("E9:I9")
    sheet["E9"] = "Asumsi dan aturan otomasi"
    sheet["E9"].fill = _fill("D9EAF7")
    sheet["E9"].font = Font(name="Aptos", bold=True, color="17365D")
    for column in range(5, 10):
        sheet.cell(9, column).fill = _fill("D9EAF7")
    for index, text in enumerate(bundle.get("assumptions", []), start=10):
        sheet.merge_cells(start_row=index, start_column=5, end_row=index, end_column=9)
        cell = sheet.cell(index, 5, f"• {_safe_value(text)}")
        fill_color = "F5F9FC" if index % 2 == 0 else "FFFFFF"
        cell.fill = _fill(fill_color)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        for column in range(5, 10):
            sheet.cell(index, column).fill = _fill(fill_color)

    review_columns = [
        "Severity", "Issue", "Source", "Source Row", "ID Vendor", "NO SAP",
        "Nama Rekanan", "Match Method", "Detail",
    ]
    partitions = _partition_review_rows(bundle["review_rows"])
    assumptions_end = 9 + len(bundle.get("assumptions", []))
    review_start = max(summary_end, assumptions_end) + 3
    sheet.freeze_panes = f"C{review_start + 2}"
    review_sections = [
        ("Duplikasi, Konflik, dan Pencocokan Ambigu", "duplicates", "AuditDuplicatesTable"),
        ("Data Kosong dan Format Tidak Valid", "completeness", "AuditCompletenessTable"),
        ("Pencocokan Seluruh Record Sumber terhadap PO", "matching", "AuditMatchingTable"),
        ("Rekonsiliasi Klasifikasi", "classification", "AuditClassificationReviewTable"),
        ("Temuan Lain", "other", "AuditOtherFindingsTable"),
    ]
    review_end = review_start - 3
    footer_columns = [
        "Company", "Source File", "Source Row", "Currency", "Nilai PO Footer",
        "Nilai PO Hitung Ulang", "Status Nilai PO", "Harga Satuan Footer",
        "Harga Satuan Hitung Ulang", "Status Harga Satuan", "Keterangan",
    ]
    footer_rows = bundle.get("po_footer_rows", [])
    if footer_rows:
        review_end = _write_audit_detail_table(
            sheet,
            review_end + 3,
            "Rekonsiliasi Baris Total PO",
            footer_columns,
            footer_rows,
            "AuditPOFooterTable",
        )
    for title_text, key, table_name in review_sections:
        section_rows = partitions[key]
        if not section_rows:
            continue
        review_end = _write_audit_detail_table(
            sheet,
            review_end + 3,
            title_text,
            review_columns,
            section_rows,
            table_name,
        )

    evidence_columns = [
        "NO SAP",
        "Nama Vendor PO",
        "Rank",
        "Klasifikasi",
        "Jumlah PO Berbeda",
        "Jumlah Item PO",
        "Rule ID",
        "Confidence Rule",
        "Dukungan Circle",
        "Sumber Final",
        "Contoh Deskripsi",
    ]
    evidence_end = _write_audit_detail_table(
        sheet,
        review_end + 3,
        "Bukti Klasifikasi PO",
        evidence_columns,
        bundle.get("evidence_rows", []),
        "AuditClassificationEvidenceTable",
    )

    hierarchy_evidence_columns = [
        "NO SAP", "Rank", "Level 1", "Kode Level 2", "Level 2", "Level 3",
        "Jumlah PO Berbeda", "Jumlah Item PO", "Baris Sumber PO", "Bukti Istilah",
        "Contoh Deskripsi", "Boundary Diterapkan",
    ]
    hierarchy_evidence_end = _write_audit_detail_table(
        sheet,
        evidence_end + 3,
        "Bukti Kelompok Klasifikasi Level 1-3 dari Item PO",
        hierarchy_evidence_columns,
        bundle.get("hierarchy_evidence_rows", []),
        "AuditHierarchyEvidenceTable",
    )

    hierarchy_unresolved_columns = [
        "Company", "NO SAP", "Nama Vendor", "Deskripsi Belum Memiliki Level",
        "Alasan", "Detail", "Jumlah Item", "Contoh PO", "Contoh Item PO",
        "Baris Sumber PO", "Contoh Project", "Tindakan",
    ]
    hierarchy_unresolved_end = _write_audit_detail_table(
        sheet,
        hierarchy_evidence_end + 3,
        "Item PO Belum Memiliki Kelompok Level 1-3",
        hierarchy_unresolved_columns,
        bundle.get("hierarchy_unresolved_rows", []),
        "AuditHierarchyUnresolvedTable",
    )

    unresolved_columns = [
        "Company",
        "NO SAP",
        "Nama Vendor",
        "Deskripsi Belum Terklasifikasi",
        "Jumlah Item",
        "Contoh PO",
        "Contoh Item PO",
        "Contoh Project",
        "Tindakan",
    ]
    _write_audit_detail_table(
        sheet,
        hierarchy_unresolved_end + 3,
        "Item PO Belum Terklasifikasi",
        unresolved_columns,
        bundle.get("unresolved_rows", []),
        "AuditUnresolvedPOTable",
    )
def build_workbook(bundle_path: Path, output_path: Path) -> Path:
    bundle = json.loads(bundle_path.read_text(encoding="utf-8"))
    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.properties.title = "Data Cleansing Vendor Otomatis"
    workbook.properties.subject = "Hasil klasifikasi vendor dan laporan audit"
    workbook.properties.creator = "SCM Vendor Cleansing"

    _build_data_sheet(workbook, bundle)
    _build_audit_sheet(workbook, bundle)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path
