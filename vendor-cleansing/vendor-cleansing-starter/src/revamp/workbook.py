from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


DATA_WIDTHS = [
    14, 14, 30, 18, 8, 15, 9, 11, 9, 9, 9, 9, 10, 23, 15, 18, 20, 20,
    30, 34, 55, 28, 32, 38, 18,
]
AUDIT_WIDTHS = [31, 31, 45, 14, 16, 16, 32, 25, 30, 22, 55]


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
        FormulaRule(formula=["LEN(TRIM(A2))=0"], fill=red_fill, font=red_font),
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

    end_row = header_row + len(rows)
    if rows:
        _add_table(
            sheet,
            f"A{header_row}:{get_column_letter(last_column)}{end_row}",
            table_name,
            "TableStyleMedium2",
        )
    return end_row


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
        ("Ungu", "Klasifikasi Final belum terdeteksi", "E4DFEC"),
        ("Oranye", "Vendor hasil Inject / belum di master", "FCE5CD"),
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
    review_start = summary_end + 3
    for column, value in enumerate(review_columns, start=1):
        sheet.cell(review_start, column, value)
    for offset, row in enumerate(bundle["review_rows"], start=1):
        for column, value in enumerate(_row_values(review_columns, row), start=1):
            sheet.cell(review_start + offset, column, value)
    review_end = review_start + len(bundle["review_rows"])
    if bundle["review_rows"]:
        _add_table(
            sheet,
            f"A{review_start}:I{review_end}",
            "AuditFindingsTable",
            "TableStyleMedium4",
        )
    _style_header(sheet, review_start, 1, 9, "17365D")
    sheet.row_dimensions[review_start].height = 42
    sheet.freeze_panes = f"C{review_start + 1}"

    for row in range(review_start + 1, review_end + 1):
        for column in range(1, 10):
            cell = sheet.cell(row, column)
            cell.font = Font(name="Aptos", size=9)
            cell.alignment = Alignment(vertical="top")
        for column in (4, 5, 6):
            cell = sheet.cell(row, column)
            cell.value = "" if cell.value is None else str(cell.value)
            cell.number_format = "@"
        for column in (7, 8, 9):
            sheet.cell(row, column).alignment = Alignment(vertical="top", wrap_text=True)

    if bundle["review_rows"]:
        severity_range = f"A{review_start + 1}:A{review_end}"
        sheet.conditional_formatting.add(
            severity_range,
            CellIsRule(
                operator="equal",
                formula=['"HIGH"'],
                fill=_fill("F4CCCC"),
                font=Font(color="9C0006", bold=True),
            ),
        )
        sheet.conditional_formatting.add(
            severity_range,
            CellIsRule(
                operator="equal",
                formula=['"MEDIUM"'],
                fill=_fill("FFF2CC"),
                font=Font(color="9C6500", bold=True),
            ),
        )
        sheet.conditional_formatting.add(
            severity_range,
            CellIsRule(
                operator="equal",
                formula=['"LOW"'],
                fill=_fill("D9EAF7"),
                font=Font(color="1F4E78"),
            ),
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
        evidence_end + 3,
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
