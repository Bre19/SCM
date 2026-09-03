import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [bundlePath, outputPath] = process.argv.slice(2);
if (!bundlePath || !outputPath) {
  throw new Error("Usage: node build_vendor_cleansing_workbook.mjs <bundle.json> <output.xlsx>");
}

const bundle = JSON.parse(await fs.readFile(bundlePath, "utf8"));
const workbook = Workbook.create();

function columnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function valuesForRows(columns, rows) {
  return rows.map((row) => columns.map((column) => row[column] ?? ""));
}

function applyColumnWidths(sheet, widths, lastRow) {
  widths.forEach((width, index) => {
    const column = columnLetter(index);
    sheet.getRange(`${column}1:${column}${lastRow}`).format.columnWidth = width;
  });
}

const dataSheet = workbook.worksheets.add("Data Cleansing");
dataSheet.showGridLines = false;
const dataMatrix = [bundle.output_columns, ...valuesForRows(bundle.output_columns, bundle.data_rows)];
const dataLastRow = dataMatrix.length;
dataSheet.getRange(`A2:B${dataLastRow}`).format.numberFormat = "@";
dataSheet.getRange(`D2:D${dataLastRow}`).format.numberFormat = "@";
dataSheet.getRange(`A1:Y${dataLastRow}`).values = dataMatrix;
const dataTable = dataSheet.tables.add(`A1:Y${dataLastRow}`, true, "DataCleansingTable");
dataTable.style = "TableStyleMedium2";
dataTable.showFilterButton = true;
dataSheet.freezePanes.freezeRows(1);
dataSheet.freezePanes.freezeColumns(3);

dataSheet.getRange("A1:Y1").format = {
  font: { bold: true, color: "#FFFFFF", fontSize: 10, typeface: "Aptos" },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#0F243E" },
};
dataSheet.getRange("A1:D1").format.fill = "#5B9BD5";
dataSheet.getRange("E1:L1").format.fill = "#4472C4";
dataSheet.getRange("M1:T1").format.fill = "#BF9000";
dataSheet.getRange("U1:Y1").format.fill = "#17365D";
dataSheet.getRange("A1:Y1").format.rowHeight = 72;
dataSheet.getRange(`A2:Y${dataLastRow}`).format = {
  font: { fontSize: 9, typeface: "Aptos" },
  verticalAlignment: "top",
};
dataSheet.getRange(`A2:Y${dataLastRow}`).format.rowHeight = 18;
dataSheet.getRange(`E2:M${dataLastRow}`).format.horizontalAlignment = "center";
dataSheet.getRange(`S2:Y${dataLastRow}`).format.wrapText = true;
dataSheet.getRange(`Y2:Y${dataLastRow}`).format.numberFormat = "#,##0";
applyColumnWidths(
  dataSheet,
  [14, 14, 30, 18, 8, 15, 9, 11, 9, 9, 9, 9, 10, 23, 15, 18, 20, 20, 30, 34, 55, 28, 32, 38, 18],
  dataLastRow,
);

const auditSheet = workbook.worksheets.add("Audit");
auditSheet.showGridLines = false;
auditSheet.getRange("A1:I1").merge();
auditSheet.getRange("A1").values = [["Laporan Audit Vendor Cleansing"]];
auditSheet.getRange("A1:I1").format = {
  fill: "#17365D",
  font: { bold: true, color: "#FFFFFF", fontSize: 16, typeface: "Aptos Display" },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
auditSheet.getRange("A1:I1").format.rowHeight = 32;

const summaryColumns = ["Metrik", "Nilai", "Keterangan"];
const summaryStart = 3;
const summaryMatrix = [summaryColumns, ...valuesForRows(summaryColumns, bundle.summary_rows)];
const summaryEnd = summaryStart + summaryMatrix.length - 1;
auditSheet.getRange(`A${summaryStart}:C${summaryEnd}`).values = summaryMatrix;
const summaryTable = auditSheet.tables.add(
  `A${summaryStart}:C${summaryEnd}`,
  true,
  "AuditSummaryTable",
);
summaryTable.style = "TableStyleMedium2";
auditSheet.getRange(`A${summaryStart}:C${summaryStart}`).format = {
  fill: "#4472C4",
  font: { bold: true, color: "#FFFFFF" },
};
auditSheet.getRange(`B${summaryStart + 1}:B${summaryEnd}`).format.numberFormat = "#,##0";

const severityOrder = ["HIGH", "MEDIUM", "LOW"];
const severityCounts = new Map(severityOrder.map((severity) => [severity, 0]));
for (const row of bundle.review_rows) {
  severityCounts.set(row.Severity, (severityCounts.get(row.Severity) ?? 0) + 1);
}
auditSheet.getRange("E3:F7").values = [
  ["Severity", "Jumlah"],
  ...severityOrder.map((severity) => [severity, severityCounts.get(severity) ?? 0]),
  ["TOTAL", bundle.review_rows.length],
];
const severityTable = auditSheet.tables.add("E3:F7", true, "AuditSeverityTable");
severityTable.style = "TableStyleMedium4";
auditSheet.getRange("E3:F3").format = {
  fill: "#4472C4",
  font: { bold: true, color: "#FFFFFF" },
};
auditSheet.getRange("F4:F7").format.numberFormat = "#,##0";

auditSheet.getRange("H3:I3").merge();
auditSheet.getRange("H3").values = [["Legenda Highlight"]];
auditSheet.getRange("H3:I3").format = {
  fill: "#4472C4",
  font: { bold: true, color: "#FFFFFF" },
};
auditSheet.getRange("H4:I7").values = [
  ["Merah", "Anomali HIGH / duplikasi"],
  ["Kuning", "Data wajib belum lengkap"],
  ["Ungu", "Klasifikasi Final belum terdeteksi"],
  ["Oranye", "Vendor hasil Inject / belum di master"],
];
auditSheet.getRange("H4:H4").format.fill = "#F4CCCC";
auditSheet.getRange("H5:H5").format.fill = "#FFF2CC";
auditSheet.getRange("H6:H6").format.fill = "#E4DFEC";
auditSheet.getRange("H7:H7").format.fill = "#FCE5CD";

auditSheet.getRange("E9:I9").merge();
auditSheet.getRange("E9").values = [["Asumsi dan aturan otomasi"]];
auditSheet.getRange("E9:I9").format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#17365D" },
};
bundle.assumptions.forEach((text, index) => {
  const row = 10 + index;
  auditSheet.getRange(`E${row}:I${row}`).merge();
  auditSheet.getRange(`E${row}`).values = [[`• ${text}`]];
  auditSheet.getRange(`E${row}:I${row}`).format = {
    wrapText: true,
    verticalAlignment: "center",
    fill: index % 2 === 0 ? "#F5F9FC" : "#FFFFFF",
  };
});

const reviewColumns = [
  "Severity",
  "Issue",
  "Source",
  "Source Row",
  "ID Vendor",
  "NO SAP",
  "Nama Rekanan",
  "Match Method",
  "Detail",
];
const reviewStart = summaryEnd + 3;
const reviewMatrix = [reviewColumns, ...valuesForRows(reviewColumns, bundle.review_rows)];
const reviewEnd = reviewStart + reviewMatrix.length - 1;
auditSheet.getRange(`D${reviewStart + 1}:F${reviewEnd}`).format.numberFormat = "@";
auditSheet.getRange(`A${reviewStart}:I${reviewEnd}`).values = reviewMatrix;
const reviewTable = auditSheet.tables.add(
  `A${reviewStart}:I${reviewEnd}`,
  true,
  "AuditFindingsTable",
);
reviewTable.style = "TableStyleMedium4";
reviewTable.showFilterButton = true;
auditSheet.getRange(`A${reviewStart}:I${reviewStart}`).format = {
  fill: "#17365D",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
auditSheet.getRange(`A${reviewStart}:I${reviewStart}`).format.rowHeight = 42;
auditSheet.getRange(`A${reviewStart + 1}:I${reviewEnd}`).format = {
  font: { fontSize: 9, typeface: "Aptos" },
  verticalAlignment: "top",
};
auditSheet.getRange(`G${reviewStart + 1}:I${reviewEnd}`).format.wrapText = true;
applyColumnWidths(auditSheet, [31, 31, 45, 14, 16, 16, 32, 25, 68], reviewEnd);
auditSheet.freezePanes.freezeRows(reviewStart);
auditSheet.freezePanes.freezeColumns(2);

const reviewSeverityRange = auditSheet.getRange(`A${reviewStart + 1}:A${reviewEnd}`);
reviewSeverityRange.conditionalFormats.add("containsText", {
  text: "HIGH",
  format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } },
});
reviewSeverityRange.conditionalFormats.add("containsText", {
  text: "MEDIUM",
  format: { fill: "#FFF2CC", font: { color: "#9C6500", bold: true } },
});
reviewSeverityRange.conditionalFormats.add("containsText", {
  text: "LOW",
  format: { fill: "#D9EAF7", font: { color: "#1F4E78" } },
});

dataSheet.getRange(`A2:A${dataLastRow}`).conditionalFormats.add("containsBlanks", {
  format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } },
});
dataSheet.getRange(`B2:B${dataLastRow}`).conditionalFormats.add("duplicateValues", {
  format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } },
});
dataSheet.getRange(`D2:D${dataLastRow}`).conditionalFormats.add("containsBlanks", {
  format: { fill: "#FFF2CC", font: { color: "#9C6500" } },
});
dataSheet.getRange(`I2:I${dataLastRow}`).conditionalFormats.add("containsText", {
  text: "✓",
  format: { fill: "#FCE5CD", font: { color: "#783F04", bold: true } },
});
dataSheet.getRange(`T2:T${dataLastRow}`).conditionalFormats.add("containsBlanks", {
  format: { fill: "#E4DFEC", font: { color: "#3F3151", bold: true } },
});

const previewDir = path.join(path.dirname(bundlePath), "preview");
await fs.mkdir(previewDir, { recursive: true });
const previews = [
  ["Data Cleansing", `A1:Y${Math.min(dataLastRow, 18)}`, "data-cleansing.png"],
  ["Audit", "A1:I20", "audit-summary.png"],
  ["Audit", `A${reviewStart}:I${Math.min(reviewEnd, reviewStart + 14)}`, "audit-findings.png"],
];
for (const [sheetName, range, filename] of previews) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(previewDir, filename), new Uint8Array(await preview.arrayBuffer()));
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(
  JSON.stringify({
    workbook: outputPath,
    rows: bundle.data_rows.length,
    sheets: 2,
    auditFindings: bundle.review_rows.length,
    previewDir,
  }),
);
