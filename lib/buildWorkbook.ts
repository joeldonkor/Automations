import ExcelJS from "exceljs";
import { REPORT_ROWS } from "./types";

export type BuildWorkbookInput = {
  monthLabel: string;
  groupLabels: string[];
  tableValues: number[][]; // [rowIndex][groupIndex], matches REPORT_ROWS order
  remarks: string[]; // per row
  attendance: {
    weekend: number[];
    midweek: number[];
  } | null;
};

const CYAN = "FF0E7C92";
const PINK = "FFC2185B";
const GREEN = "FF1A7F37";
const BORDER: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFB0B4BC" } };

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
}

export async function buildWorkbook(input: BuildWorkbookInput): Promise<ArrayBuffer> {
  const { monthLabel, groupLabels, tableValues, remarks, attendance } = input;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Monthly Report");

  const groupCount = groupLabels.length;
  const totalCols = 1 + groupCount + 1 + 1; // label + groups + totals + remarks

  sheet.getColumn(1).width = 30;
  for (let i = 0; i < groupCount; i++) sheet.getColumn(2 + i).width = 14;
  sheet.getColumn(2 + groupCount).width = 12;
  sheet.getColumn(3 + groupCount).width = 24;

  // Title row
  const titleRow = sheet.addRow(["MONTHLY CONGREGATIONAL REPORT"]);
  sheet.mergeCells(titleRow.number, 1, titleRow.number, totalCols);
  const titleCell = titleRow.getCell(1);
  titleCell.font = { bold: true, size: 14, color: { argb: CYAN } };
  titleCell.alignment = { horizontal: "center" };

  // Header row: GROUPS | labels | TOTALS | REMARKS
  const headerRow = sheet.addRow(["GROUPS", ...groupLabels, "TOTALS", "REMARKS"]);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: CYAN } };
    cell.alignment = { horizontal: "center" };
    applyBorder(cell);
  });

  // Month row
  const monthRow = sheet.addRow([
    "MONTH",
    ...groupLabels.map(() => monthLabel),
    monthLabel,
    "",
  ]);
  monthRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: CYAN } };
    cell.alignment = { horizontal: "center" };
    applyBorder(cell);
  });

  // Category rows
  REPORT_ROWS.forEach((rowDef, rowIndex) => {
    const groupVals = tableValues[rowIndex] ?? groupLabels.map(() => 0);
    const total = groupVals.reduce((a, b) => a + (b || 0), 0);
    const row = sheet.addRow([
      rowDef.label,
      ...groupVals,
      total,
      remarks[rowIndex] ?? "",
    ]);

    const labelCell = row.getCell(1);
    labelCell.font = { bold: true, color: { argb: PINK } };
    labelCell.alignment = { horizontal: "left" };
    applyBorder(labelCell);

    for (let g = 0; g < groupCount; g++) {
      const cell = row.getCell(2 + g);
      cell.font = { color: { argb: GREEN }, bold: true };
      cell.alignment = { horizontal: "center" };
      applyBorder(cell);
    }

    const totalCell = row.getCell(2 + groupCount);
    totalCell.font = { bold: true, color: { argb: PINK } };
    totalCell.alignment = { horizontal: "center" };
    applyBorder(totalCell);

    const remarksCell = row.getCell(3 + groupCount);
    remarksCell.alignment = { horizontal: "left" };
    applyBorder(remarksCell);
  });

  if (attendance) {
    sheet.addRow([]);
    const totalWeekend = attendance.weekend.reduce((a, b) => a + (b || 0), 0);
    const totalMidweek = attendance.midweek.reduce((a, b) => a + (b || 0), 0);
    const weeksWeekend = attendance.weekend.filter((v) => v > 0).length || 1;
    const weeksMidweek = attendance.midweek.filter((v) => v > 0).length || 1;
    const avgWeekend = Math.round(totalWeekend / weeksWeekend);
    const avgMidweek = Math.round(totalMidweek / weeksMidweek);

    const attStartRow = sheet.rowCount + 1;
    const rowWeekend = sheet.addRow([
      "MEETING ATTENDANCE",
      "TOTAL",
      totalWeekend,
      "÷",
      weeksWeekend,
      "=",
      "AVERAGE",
      avgWeekend,
      "WEEKEND",
    ]);
    const rowMidweek = sheet.addRow([
      "",
      "",
      totalMidweek,
      "÷",
      weeksMidweek,
      "=",
      "AVERAGE",
      avgMidweek,
      "MIDWEEK",
    ]);
    sheet.mergeCells(attStartRow, 1, attStartRow + 1, 1);

    [rowWeekend, rowMidweek].forEach((row) => {
      row.eachCell((cell) => {
        applyBorder(cell);
        cell.alignment = { horizontal: "center" };
      });
      row.getCell(1).font = { bold: true, color: { argb: CYAN } };
      row.getCell(1).alignment = { horizontal: "left" };
      row.getCell(8).font = { bold: true, color: { argb: PINK } };
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}
