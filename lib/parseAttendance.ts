import ExcelJS from "exceljs";
import { cellText, toNumber } from "./excelUtils";

export type AttendanceWeek = {
  week: number;
  attendance: number;
};

export type ParsedAttendance = {
  monthLabel: string | null;
  weekend: AttendanceWeek[];
  midweek: AttendanceWeek[];
  warnings: string[];
};

const MONTH_YEAR = /^[A-Za-z]+\s+\d{4}$/;

function findHeaderColumns(
  worksheet: ExcelJS.Worksheet
): { headerRowNumber: number; weekCol: number; attendanceCol: number } | null {
  let result: { headerRowNumber: number; weekCol: number; attendanceCol: number } | null = null;

  worksheet.eachRow((row, rowNumber) => {
    if (result) return;
    const values = row.values as unknown[];
    let weekCol = -1;
    let attendanceCol = -1;
    values.forEach((v, colIndex) => {
      const text = cellText(v).trim().toLowerCase();
      if (text === "week") weekCol = colIndex;
      if (text.startsWith("attendance")) attendanceCol = colIndex;
    });
    if (weekCol !== -1 && attendanceCol !== -1) {
      result = { headerRowNumber: rowNumber, weekCol, attendanceCol };
    }
  });

  return result;
}

function findMonthLabel(worksheet: ExcelJS.Worksheet, beforeRow: number): string | null {
  let label: string | null = null;

  worksheet.eachRow((row, rowNumber) => {
    if (label || rowNumber >= beforeRow) return;
    const values = (row.values as unknown[]).filter((v) => cellText(v).trim() !== "");
    if (values.length !== 1) return;
    const text = cellText(values[0]).trim();
    if (MONTH_YEAR.test(text)) label = text;
  });

  return label;
}

export async function parseAttendanceWorkbook(file: File): Promise<ParsedAttendance> {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  for (const worksheet of workbook.worksheets) {
    const header = findHeaderColumns(worksheet);
    if (!header) continue;

    const { headerRowNumber, weekCol, attendanceCol } = header;
    const monthLabel = findMonthLabel(worksheet, headerRowNumber);
    const weekend: AttendanceWeek[] = [];
    const midweek: AttendanceWeek[] = [];
    let section: "weekend" | "midweek" | null = null;

    for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const values = row.values as unknown[];
      const label = cellText(values[1]).trim().toLowerCase();

      if (label.includes("midweek")) {
        section = "midweek";
        continue;
      }
      if (label.includes("weekend")) {
        section = "weekend";
        continue;
      }
      if (!section) continue;

      const week = toNumber(values[weekCol]);
      const attendance = toNumber(values[attendanceCol]);
      if (week === null || attendance === null) continue;

      (section === "midweek" ? midweek : weekend).push({ week, attendance });
    }

    midweek.sort((a, b) => a.week - b.week);
    weekend.sort((a, b) => a.week - b.week);

    const warnings: string[] = [];
    if (midweek.length === 0 && weekend.length === 0) {
      warnings.push(`${file.name}: found a header row but no "Midweek meeting" / "Weekend meeting" attendance rows`);
    }

    return { monthLabel, weekend, midweek, warnings };
  }

  return {
    monthLabel: null,
    weekend: [],
    midweek: [],
    warnings: [`${file.name}: could not find a "Week" / "Attendance" header row in any sheet`],
  };
}
