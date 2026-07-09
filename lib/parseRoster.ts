import ExcelJS from "exceljs";
import type { PublisherRow } from "./types";

function isMarked(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const s = String(value).trim().toLowerCase();
  if (s === "") return false;
  if (["0", "no", "false", "n"].includes(s)) return false;
  return true;
}

function toNumber(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  if (typeof value === "object" && value !== null && "result" in (value as Record<string, unknown>)) {
    // ExcelJS formula cell result
    return toNumber((value as { result: unknown }).result);
  }
  const n = Number(String(value).trim());
  return Number.isFinite(n) ? n : null;
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "text" in (value as Record<string, unknown>)) {
    return String((value as { text: unknown }).text ?? "");
  }
  if (typeof value === "object" && "result" in (value as Record<string, unknown>)) {
    return String((value as { result: unknown }).result ?? "");
  }
  return String(value);
}

const COLUMN_ALIASES: Record<string, string> = {
  publishers: "name",
  publisher: "name",
  name: "name",
  number: "number",
  shared: "reported",
  "b.studies": "bibleStudies",
  bstudies: "bibleStudies",
  "bible studies": "bibleStudies",
  studies: "bibleStudies",
  "aux.": "aux",
  aux: "aux",
  auxiliary: "aux",
  hours: "hours",
  notes: "notes",
  remarks: "notes",
};

export type ParsedGroup = {
  label: string;
  rows: PublisherRow[];
};

export type ParsedWorkbook = {
  groups: ParsedGroup[];
  warnings: string[];
};

function parseWorksheet(
  worksheet: ExcelJS.Worksheet,
  inactiveKeywords: string[]
): PublisherRow[] | null {
  let headerRowNumber = -1;
  const columnMap: Record<number, string> = {};

  worksheet.eachRow((row, rowNumber) => {
    if (headerRowNumber !== -1) return;
    const values = row.values as unknown[];
    const hasPublishersHeader = values.some(
      (v) => cellText(v).trim().toLowerCase() === "publishers"
    );
    if (hasPublishersHeader) {
      headerRowNumber = rowNumber;
      values.forEach((v, colIndex) => {
        const key = cellText(v).trim().toLowerCase();
        if (COLUMN_ALIASES[key]) {
          columnMap[colIndex] = COLUMN_ALIASES[key];
        }
      });
    }
  });

  if (headerRowNumber === -1) return null;

  const nameCol = Object.entries(columnMap).find(([, v]) => v === "name")?.[0];
  if (!nameCol) return null;

  const reportedCol = Object.entries(columnMap).find(([, v]) => v === "reported")?.[0];
  const bibleStudiesCol = Object.entries(columnMap).find(([, v]) => v === "bibleStudies")?.[0];
  const auxCol = Object.entries(columnMap).find(([, v]) => v === "aux")?.[0];
  const hoursCol = Object.entries(columnMap).find(([, v]) => v === "hours")?.[0];
  const notesCol = Object.entries(columnMap).find(([, v]) => v === "notes")?.[0];

  const rows: PublisherRow[] = [];
  const keywordsLower = inactiveKeywords.map((k) => k.trim().toLowerCase()).filter(Boolean);

  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const values = row.values as unknown[];

    const name = cellText(values[Number(nameCol)]).trim();
    if (!name) continue;
    if (name.toLowerCase() === "total") break;

    const reportedRaw = reportedCol ? values[Number(reportedCol)] : undefined;
    const bibleStudiesRaw = bibleStudiesCol ? values[Number(bibleStudiesCol)] : undefined;
    const auxRaw = auxCol ? values[Number(auxCol)] : undefined;
    const hoursRaw = hoursCol ? values[Number(hoursCol)] : undefined;
    const notesRaw = notesCol ? values[Number(notesCol)] : undefined;

    const isLabelRow =
      !isMarked(reportedRaw) &&
      toNumber(bibleStudiesRaw) === null &&
      !isMarked(auxRaw) &&
      toNumber(hoursRaw) === null &&
      cellText(notesRaw).trim() === "";
    if (isLabelRow) continue;

    const reported = isMarked(reportedRaw);
    const bibleStudies = toNumber(bibleStudiesRaw) ?? 0;
    const isAux = isMarked(auxRaw);
    const hours = toNumber(hoursRaw);
    const notes = cellText(notesRaw).trim();

    const inactive =
      keywordsLower.length > 0 &&
      keywordsLower.some((kw) => notes.toLowerCase().includes(kw));

    const category = isAux ? "auxPioneer" : hours !== null && hours > 0 ? "regularPioneer" : "publisher";

    rows.push({ name, reported, bibleStudies, hours, notes, inactive, category });
  }

  return rows;
}

function labelFromFileName(name: string): string {
  return name.replace(/\.xlsx$/i, "");
}

/**
 * Parses one uploaded workbook into one or more groups: a workbook with a single
 * roster sheet becomes one group (named after the file), while a workbook with
 * several roster sheets (one per group) becomes one group per sheet (named after
 * the sheet tab).
 */
export async function parseRosterWorkbook(
  file: File,
  inactiveKeywords: string[]
): Promise<ParsedWorkbook> {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const found: ParsedGroup[] = [];
  for (const worksheet of workbook.worksheets) {
    const rows = parseWorksheet(worksheet, inactiveKeywords);
    if (rows !== null) {
      found.push({ label: worksheet.name, rows });
    }
  }

  if (found.length === 0) {
    return {
      groups: [],
      warnings: [`${file.name}: could not find a "Publishers" header row in any sheet`],
    };
  }

  // A single usable sheet: the file name is a more meaningful group label than "Sheet1".
  if (found.length === 1) {
    found[0].label = labelFromFileName(file.name);
  }

  return { groups: found, warnings: [] };
}
