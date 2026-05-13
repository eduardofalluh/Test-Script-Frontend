"use client";

import * as XLSX from "@e965/xlsx";
import type { EditableWorkbook, WorkbookPreview } from "@/lib/types";

export async function parseWorkbook(file: File): Promise<WorkbookPreview> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0] ?? "Sheet1";
  const firstSheet = workbook.Sheets[firstSheetName];
  const firstSheetRows = firstSheet
    ? (XLSX.utils.sheet_to_json(firstSheet, {
        header: 1,
        blankrows: false,
        defval: ""
      }) as unknown[][])
    : [];

  return {
    sheetNames: workbook.SheetNames,
    firstSheetName,
    firstSheetRows: firstSheetRows.slice(0, 20),
    firstSheetRowCount: firstSheetRows.length
  };
}

export function parseWorkbookFromBase64(base64: string): WorkbookPreview {
  const workbook = XLSX.read(base64, { type: "base64" });
  const firstSheetName = workbook.SheetNames[0] ?? "Sheet1";
  const firstSheet = workbook.Sheets[firstSheetName];
  const firstSheetRows = firstSheet
    ? (XLSX.utils.sheet_to_json(firstSheet, {
        header: 1,
        blankrows: false,
        defval: ""
      }) as unknown[][])
    : [];

  return {
    sheetNames: workbook.SheetNames,
    firstSheetName,
    firstSheetRows: firstSheetRows.slice(0, 20),
    firstSheetRowCount: firstSheetRows.length
  };
}

export function parseEditableWorkbookFromBase64(base64: string): EditableWorkbook {
  const workbook = XLSX.read(base64, { type: "base64" });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = worksheet
      ? (XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          blankrows: true,
          defval: ""
        }) as unknown[][]).map((row) => row.map(formatEditableCell))
      : [];
    const columnCount = Math.max(...rows.map((row) => row.length), 1);

    return {
      name: sheetName,
      rows: padRows(rows, columnCount),
      rowCount: rows.length,
      columnCount
    };
  });

  return {
    sheets,
    activeSheetName: sheets[0]?.name ?? "Sheet1"
  };
}

export function editableWorkbookToBase64(workbook: EditableWorkbook) {
  const nextWorkbook = XLSX.utils.book_new();

  workbook.sheets.forEach((sheet) => {
    const worksheet = XLSX.utils.aoa_to_sheet(trimEmptyEdges(sheet.rows));
    XLSX.utils.book_append_sheet(nextWorkbook, worksheet, sheet.name);
  });

  return XLSX.write(nextWorkbook, {
    type: "base64",
    bookType: "xlsx"
  }) as string;
}

function formatEditableCell(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value);
}

function padRows(rows: string[][], columnCount: number) {
  return rows.map((row) => {
    if (row.length >= columnCount) {
      return row;
    }

    return [...row, ...Array.from({ length: columnCount - row.length }, () => "")];
  });
}

function trimEmptyEdges(rows: string[][]) {
  const lastNonEmptyRow = rows.reduce((lastIndex, row, index) => {
    return row.some((cell) => cell.trim() !== "") ? index : lastIndex;
  }, -1);

  if (lastNonEmptyRow === -1) {
    return [[]];
  }

  const usedRows = rows.slice(0, lastNonEmptyRow + 1);
  const lastNonEmptyColumn = usedRows.reduce((lastColumn, row) => {
    const rowLastColumn = row.reduce((lastIndex, cell, index) => {
      return cell.trim() !== "" ? index : lastIndex;
    }, -1);
    return Math.max(lastColumn, rowLastColumn);
  }, 0);

  return usedRows.map((row) => row.slice(0, lastNonEmptyColumn + 1));
}
