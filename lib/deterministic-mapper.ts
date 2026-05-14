import * as XLSX from "@e965/xlsx";
import { EXCEL_MIME_TYPE } from "@/lib/constants";
import { sanitizeFilename } from "@/lib/utils";
import type { AgentFilePayload } from "@/lib/types";

type WorkbookInput = {
  filename: string;
  buffer: Buffer;
};

type DeterministicMappingInput = {
  template: WorkbookInput;
  sources: WorkbookInput[];
  prompt: string;
};

type ColumnMapping = {
  sourceColumnIndex: number | null;
  targetColumnIndex: number;
  sourceLabel: string;
  targetLabel: string;
  generated?: "source-row" | "validation-note";
};

type MappingRule = {
  sourceSheetName: string;
  targetSheetName: string;
  targetStartRow: number;
  mappings: ColumnMapping[];
  convertCountryToIso2: boolean;
};

const COUNTRY_TO_ISO2: Record<string, string> = {
  brazil: "BR",
  canada: "CA",
  france: "FR",
  germany: "DE",
  india: "IN",
  japan: "JP",
  mexico: "MX",
  netherlands: "NL",
  spain: "ES",
  sweden: "SE",
  "united kingdom": "GB",
  uk: "GB",
  "united states": "US",
  usa: "US"
};

const HEADER_ALIASES: Record<string, string[]> = {
  customerid: ["legacyid", "customerid", "customer", "id"],
  name: ["customername", "name", "customer"],
  contactemail: ["email", "contactemail", "e-mail"],
  email: ["email", "contactemail", "e-mail"],
  countrynametext: ["country", "countryname", "countrytext"],
  country: ["country", "countryname", "countrytext"],
  currency: ["currency", "curr"],
  riskscore: ["riskscore", "risk"],
  createdon: ["createdon", "createddate", "date"]
};

export function mapWorkbookDeterministically(input: DeterministicMappingInput) {
  if (input.sources.length === 0) {
    throw new Error("At least one source workbook is required for deterministic mapping.");
  }

  const templateWorkbook = XLSX.read(input.template.buffer, { type: "buffer", cellDates: true });
  const sourceWorkbooks = input.sources.map((source) => ({
    filename: source.filename,
    workbook: XLSX.read(source.buffer, { type: "buffer", cellDates: true })
  }));

  const rule = parseMappingRule(input.prompt, templateWorkbook, sourceWorkbooks);
  const sourceWorkbook = sourceWorkbooks.find((source) => findSheetName(source.workbook, rule.sourceSheetName));
  if (!sourceWorkbook) {
    throw new Error(`Could not find source sheet "${rule.sourceSheetName}" in the uploaded source workbooks.`);
  }

  const resolvedSourceSheetName = findSheetName(sourceWorkbook.workbook, rule.sourceSheetName);
  const resolvedTargetSheetName = findSheetName(templateWorkbook, rule.targetSheetName);
  if (!resolvedSourceSheetName) {
    throw new Error(`Could not find source sheet "${rule.sourceSheetName}".`);
  }
  if (!resolvedTargetSheetName) {
    throw new Error(`Could not find target sheet "${rule.targetSheetName}" in the template.`);
  }

  const sourceSheet = sourceWorkbook.workbook.Sheets[resolvedSourceSheetName];
  const targetSheet = templateWorkbook.Sheets[resolvedTargetSheetName];
  const sourceRows = sheetToRows(sourceSheet);
  const sourceHeaderRowIndex = detectHeaderRowIndex(sourceRows);
  const sourceDataRows = sourceRows.slice(sourceHeaderRowIndex + 1);
  const mappedRows = sourceDataRows.filter((row) => row.some((cell) => cell !== ""));

  mappedRows.forEach((sourceRow, rowOffset) => {
    const targetRowIndex = rule.targetStartRow - 1 + rowOffset;
    rule.mappings.forEach((mapping) => {
      const value = resolveMappedValue({
        mapping,
        sourceRow,
        sourceExcelRowNumber: sourceHeaderRowIndex + rowOffset + 2,
        convertCountryToIso2: rule.convertCountryToIso2
      });
      writeCell(targetSheet, targetRowIndex, mapping.targetColumnIndex, value);
    });
  });

  const base64 = XLSX.write(templateWorkbook, {
    type: "base64",
    bookType: "xlsx"
  }) as string;

  const filename = `populated-${sanitizeFilename(input.template.filename).replace(/\.(xlsx|xlsm)$/i, "")}.xlsx`;
  const file: AgentFilePayload = {
    filename,
    mimeType: EXCEL_MIME_TYPE,
    base64
  };

  return {
    file,
    responseText: [
      "Deterministic mapper completed without using the external AI agent.",
      "",
      `Source workbook: ${sourceWorkbook.filename}`,
      `Source sheet: ${resolvedSourceSheetName}`,
      `Target sheet: ${resolvedTargetSheetName}`,
      `Target start row: ${rule.targetStartRow}`,
      `Rows mapped: ${mappedRows.length}`,
      `Column mappings: ${rule.mappings.map((mapping) => `${mapping.sourceLabel}->${mapping.targetLabel}`).join(", ")}`,
      rule.convertCountryToIso2 ? "Transformations: country names converted to ISO-2 where recognized." : "Transformations: none."
    ].join("\n"),
    rawResponse: {
      mode: "deterministic",
      rule,
      rowsMapped: mappedRows.length
    }
  };
}

function parseMappingRule(
  prompt: string,
  templateWorkbook: XLSX.WorkBook,
  sources: Array<{ filename: string; workbook: XLSX.WorkBook }>
): MappingRule {
  const sourceSheetName = extractSourceSheetName(prompt) ?? sources[0]?.workbook.SheetNames[0];
  const targetSheetName = extractTargetSheetName(prompt) ?? templateWorkbook.SheetNames[0];
  const targetStartRow = extractTargetStartRow(prompt);

  if (!sourceSheetName) {
    throw new Error("Could not determine the source sheet. Add wording like: source sheet 'Customers'.");
  }
  if (!targetSheetName) {
    throw new Error("Could not determine the target sheet. Add wording like: target sheet 'Migration Input'.");
  }
  if (!targetStartRow) {
    throw new Error("Could not determine the target start row. Add wording like: starting row 5.");
  }

  const sourceWorkbook = sources.find((source) => findSheetName(source.workbook, sourceSheetName)) ?? sources[0];
  const resolvedSourceSheetName = findSheetName(sourceWorkbook.workbook, sourceSheetName);
  const resolvedTargetSheetName = findSheetName(templateWorkbook, targetSheetName);
  if (!resolvedSourceSheetName || !resolvedTargetSheetName) {
    throw new Error("Could not resolve the requested source or target sheet name.");
  }

  const sourceRows = sheetToRows(sourceWorkbook.workbook.Sheets[resolvedSourceSheetName]);
  const targetRows = sheetToRows(templateWorkbook.Sheets[resolvedTargetSheetName]);
  const sourceHeaders = sourceRows[detectHeaderRowIndex(sourceRows)] ?? [];
  const targetHeaders = targetRows[targetStartRow - 2] ?? [];
  const explicitMappings = extractExplicitColumnMappings(prompt, sourceHeaders, targetHeaders);
  const mappings = explicitMappings.length > 0 ? explicitMappings : inferHeaderMappings(sourceHeaders, targetHeaders);

  if (mappings.length === 0) {
    throw new Error(
      "Could not determine column mappings. Add explicit mappings like A->C, B->D, or use matching source and target header names."
    );
  }

  return {
    sourceSheetName,
    targetSheetName,
    targetStartRow,
    mappings,
    convertCountryToIso2: /iso-?2|country names?\s+to\s+iso/i.test(prompt)
  };
}

function extractSourceSheetName(prompt: string) {
  return (
    matchFirst(prompt, /source(?: workbook| file| data)?[^.]*?sheet\s+['"]([^'"]+)['"]/i) ??
    matchFirst(prompt, /source(?: workbook| file| data)?[^.]*?sheet\s+([A-Za-z0-9 _-]+?)(?:\s+(?:rows?|columns?|into|starting|and)|[.,]|$)/i)
  );
}

function extractTargetSheetName(prompt: string) {
  return (
    matchFirst(prompt, /(?:target|template)(?:'s)?\s+['"]([^'"]+)['"]\s+sheet/i) ??
    matchFirst(prompt, /(?:target|template)[^.]*?sheet\s+['"]([^'"]+)['"]/i) ??
    matchFirst(prompt, /into[^.]*?sheet\s+['"]([^'"]+)['"]/i) ??
    matchFirst(prompt, /into[^.]*?sheet\s+([A-Za-z0-9 _-]+?)\s+starting row/i)
  );
}

function extractTargetStartRow(prompt: string) {
  const match = prompt.match(/starting\s+row\s+(\d+)/i);
  return match?.[1] ? Number.parseInt(match[1], 10) : null;
}

function extractExplicitColumnMappings(prompt: string, sourceHeaders: unknown[], targetHeaders: unknown[]): ColumnMapping[] {
  const mappings: ColumnMapping[] = [];
  const arrowPattern = /\b([A-Za-z]{1,3})\s*(?:->|=>)\s*([A-Za-z]{1,3})\b/g;
  const wordPattern = /\bsource\s+([A-Za-z]{1,3})\s+to\s+target\s+([A-Za-z]{1,3})\b/gi;

  collectColumnMappings(prompt, arrowPattern, sourceHeaders, targetHeaders, mappings);
  collectColumnMappings(prompt, wordPattern, sourceHeaders, targetHeaders, mappings);

  return uniqueTargetMappings(mappings);
}

function collectColumnMappings(
  prompt: string,
  pattern: RegExp,
  sourceHeaders: unknown[],
  targetHeaders: unknown[],
  mappings: ColumnMapping[]
) {
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(prompt)) !== null) {
    const sourceColumnIndex = columnNameToIndex(match[1]);
    const targetColumnIndex = columnNameToIndex(match[2]);
    if (!Number.isSafeInteger(sourceColumnIndex) || !Number.isSafeInteger(targetColumnIndex)) {
      continue;
    }

    mappings.push({
      sourceColumnIndex,
      targetColumnIndex,
      sourceLabel: String(sourceHeaders[sourceColumnIndex] || match[1]),
      targetLabel: String(targetHeaders[targetColumnIndex] || match[2])
    });
  }
}

function inferHeaderMappings(sourceHeaders: unknown[], targetHeaders: unknown[]): ColumnMapping[] {
  const sourceHeaderNames = sourceHeaders.map((header) => String(header || ""));
  const mappings: ColumnMapping[] = [];

  targetHeaders.forEach((targetHeader, targetColumnIndex) => {
    const targetLabel = String(targetHeader || "").trim();
    if (!targetLabel) {
      return;
    }

    const normalizedTarget = normalizeHeader(targetLabel);
    if (normalizedTarget === "sourcerow") {
      mappings.push({
        sourceColumnIndex: null,
        targetColumnIndex,
        sourceLabel: "source row number",
        targetLabel,
        generated: "source-row"
      });
      return;
    }

    if (normalizedTarget.includes("validation")) {
      mappings.push({
        sourceColumnIndex: null,
        targetColumnIndex,
        sourceLabel: "validation note",
        targetLabel,
        generated: "validation-note"
      });
      return;
    }

    const sourceColumnIndex = sourceHeaderNames.findIndex((sourceHeader) => headersMatch(sourceHeader, targetLabel));
    if (sourceColumnIndex >= 0) {
      mappings.push({
        sourceColumnIndex,
        targetColumnIndex,
        sourceLabel: sourceHeaderNames[sourceColumnIndex],
        targetLabel
      });
    }
  });

  return mappings;
}

function headersMatch(sourceHeader: string, targetHeader: string) {
  const source = normalizeHeader(sourceHeader);
  const target = normalizeHeader(targetHeader);
  if (!source || !target) {
    return false;
  }
  if (source === target) {
    return true;
  }

  const sourceAliases = HEADER_ALIASES[source] ?? [];
  const targetAliases = HEADER_ALIASES[target] ?? [];
  return sourceAliases.includes(target) || targetAliases.includes(source) || sourceAliases.some((alias) => targetAliases.includes(alias));
}

function resolveMappedValue({
  mapping,
  sourceRow,
  sourceExcelRowNumber,
  convertCountryToIso2
}: {
  mapping: ColumnMapping;
  sourceRow: unknown[];
  sourceExcelRowNumber: number;
  convertCountryToIso2: boolean;
}) {
  if (mapping.generated === "source-row") {
    return sourceExcelRowNumber;
  }
  if (mapping.generated === "validation-note") {
    return "Mapped deterministically";
  }

  const rawValue = mapping.sourceColumnIndex === null ? "" : sourceRow[mapping.sourceColumnIndex];
  if (convertCountryToIso2 && normalizeHeader(mapping.targetLabel).includes("country")) {
    return COUNTRY_TO_ISO2[String(rawValue || "").trim().toLowerCase()] ?? rawValue;
  }
  return rawValue ?? "";
}

function sheetToRows(sheet: XLSX.WorkSheet) {
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false,
    defval: ""
  }) as unknown[][];
}

function detectHeaderRowIndex(rows: unknown[][]) {
  const index = rows.findIndex((row) => row.filter((cell) => String(cell || "").trim() !== "").length >= 2);
  return index >= 0 ? index : 0;
}

function findSheetName(workbook: XLSX.WorkBook, requestedName: string) {
  return workbook.SheetNames.find((sheetName) => sheetName.toLowerCase() === requestedName.toLowerCase());
}

function writeCell(sheet: XLSX.WorkSheet, rowIndex: number, columnIndex: number, value: unknown) {
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  sheet[cellRef] = makeCell(value);
  const existingRange = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  existingRange.e.r = Math.max(existingRange.e.r, rowIndex);
  existingRange.e.c = Math.max(existingRange.e.c, columnIndex);
  sheet["!ref"] = XLSX.utils.encode_range(existingRange);
}

function makeCell(value: unknown) {
  if (typeof value === "number") {
    return { t: "n", v: value };
  }
  if (typeof value === "boolean") {
    return { t: "b", v: value };
  }
  if (value instanceof Date) {
    return { t: "d", v: value };
  }
  return { t: "s", v: String(value ?? "") };
}

function columnNameToIndex(columnName: string) {
  return columnName
    .toUpperCase()
    .split("")
    .reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function normalizeHeader(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function matchFirst(prompt: string, pattern: RegExp) {
  const value = prompt.match(pattern)?.[1]?.trim();
  return value || null;
}

function uniqueTargetMappings(mappings: ColumnMapping[]) {
  const seen = new Set<number>();
  return mappings.filter((mapping) => {
    if (seen.has(mapping.targetColumnIndex)) {
      return false;
    }
    seen.add(mapping.targetColumnIndex);
    return true;
  });
}
