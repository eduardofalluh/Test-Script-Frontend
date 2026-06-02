import * as XLSX from "@e965/xlsx";
import { EXCEL_MIME_TYPE } from "@/lib/constants";
import { sanitizeFilename } from "@/lib/utils";
import type { AgentFilePayload } from "@/lib/types";

export type WorkbookInput = {
  filename: string;
  buffer: Buffer;
};

export type DeterministicMappingInput = {
  template: WorkbookInput;
  sources: WorkbookInput[];
  prompt: string;
};

export type ColumnMapping = {
  sourceColumnIndex: number | null;
  targetColumnIndex: number;
  sourceLabel: string;
  targetLabel: string;
  generated?: "source-row" | "validation-note";
  constantValue?: string | number | boolean | null;
};

export type MappingRule = {
  sourceSheetName: string;
  targetSheetName: string;
  targetStartRow: number;
  mappings: ColumnMapping[];
  convertCountryToIso2: boolean;
  clearTargetRowsBeforeMapping: boolean;
};

export type StructuredMappingSpec = {
  sourceSheetName: string;
  targetSheetName: string;
  targetStartRow: number;
  mappings: Array<{
    sourceColumn?: string;
    targetColumn?: string;
    sourceHeader?: string;
    targetHeader?: string;
    generated?: "source-row" | "validation-note";
    constantValue?: string | number | boolean | null;
  }>;
  transformations?: {
    convertCountryToIso2?: boolean;
    clearTargetRowsBeforeMapping?: boolean;
  };
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

  let templateWorkbook: XLSX.WorkBook;
  try {
    templateWorkbook = XLSX.read(input.template.buffer, { type: "buffer", cellDates: true, cellStyles: true });
    if (!templateWorkbook.SheetNames || templateWorkbook.SheetNames.length === 0) {
      throw new Error("Template workbook has no sheets. Please upload a valid Excel file.");
    }
  } catch (error) {
    throw new Error(`Failed to read template workbook: ${error instanceof Error ? error.message : "Unknown error"}`);
  }

  const sourceWorkbooks = input.sources.map((source) => {
    try {
      const workbook = XLSX.read(source.buffer, { type: "buffer", cellDates: true });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error(`Source file "${source.filename}" has no sheets.`);
      }
      return { filename: source.filename, workbook };
    } catch (error) {
      throw new Error(`Failed to read source file "${source.filename}": ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  });

  const rule = parseMappingRule(input.prompt, templateWorkbook, sourceWorkbooks);
  return executeMapping({
    input,
    templateWorkbook,
    sourceWorkbooks,
    rule,
    responseHeader: "Deterministic mapper completed without using the external AI agent.",
    rawMode: "deterministic"
  });
}

export function mapWorkbookFromStructuredSpec(input: DeterministicMappingInput, spec: StructuredMappingSpec) {
  if (input.sources.length === 0) {
    throw new Error("At least one source workbook is required for deterministic mapping.");
  }

  let templateWorkbook: XLSX.WorkBook;
  try {
    templateWorkbook = XLSX.read(input.template.buffer, { type: "buffer", cellDates: true, cellStyles: true });
    if (!templateWorkbook.SheetNames || templateWorkbook.SheetNames.length === 0) {
      throw new Error("Template workbook has no sheets. Please upload a valid Excel file.");
    }
  } catch (error) {
    throw new Error(`Failed to read template workbook: ${error instanceof Error ? error.message : "Unknown error"}`);
  }

  const sourceWorkbooks = input.sources.map((source) => {
    try {
      const workbook = XLSX.read(source.buffer, { type: "buffer", cellDates: true });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error(`Source file "${source.filename}" has no sheets.`);
      }
      return { filename: source.filename, workbook };
    } catch (error) {
      throw new Error(`Failed to read source file "${source.filename}": ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  });
  const rule = structuredSpecToMappingRule(spec, templateWorkbook, sourceWorkbooks);

  return executeMapping({
    input,
    templateWorkbook,
    sourceWorkbooks,
    rule,
    responseHeader: "AI-assisted mapping plan created by Test Script IQ. Workbook populated by deterministic mapper.",
    rawMode: "ai-assisted-deterministic",
    plan: spec
  });
}

function executeMapping({
  input,
  templateWorkbook,
  sourceWorkbooks,
  rule,
  responseHeader,
  rawMode,
  plan
}: {
  input: DeterministicMappingInput;
  templateWorkbook: XLSX.WorkBook;
  sourceWorkbooks: Array<{ filename: string; workbook: XLSX.WorkBook }>;
  rule: MappingRule;
  responseHeader: string;
  rawMode: string;
  plan?: StructuredMappingSpec;
}) {
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

  // SAP CALM specific: Set up proper CALM export format
  if (rule.targetSheetName === "Test Cases") {
    // Clear existing content
    const range = XLSX.utils.decode_range(targetSheet["!ref"] || "A1:M1");

    // Row 1: #
    targetSheet["A1"] = { t: "s", v: "# ", w: "# ", h: "# " };

    // Row 2: # with warning message
    targetSheet["A2"] = { t: "s", v: "# ", w: "# ", h: "# " };
    targetSheet["B2"] = {
      t: "s",
      v: "Any changes you make to the values of columns marked with [---] will be ignored.",
      w: "Any changes you make to the values of columns marked with [---] will be ignored.",
      h: "Any changes you make to the values of columns marked with [---] will be ignored."
    };

    // Row 3: # created at with current timestamp
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const timeStr = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
    const timestamp = `${year}-${month}-${day}, ${timeStr}`;
    targetSheet["A3"] = { t: "s", v: "# created at", w: "# created at", h: "# created at" };
    targetSheet["B3"] = { t: "s", v: timestamp, w: timestamp, h: timestamp };

    // Row 4: Blank (already cleared)

    // Row 5 will be the header row - we'll set this up before mapping
  }

  const sourceRows = sheetToRows(sourceSheet);
  const sourceHeaderRowIndex = detectHeaderRowIndex(sourceRows);
  const sourceDataRows = sourceRows.slice(sourceHeaderRowIndex + 1);

  // Filter out empty rows and header-like rows (rows with mostly text that matches headers)
  const mappedRows = sourceDataRows.filter((row) => {
    // Must have at least one non-empty cell
    if (!row.some((cell) => cell !== "")) {
      return false;
    }

    // For SAP CALM: Skip rows that look like header rows (contain asterisks, brackets, or all-caps)
    if (rule.targetSheetName === "Test Cases") {
      // Check first cell
      const firstCell = String(row[0] || "").trim();

      // Skip timestamp rows - check if it looks like a date/time
      // Patterns: "6/2/2026", "6/2/2026, 8:41:41", "8:41:41 AM", etc.
      if (
        /\d{1,2}\/\d{1,2}\/\d{4}/.test(firstCell) || // Date pattern
        /\d{1,2}:\d{1,2}(:\d{1,2})?/.test(firstCell) || // Time pattern
        /^\d{1,2}\/\d{1,2}\/\d{2,4}[,\s]/.test(firstCell) // Date followed by comma/space
      ) {
        console.log(`[CALM Filter] Skipping timestamp row: ${firstCell}`);
        return false;
      }

      // Skip if first cell contains decorators like asterisks or square brackets
      if (firstCell.includes("*") || /^\[.*\]$/.test(firstCell)) {
        console.log(`[CALM Filter] Skipping header-like row: ${firstCell}`);
        return false;
      }

      // Skip rows where multiple cells contain asterisks or brackets (likely a header row)
      const cellsWithSymbols = row.filter((cell) => {
        const str = String(cell || "");
        return str.includes("*") || /^\[.*\]$/.test(str);
      });
      if (cellsWithSymbols.length >= 3) {
        console.log(`[CALM Filter] Skipping multi-symbol header row`);
        return false;
      }

      // Skip rows where first cell is very short and looks generic (like just a number)
      if (firstCell.length > 0 && firstCell.length < 3 && /^\d+$/.test(firstCell)) {
        console.log(`[CALM Filter] Skipping generic numeric row: ${firstCell}`);
        return false;
      }
    }

    return true;
  });

  if (rule.clearTargetRowsBeforeMapping) {
    clearTargetRows(targetSheet, rule.targetStartRow - 1);
  }

  // SAP CALM specific: Set up row 5 with proper CALM headers
  if (rule.targetSheetName === "Test Cases") {
    // CALM requires specific column order and naming
    const calmHeaders = [
      "Test Case GUID",
      "Test Case Name*",
      "[Scope GUID]",
      "[Scope Name]",
      "[Solution Process GUID]",
      "[Solution Process Name]",
      "Test Case Status",
      "Test Case Priority",
      "Test Case References",
      "Test Case Owner",
      "Tag",
      "Activity Title",
      "Activity Target Name",
      "Activity Target URL",
      "Action Title",
      "Action Instructions",
      "Action Expected Result",
      "Action Evidence"
    ];

    // Write headers to row 5
    calmHeaders.forEach((header, colIndex) => {
      const cellAddr = XLSX.utils.encode_cell({ r: 4, c: colIndex }); // Row 5 = index 4
      targetSheet[cellAddr] = {
        t: "s",
        v: header,
        w: header,
        h: header
      };
    });

    // Update range
    const range = XLSX.utils.decode_range(targetSheet["!ref"] || "A1:A1");
    range.e.c = Math.max(range.e.c, calmHeaders.length - 1);
    range.e.r = Math.max(range.e.r, 4);
    targetSheet["!ref"] = XLSX.utils.encode_range(range);
  }

  mappedRows.forEach((sourceRow, rowOffset) => {
    // SAP CALM specific: Data starts at row 6 (index 5)
    const actualTargetRowIndex = rule.targetSheetName === "Test Cases"
      ? 5 + rowOffset // Row 6 = index 5
      : rule.targetStartRow - 1 + rowOffset;

    rule.mappings.forEach((mapping) => {
      let value = resolveMappedValue({
        mapping,
        sourceRow,
        sourceExcelRowNumber: sourceHeaderRowIndex + rowOffset + 2,
        convertCountryToIso2: rule.convertCountryToIso2
      });

      // SAP CALM specific: Map to correct column indices
      // The mapping targetColumnIndex needs to be adjusted for CALM column structure
      // Test Case Name* is now at index 1 (column B) instead of index 0
      let targetColIndex = mapping.targetColumnIndex;
      if (rule.targetSheetName === "Test Cases") {
        // Shift regular columns to account for GUID columns at the start
        if (mapping.targetLabel === "Test Case Name") {
          targetColIndex = 1; // Column B: Test Case Name*
        } else if (mapping.targetLabel === "Test Case Status") {
          targetColIndex = 6;
        } else if (mapping.targetLabel === "Test Case Priority") {
          targetColIndex = 7;
        } else if (mapping.targetLabel === "Test Case References") {
          targetColIndex = 8;
        } else if (mapping.targetLabel === "Test Case Owner") {
          targetColIndex = 9;
        } else if (mapping.targetLabel === "Tag") {
          targetColIndex = 10;
        } else if (mapping.targetLabel === "Activity Title") {
          targetColIndex = 11;
        } else if (mapping.targetLabel === "Activity Target Name") {
          targetColIndex = 12;
        } else if (mapping.targetLabel === "Activity Target URL") {
          targetColIndex = 13;
        } else if (mapping.targetLabel === "Action Title") {
          targetColIndex = 14;
        } else if (mapping.targetLabel === "Action Instructions") {
          targetColIndex = 15;
        } else if (mapping.targetLabel === "Action Expected Result") {
          targetColIndex = 16;
        } else if (mapping.targetLabel === "Action Evidence") {
          targetColIndex = 17;
        }
      }

      writeCell(targetSheet, actualTargetRowIndex, targetColIndex, value);
    });
  });

  const base64 = XLSX.write(templateWorkbook, {
    type: "base64",
    bookType: "xlsx",
    cellStyles: true
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
      responseHeader,
      "",
      `Source workbook: ${sourceWorkbook.filename}`,
      `Source sheet: ${resolvedSourceSheetName}`,
      `Target sheet: ${resolvedTargetSheetName}`,
      `Target start row: ${rule.targetStartRow}`,
      `Rows mapped: ${mappedRows.length}`,
      `Column mappings: ${rule.mappings.map((mapping) => `${mapping.sourceLabel}->${mapping.targetLabel}`).join(", ")}`,
      [
        rule.convertCountryToIso2 ? "country names converted to ISO-2 where recognized" : null,
        rule.clearTargetRowsBeforeMapping ? "target sample rows cleared before mapping" : null
      ].filter(Boolean).length > 0
        ? `Transformations: ${[
            rule.convertCountryToIso2 ? "country names converted to ISO-2 where recognized" : null,
            rule.clearTargetRowsBeforeMapping ? "target sample rows cleared before mapping" : null
          ]
            .filter(Boolean)
            .join("; ")}.`
        : "Transformations: none."
    ].join("\n"),
    rawResponse: {
      mode: rawMode,
      rule,
      rowsMapped: mappedRows.length,
      plan
    }
  };
}

function structuredSpecToMappingRule(
  spec: StructuredMappingSpec,
  templateWorkbook: XLSX.WorkBook,
  sources: Array<{ filename: string; workbook: XLSX.WorkBook }>
): MappingRule {
  if (!spec.sourceSheetName || !spec.targetSheetName || !spec.targetStartRow) {
    throw new Error("AI mapping plan is missing sourceSheetName, targetSheetName, or targetStartRow.");
  }

  const sourceWorkbook = sources.find((source) => findSheetName(source.workbook, spec.sourceSheetName)) ?? sources[0];
  const resolvedSourceSheetName = findSheetName(sourceWorkbook.workbook, spec.sourceSheetName);
  const resolvedTargetSheetName = findSheetName(templateWorkbook, spec.targetSheetName);
  if (!resolvedSourceSheetName || !resolvedTargetSheetName) {
    throw new Error("AI mapping plan references a source or target sheet that was not found.");
  }

  const sourceRows = sheetToRows(sourceWorkbook.workbook.Sheets[resolvedSourceSheetName]);
  const targetRows = sheetToRows(templateWorkbook.Sheets[resolvedTargetSheetName]);
  const sourceHeaders = sourceRows[detectHeaderRowIndex(sourceRows)] ?? [];
  const targetHeaders = targetRows[spec.targetStartRow - 2] ?? [];

  const mappings = spec.mappings.map((mapping) => resolveStructuredColumnMapping(mapping, sourceHeaders, targetHeaders));
  if (mappings.length === 0) {
    throw new Error("AI mapping plan did not include any executable column mappings.");
  }

  return {
    sourceSheetName: spec.sourceSheetName,
    targetSheetName: spec.targetSheetName,
    targetStartRow: spec.targetStartRow,
    mappings: uniqueTargetMappings(mappings),
    convertCountryToIso2: Boolean(spec.transformations?.convertCountryToIso2),
    clearTargetRowsBeforeMapping: Boolean(spec.transformations?.clearTargetRowsBeforeMapping)
  };
}

function resolveStructuredColumnMapping(
  mapping: StructuredMappingSpec["mappings"][number],
  sourceHeaders: unknown[],
  targetHeaders: unknown[]
): ColumnMapping {
  const targetColumnIndex = mapping.targetColumn
    ? columnNameToIndex(mapping.targetColumn)
    : findHeaderIndex(targetHeaders, mapping.targetHeader ?? "");

  if (targetColumnIndex < 0) {
    const requestedName = mapping.targetColumn ?? mapping.targetHeader ?? "unknown";
    const availableHeaders = targetHeaders
      .map((h, i) => `${String(h)} (column ${XLSX.utils.encode_col(i)})`)
      .slice(0, 10)
      .join(", ");
    throw new Error(
      `AI mapping plan target column could not be resolved: ${requestedName}. ` +
      `Available target columns: ${availableHeaders}. ` +
      `This usually means the AI added symbols (* or parentheses) to the column name. ` +
      `Try re-prompting with the exact error from CALM.`
    );
  }

  if (mapping.generated) {
    return {
      sourceColumnIndex: null,
      targetColumnIndex,
      sourceLabel: mapping.generated,
      targetLabel: String(targetHeaders[targetColumnIndex] || mapping.targetHeader || mapping.targetColumn || ""),
      generated: mapping.generated
    };
  }

  if (Object.prototype.hasOwnProperty.call(mapping, "constantValue")) {
    return {
      sourceColumnIndex: null,
      targetColumnIndex,
      sourceLabel: `constant ${String(mapping.constantValue ?? "")}`,
      targetLabel: String(targetHeaders[targetColumnIndex] || mapping.targetHeader || mapping.targetColumn || ""),
      constantValue: mapping.constantValue
    };
  }

  const sourceColumnIndex = mapping.sourceColumn
    ? columnNameToIndex(mapping.sourceColumn)
    : findHeaderIndex(sourceHeaders, mapping.sourceHeader ?? "");

  if (sourceColumnIndex < 0) {
    const requestedName = mapping.sourceColumn ?? mapping.sourceHeader ?? "unknown";
    const availableHeaders = sourceHeaders
      .map((h, i) => `${String(h)} (column ${XLSX.utils.encode_col(i)})`)
      .slice(0, 10)
      .join(", ");
    throw new Error(
      `AI mapping plan source column could not be resolved: ${requestedName}. ` +
      `Available source columns: ${availableHeaders}. ` +
      `Check your source file has this column.`
    );
  }

  return {
    sourceColumnIndex,
    targetColumnIndex,
    sourceLabel: String(sourceHeaders[sourceColumnIndex] || mapping.sourceHeader || mapping.sourceColumn || ""),
    targetLabel: String(targetHeaders[targetColumnIndex] || mapping.targetHeader || mapping.targetColumn || "")
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
    convertCountryToIso2: /iso-?2|country names?\s+to\s+iso/i.test(prompt),
    clearTargetRowsBeforeMapping: /clear (?:the )?(?:existing |sample )?(?:target )?rows|sap\s+(?:cloud\s+alm|calm)/i.test(prompt)
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
  if (Object.prototype.hasOwnProperty.call(mapping, "constantValue")) {
    return mapping.constantValue ?? "";
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

function clearTargetRows(sheet: XLSX.WorkSheet, startRowIndex: number) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let rowIndex = startRowIndex; rowIndex <= range.e.r; rowIndex += 1) {
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      delete sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
    }
  }

  range.e.r = Math.max(range.s.r, startRowIndex - 1);
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function detectHeaderRowIndex(rows: unknown[][]) {
  const index = rows.findIndex((row) => row.filter((cell) => String(cell || "").trim() !== "").length >= 2);
  return index >= 0 ? index : 0;
}

function findSheetName(workbook: XLSX.WorkBook, requestedName: string) {
  return workbook.SheetNames.find((sheetName) => sheetName.toLowerCase() === requestedName.toLowerCase());
}

function findHeaderIndex(headers: unknown[], requestedHeader: string) {
  return headers.findIndex((header) => headersMatch(String(header || ""), requestedHeader));
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
