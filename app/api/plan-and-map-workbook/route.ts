import { NextResponse } from "next/server";
import * as XLSX from "@e965/xlsx";
import { AGENT_ENDPOINT, MAX_FILE_SIZE_BYTES, MAX_TOTAL_PAYLOAD_BYTES, REQUEST_TIMEOUT_MS } from "@/lib/constants";
import { mapWorkbookFromStructuredSpec, type StructuredMappingSpec, type WorkbookInput } from "@/lib/deterministic-mapper";
import type { InvokeAgentResponse } from "@/lib/types";
import { estimatedBase64Size, sanitizeFilename } from "@/lib/utils";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const template = formData.get("template");
    const sources = formData.getAll("sources");
    const prompt = getString(formData.get("prompt")).trim();
    const apiKey = getString(formData.get("apiKey")).trim();
    const sessionId = getString(formData.get("sessionId")).trim();

    if (!(template instanceof File)) {
      return badRequest("Upload one target Excel template.");
    }
    if (!prompt) {
      return badRequest("Mapping prompt is required.");
    }
    if (!apiKey) {
      return badRequest("Syntax GenAI Studio API key is required for AI-assisted planning.");
    }
    if (!sessionId) {
      return badRequest("Session ID is required.");
    }

    const sourceFiles = sources.filter((source): source is File => source instanceof File);
    if (sourceFiles.length === 0) {
      return badRequest("Upload at least one source Excel workbook.");
    }

    const allFiles = [template, ...sourceFiles];
    const totalEstimatedPayloadSize = allFiles.reduce((sum, file) => sum + estimatedBase64Size(file.size), 0);
    for (const file of allFiles) {
      validateExcelFile(file);
    }
    if (totalEstimatedPayloadSize > MAX_TOTAL_PAYLOAD_BYTES) {
      return badRequest(`Estimated encoded payload is too large. Limit is ${MAX_TOTAL_PAYLOAD_BYTES} bytes.`);
    }

    const templateInput: WorkbookInput = {
      filename: sanitizeFilename(template.name),
      buffer: Buffer.from(await template.arrayBuffer())
    };
    const sourceInputs: WorkbookInput[] = await Promise.all(
      sourceFiles.map(async (file) => ({
        filename: sanitizeFilename(file.name),
        buffer: Buffer.from(await file.arrayBuffer())
      }))
    );

    const planningPrompt = buildPlanningPrompt({
      userPrompt: prompt,
      templateSummary: summarizeWorkbook(templateInput),
      sourceSummaries: sourceInputs.map(summarizeWorkbook)
    });

    const response = await fetchWithTimeout(
      getAgentEndpoint(),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey
        },
        body: JSON.stringify({
          input: [{ type: "text", text: planningPrompt }],
          session_id: sessionId
        })
      },
      REQUEST_TIMEOUT_MS
    );

    const responseText = await response.text();
    const rawResponse = parseJsonOrText(responseText);
    if (!response.ok) {
      return NextResponse.json<InvokeAgentResponse>(
        {
          ok: false,
          error: `Agent API returned ${response.status} ${response.statusText}`,
          details: responseText
        },
        { status: response.status }
      );
    }

    const spec = extractStructuredMappingSpec(rawResponse);
    const result = mapWorkbookFromStructuredSpec(
      {
        template: templateInput,
        sources: sourceInputs,
        prompt
      },
      spec
    );

    return NextResponse.json<InvokeAgentResponse>({
      ok: true,
      status: 200,
      responseText: `${result.responseText}\n\nMapping plan:\n${JSON.stringify(spec, null, 2)}`,
      rawResponse: {
        plannerResponse: rawResponse,
        execution: result.rawResponse
      },
      file: result.file
    });
  } catch (error) {
    return NextResponse.json<InvokeAgentResponse>(
      {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to plan and map workbook."
      },
      { status: 400 }
    );
  }
}

function buildPlanningPrompt({
  userPrompt,
  templateSummary,
  sourceSummaries
}: {
  userPrompt: string;
  templateSummary: unknown;
  sourceSummaries: unknown[];
}) {
  return [
    "You are a mapping planner. Return JSON only. Do not return a populated file.",
    "Your job is to convert the user's mapping request into this JSON schema:",
    "{",
    '  "sourceSheetName": "Customers",',
    '  "targetSheetName": "Migration Input",',
    '  "targetStartRow": 5,',
    '  "mappings": [',
    '    { "sourceColumn": "A", "targetColumn": "C" },',
    '    { "sourceHeader": "Contact Email", "targetHeader": "Email" },',
    '    { "generated": "source-row", "targetHeader": "Source Row" },',
    '    { "generated": "validation-note", "targetHeader": "Validation Notes" },',
    '    { "constantValue": "In Preparation", "targetHeader": "Test Case Status" }',
    "  ],",
    '  "transformations": { "convertCountryToIso2": true, "clearTargetRowsBeforeMapping": false }',
    "}",
    "Use only sheets and headers present in the workbook summaries. The deterministic mapper will execute your plan.",
    "For SAP Cloud ALM Test Script mode, targetSheetName must be \"Test Cases\", targetStartRow must be 2, and transformations.clearTargetRowsBeforeMapping must be true so the sample rows are removed before writing. Use constantValue mappings for SAP defaults such as Test Case Status = In Preparation or Test Case Priority = Medium when the source does not provide those fields.",
    "",
    "User mapping request:",
    userPrompt,
    "",
    "Template workbook summary:",
    JSON.stringify(templateSummary, null, 2),
    "",
    "Source workbook summaries:",
    JSON.stringify(sourceSummaries, null, 2)
  ].join("\n");
}

function summarizeWorkbook(input: WorkbookInput) {
  const workbook = XLSX.read(input.buffer, { type: "buffer", cellDates: true });
  return {
    filename: input.filename,
    sheets: workbook.SheetNames.map((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
        blankrows: false,
        defval: ""
      }) as unknown[][];
      const headerRowIndex = rows.findIndex((row) => row.filter((cell) => String(cell || "").trim() !== "").length >= 2);
      return {
        sheetName,
        rowCount: rows.length,
        likelyHeaders: rows[Math.max(headerRowIndex, 0)] ?? [],
        sampleRows: rows.slice(Math.max(headerRowIndex, 0) + 1, Math.max(headerRowIndex, 0) + 6)
      };
    })
  };
}

function extractStructuredMappingSpec(rawResponse: unknown): StructuredMappingSpec {
  const candidate = unwrapResponseText(rawResponse);
  const parsed = typeof candidate === "string" ? parseJsonFromText(candidate) : candidate;
  if (!parsed || typeof parsed !== "object") {
    throw new Error("AI planner did not return a JSON mapping spec.");
  }

  const spec = parsed as StructuredMappingSpec;
  if (!spec.sourceSheetName || !spec.targetSheetName || !Array.isArray(spec.mappings)) {
    throw new Error("AI planner JSON is missing sourceSheetName, targetSheetName, or mappings.");
  }
  return spec;
}

function unwrapResponseText(rawResponse: unknown) {
  if (typeof rawResponse === "string") {
    return rawResponse;
  }
  if (rawResponse && typeof rawResponse === "object") {
    const record = rawResponse as Record<string, unknown>;
    return record.mapping_spec ?? record.mappingSpec ?? record.message ?? record.output ?? record.response ?? rawResponse;
  }
  return rawResponse;
}

function parseJsonFromText(text: string) {
  const trimmed = text.trim();
  try {
    return JSON.parse(trimmed);
  } catch {
    const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
    if (fenced) {
      return JSON.parse(fenced);
    }
    const objectText = trimmed.match(/\{[\s\S]*\}/)?.[0];
    if (objectText) {
      return JSON.parse(objectText);
    }
    throw new Error("AI planner response did not contain parseable JSON.");
  }
}

function getAgentEndpoint() {
  return process.env.SYNTAX_AGENT_ENDPOINT || AGENT_ENDPOINT;
}

async function fetchWithTimeout(url: string, init: RequestInit, timeoutMs: number) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, {
      ...init,
      signal: controller.signal
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error("Timed out waiting for Test Script IQ after 5 minutes.");
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function badRequest(error: string) {
  return NextResponse.json<InvokeAgentResponse>({ ok: false, error }, { status: 400 });
}

function getString(value: FormDataEntryValue | null) {
  return typeof value === "string" ? value : "";
}

function validateExcelFile(file: File) {
  const lowerName = file.name.toLowerCase();
  if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".xlsm")) {
    throw new Error(`${file.name} is not a supported .xlsx or .xlsm workbook.`);
  }
  if (file.size > MAX_FILE_SIZE_BYTES) {
    throw new Error(`${file.name} exceeds the 50 MB file limit.`);
  }
}

function parseJsonOrText(text: string) {
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}
