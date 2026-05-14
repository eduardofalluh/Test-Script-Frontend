import { NextResponse } from "next/server";
import { MAX_FILE_SIZE_BYTES, MAX_TOTAL_PAYLOAD_BYTES } from "@/lib/constants";
import { mapWorkbookDeterministically } from "@/lib/deterministic-mapper";
import type { InvokeAgentResponse } from "@/lib/types";
import { estimatedBase64Size, sanitizeFilename } from "@/lib/utils";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const template = formData.get("template");
    const sources = formData.getAll("sources");
    const prompt = getString(formData.get("prompt")).trim();

    if (!(template instanceof File)) {
      return badRequest("Upload one target Excel template.");
    }
    if (!prompt) {
      return badRequest("Mapping prompt is required.");
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

    const result = mapWorkbookDeterministically({
      template: {
        filename: sanitizeFilename(template.name),
        buffer: Buffer.from(await template.arrayBuffer())
      },
      sources: await Promise.all(
        sourceFiles.map(async (file) => ({
          filename: sanitizeFilename(file.name),
          buffer: Buffer.from(await file.arrayBuffer())
        }))
      ),
      prompt
    });

    return NextResponse.json<InvokeAgentResponse>({
      ok: true,
      status: 200,
      responseText: result.responseText,
      rawResponse: result.rawResponse,
      file: result.file
    });
  } catch (error) {
    return NextResponse.json<InvokeAgentResponse>(
      {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to map workbook deterministically."
      },
      { status: 400 }
    );
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
