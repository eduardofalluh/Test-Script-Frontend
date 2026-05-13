import { NextResponse } from "next/server";
import {
  AGENT_ENDPOINT,
  EXCEL_MIME_TYPE,
  MAX_FILE_SIZE_BYTES,
  MAX_TOTAL_PAYLOAD_BYTES,
  REQUEST_TIMEOUT_MS
} from "@/lib/constants";
import { parseAgentResponseFile, responseToText } from "@/lib/agent-response";
import { estimatedBase64Size, sanitizeFilename } from "@/lib/utils";
import type { InvokeAgentResponse } from "@/lib/types";

export const runtime = "nodejs";

type AgentInput =
  | {
      type: "text";
      text: string;
    }
  | {
      type: "image_url";
      image_url: {
        url: string;
      };
    };

type PreparedFile = {
  filename: string;
  role: "TARGET TEMPLATE" | "SOURCE DATA";
  dataUrl: string;
  size: number;
};

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
      return badRequest("Syntax GenAI Studio API key is required.");
    }
    if (!sessionId) {
      return badRequest("Session ID is required.");
    }

    const sourceFiles = sources.filter((source): source is File => source instanceof File);
    const allFiles = [template, ...sourceFiles];
    const totalEstimatedPayloadSize = allFiles.reduce((sum, file) => sum + estimatedBase64Size(file.size), 0);

    for (const file of allFiles) {
      validateExcelFile(file);
    }
    if (totalEstimatedPayloadSize > MAX_TOTAL_PAYLOAD_BYTES) {
      return badRequest(`Estimated encoded payload is too large. Limit is ${MAX_TOTAL_PAYLOAD_BYTES} bytes.`);
    }

    const preparedFiles: PreparedFile[] = [
      await prepareFile(template, "TARGET TEMPLATE"),
      ...(await Promise.all(sourceFiles.map((file) => prepareFile(file, "SOURCE DATA"))))
    ];

    const textPrompt = buildAgentPrompt({
      userPrompt: prompt,
      files: preparedFiles
    });

    const input: AgentInput[] = [
      {
        type: "text",
        text: textPrompt
      },
      ...preparedFiles.map((file) => ({
        type: "image_url" as const,
        image_url: {
          url: file.dataUrl
        }
      }))
    ];

    const response = await fetchWithTimeout(
      getAgentEndpoint(),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey
        },
        body: JSON.stringify({
          input,
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

    const file = parseAgentResponseFile(rawResponse);
    return NextResponse.json<InvokeAgentResponse>({
      ok: true,
      status: response.status,
      responseText: responseToText(rawResponse),
      rawResponse,
      file
    });
  } catch (error) {
    return NextResponse.json<InvokeAgentResponse>(
      {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to invoke agent."
      },
      { status: 500 }
    );
  }
}

function getAgentEndpoint() {
  // Test and self-hosted deployments can override the endpoint without
  // changing the production default requested for Test Script IQ.
  return process.env.SYNTAX_AGENT_ENDPOINT || AGENT_ENDPOINT;
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

async function prepareFile(file: File, role: PreparedFile["role"]): Promise<PreparedFile> {
  const buffer = Buffer.from(await file.arrayBuffer());
  const filename = sanitizeFilename(file.name);

  // Assumption: the Syntax agent accepts data URLs in image_url.url even for
  // workbook content. TODO: if it rejects non-image data URLs, switch this to
  // short-lived presigned URLs using the TemporaryFileHost stub in lib/file-hosting.ts.
  const dataUrl = `data:${EXCEL_MIME_TYPE};base64,${buffer.toString("base64")}`;

  return {
    filename,
    role,
    dataUrl,
    size: file.size
  };
}

function buildAgentPrompt({ userPrompt, files }: { userPrompt: string; files: PreparedFile[] }) {
  const attachedFiles = files.map((file) => `- ${file.filename} (${file.role}, ${file.size} bytes)`).join("\n");

  return [
    "You are receiving a target Excel template and source data files. Read the template structure first, then apply the user's mapping instructions to populate it. Return the populated file along with a validation summary.",
    "",
    "User mapping instructions:",
    userPrompt,
    "",
    "Attached files:",
    attachedFiles
  ].join("\n");
}

async function fetchWithTimeout(url: string, init: RequestInit, timeoutMs: number) {
  const controller = new AbortController();
  const timeout = windowlessSetTimeout(() => controller.abort(), timeoutMs);

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

function windowlessSetTimeout(callback: () => void, ms: number) {
  return setTimeout(callback, ms);
}

function parseJsonOrText(text: string) {
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}
