import { EXCEL_MIME_TYPE } from "@/lib/constants";
import type { AgentFilePayload } from "@/lib/types";

const FILE_KEYS = ["output_file", "outputFile", "file", "result", "xlsx", "workbook"];
const XLSX_BASE64_PREFIX = "UEsDB";
const BASE64_CANDIDATE_PATTERN = /(?:data:[^;,]+;base64,)?([A-Za-z0-9+/]{80,}={0,2})/g;

type FileCandidate = {
  filename?: string;
  mimeType?: string;
  base64: string;
};

export function parseAgentResponseFile(response: unknown): AgentFilePayload | null {
  const candidates = collectCandidates(response);

  for (const candidate of candidates) {
    const normalized = normalizeBase64(candidate.base64);
    if (looksLikeXlsxBase64(normalized)) {
      return {
        filename: candidate.filename ?? "populated-workbook.xlsx",
        mimeType: candidate.mimeType ?? EXCEL_MIME_TYPE,
        base64: normalized
      };
    }
  }

  return null;
}

export function responseToText(response: unknown): string {
  if (typeof response === "string") {
    return response;
  }

  if (response && typeof response === "object") {
    const maybeText =
      getByPath(response, ["text"]) ??
      getByPath(response, ["message"]) ??
      getByPath(response, ["response"]) ??
      getByPath(response, ["output"]) ??
      getByPath(response, ["content"]);

    if (typeof maybeText === "string") {
      return maybeText;
    }
  }

  return JSON.stringify(response, null, 2);
}

function collectCandidates(value: unknown): FileCandidate[] {
  const candidates: FileCandidate[] = [];

  visit(value, (node, key) => {
    if (typeof node === "string") {
      const normalized = normalizeBase64(node);
      if (FILE_KEYS.includes(key ?? "") || looksLikeXlsxBase64(normalized)) {
        candidates.push({ base64: normalized });
      }

      let match: RegExpExecArray | null;
      BASE64_CANDIDATE_PATTERN.lastIndex = 0;
      while ((match = BASE64_CANDIDATE_PATTERN.exec(node)) !== null) {
        const matched = normalizeBase64(match[1] ?? "");
        if (looksLikeXlsxBase64(matched)) {
          candidates.push({ base64: matched });
        }
      }
    }

    if (node && typeof node === "object" && !Array.isArray(node)) {
      const record = node as Record<string, unknown>;
      const rawBase64 =
        record.base64 ??
        record.data ??
        record.content ??
        record.url ??
        record.output_file ??
        record.file;

      if (typeof rawBase64 === "string") {
        candidates.push({
          filename: typeof record.filename === "string" ? record.filename : undefined,
          mimeType: typeof record.mimeType === "string" ? record.mimeType : undefined,
          base64: rawBase64
        });
      }
    }
  });

  return candidates;
}

function visit(value: unknown, callback: (node: unknown, key?: string) => void, key?: string) {
  callback(value, key);

  if (Array.isArray(value)) {
    value.forEach((item) => visit(item, callback));
    return;
  }

  if (value && typeof value === "object") {
    Object.entries(value).forEach(([childKey, childValue]) => visit(childValue, callback, childKey));
  }
}

function getByPath(value: unknown, path: string[]) {
  let current = value;
  for (const segment of path) {
    if (!current || typeof current !== "object" || !(segment in current)) {
      return undefined;
    }
    current = (current as Record<string, unknown>)[segment];
  }
  return current;
}

function normalizeBase64(value: string) {
  const withoutDataUrl = value.includes(",") ? value.split(",").pop() ?? value : value;
  return withoutDataUrl.replace(/\s/g, "");
}

function looksLikeXlsxBase64(value: string) {
  return value.startsWith(XLSX_BASE64_PREFIX);
}
