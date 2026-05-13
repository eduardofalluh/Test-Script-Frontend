export const AGENT_ENDPOINT =
  "https://studio-api.ai.syntax-rnd.com/api/v1/agents/6d310742-9d0a-4069-8689-6c8feb61b935/invoke";
export const AGENT_NAME = "Test Script IQ";
export const REQUEST_TIMEOUT_MS = 300_000;
export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024;
export const WARN_FILE_SIZE_BYTES = 10 * 1024 * 1024;
export const MAX_TOTAL_PAYLOAD_BYTES = 50 * 1024 * 1024;

export const EXCEL_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export const ACCEPTED_EXCEL_EXTENSIONS = [".xlsx", ".xlsm"] as const;
