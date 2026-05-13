export type WorkbookPreview = {
  sheetNames: string[];
  firstSheetName: string;
  firstSheetRows: unknown[][];
  firstSheetRowCount: number;
};

export type EditableWorkbookSheet = {
  name: string;
  rows: string[][];
  rowCount: number;
  columnCount: number;
};

export type EditableWorkbook = {
  sheets: EditableWorkbookSheet[];
  activeSheetName: string;
};

export type UploadedWorkbook = {
  id: string;
  file: File;
  preview: WorkbookPreview;
  warning?: string;
};

export type AgentFilePayload = {
  filename: string;
  mimeType: string;
  base64: string;
};

export type InvokeAgentSuccess = {
  ok: true;
  status: number;
  responseText: string;
  rawResponse: unknown;
  file: AgentFilePayload | null;
};

export type InvokeAgentError = {
  ok: false;
  error: string;
  details?: string;
};

export type InvokeAgentResponse = InvokeAgentSuccess | InvokeAgentError;

export type SavedPrompt = {
  id: string;
  title: string;
  body: string;
  updatedAt: string;
};
