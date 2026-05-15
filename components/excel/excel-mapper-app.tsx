"use client";

import * as React from "react";
import { AlertTriangle, KeyRound, Loader2, RefreshCw, ShieldAlert, WandSparkles, XCircle } from "lucide-react";
import { WorkbookDropzone } from "@/components/excel/workbook-dropzone";
import { ResultErrorBoundary } from "@/components/excel/result-error-boundary";
import { ResultPanel, type ResultState } from "@/components/excel/result-panel";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Toaster } from "@/components/ui/toaster";
import { ToastStateProvider, useToast } from "@/hooks/use-toast";
import { EXCEL_MIME_TYPE } from "@/lib/constants";
import { parseWorkbook, parseWorkbookFromBase64 } from "@/lib/excel";
import type { InvokeAgentResponse, SavedPrompt, UploadedWorkbook } from "@/lib/types";
import { estimatedBase64Size, formatBytes } from "@/lib/utils";

const API_KEY_STORAGE_KEY = "excel-mapper.syntax-api-key";
const SESSION_ID_STORAGE_KEY = "excel-mapper.session-id";
const PROMPT_PLACEHOLDER =
  "Take the customer data from Source.xlsx sheet 'Customers' columns A-E and map them to the template's 'Migration Input' sheet starting row 5. Map source A to target C, source B to target D, source C to target F, and convert country names to ISO-2 codes before writing them.";
const SAP_CALM_TEMPLATE_URL = "/templates/sap-cloud-alm-test-cases-template.xlsx";
const SAP_CALM_TEMPLATE_NAME = "SAP Cloud ALM - Test Cases template.xlsx";
const SAP_CALM_PROMPT = [
  "Use SAP Cloud ALM Test Script mode with the built-in SAP Cloud ALM Test Cases template.",
  "Map the uploaded source workbook into target sheet 'Test Cases' starting at row 2.",
  "Clear the sample rows before writing new rows.",
  "Use these SAP CALM target columns when mapping: Test Case Name, Test Case Status, Test Case Priority, Test Case References, Test Case Owner, Tag, Activity Title, Activity Target Name, Activity Target URL, Action Title, Action Instructions, Action Expected Result, Action Evidence.",
  "Default missing Test Case Status to In Preparation and missing Test Case Priority to Medium.",
  "",
  "User clarification:"
].join("\n");
type MappingMode = "deterministic" | "assisted" | "sap-calm" | "agent";

export function ExcelMapperApp() {
  return (
    <ToastStateProvider>
      <ExcelMapperContent />
      <Toaster />
    </ToastStateProvider>
  );
}

function ExcelMapperContent() {
  const { toast } = useToast();
  const [templateFiles, setTemplateFiles] = React.useState<UploadedWorkbook[]>([]);
  const [sourceFiles, setSourceFiles] = React.useState<UploadedWorkbook[]>([]);
  const [prompt, setPrompt] = React.useState("");
  const [apiKey, setApiKey] = React.useState("");
  const [sessionId, setSessionId] = React.useState("");
  const [mappingMode, setMappingMode] = React.useState<MappingMode>("deterministic");
  const [abortController, setAbortController] = React.useState<AbortController | null>(null);
  const [result, setResult] = React.useState<ResultState>({
    status: "idle",
    responseText: "",
    file: null,
    preview: null
  });

  React.useEffect(() => {
    const storedSessionId = window.localStorage.getItem(SESSION_ID_STORAGE_KEY);
    const nextSessionId = storedSessionId || createSessionId();
    window.localStorage.setItem(SESSION_ID_STORAGE_KEY, nextSessionId);
    setSessionId(nextSessionId);
  }, []);

  React.useEffect(() => {
    setApiKey(window.localStorage.getItem(API_KEY_STORAGE_KEY) ?? "");
  }, []);

  const totalRawBytes = [...templateFiles, ...sourceFiles].reduce((sum, uploaded) => sum + uploaded.file.size, 0);
  const estimatedRequestBytes = [...templateFiles, ...sourceFiles].reduce(
    (sum, uploaded) => sum + estimatedBase64Size(uploaded.file.size),
    0
  );

  const handleTemplateDrop = React.useCallback(
    async (files: File[]) => {
      const uploaded = await buildUploadedWorkbook(files[0]);
      setTemplateFiles(uploaded ? [uploaded] : []);
      if (uploaded) {
        toast({ title: "Template uploaded", description: uploaded.file.name });
      }
    },
    [toast]
  );

  const handleSourceDrop = React.useCallback(
    async (files: File[]) => {
      const uploaded = (await Promise.all(files.map(buildUploadedWorkbook))).filter(Boolean) as UploadedWorkbook[];
      setSourceFiles((current) => [...current, ...uploaded]);
      if (uploaded.length > 0) {
        toast({ title: "Source files uploaded", description: `${uploaded.length} workbook${uploaded.length === 1 ? "" : "s"} added.` });
      }
    },
    [toast]
  );

  const removeTemplate = React.useCallback(
    (id: string) => {
      const removed = templateFiles.find((file) => file.id === id);
      setTemplateFiles((current) => current.filter((file) => file.id !== id));
      toast({ title: "Template removed", description: removed?.file.name });
    },
    [templateFiles, toast]
  );

  const removeSource = React.useCallback(
    (id: string) => {
      const removed = sourceFiles.find((file) => file.id === id);
      setSourceFiles((current) => current.filter((file) => file.id !== id));
      toast({ title: "Source file removed", description: removed?.file.name });
    },
    [sourceFiles, toast]
  );

  const forgetApiKey = React.useCallback(() => {
    window.localStorage.removeItem(API_KEY_STORAGE_KEY);
    setApiKey("");
    toast({ title: "API key forgotten" });
  }, [toast]);

  const updateApiKey = React.useCallback((value: string) => {
    setApiKey(value);
    window.localStorage.setItem(API_KEY_STORAGE_KEY, value);
  }, []);

  const activateSapCalmMode = React.useCallback(async () => {
    if (result.status === "processing") {
      return;
    }

    setMappingMode("sap-calm");
    setPrompt((current) => ensureSapCalmPrompt(current));

    try {
      const response = await fetch(SAP_CALM_TEMPLATE_URL);
      if (!response.ok) {
        throw new Error(`Could not load SAP CALM template (${response.status}).`);
      }

      const blob = await response.blob();
      const templateFile = new File([blob], SAP_CALM_TEMPLATE_NAME, { type: EXCEL_MIME_TYPE });
      const uploaded = await buildUploadedWorkbook(templateFile);
      setTemplateFiles(uploaded ? [uploaded] : []);
      toast({
        title: "SAP CALM mode enabled",
        description: "The SAP Cloud ALM Test Cases template is ready. Upload the source workbook and add mapping clarification."
      });
    } catch (error) {
      toast({ title: "SAP CALM template failed to load", description: getErrorMessage(error) });
    }
  }, [result.status, toast]);

  const regenerateSession = React.useCallback(() => {
    const nextSessionId = createSessionId();
    window.localStorage.setItem(SESSION_ID_STORAGE_KEY, nextSessionId);
    setSessionId(nextSessionId);
    toast({ title: "Session regenerated" });
  }, [toast]);

  const cancelRequest = React.useCallback(() => {
    abortController?.abort();
    setAbortController(null);
    setResult((current) => ({
      ...current,
      status: current.status === "processing" ? "idle" : current.status
    }));
    toast({ title: "Generation canceled" });
  }, [abortController, toast]);

  const submit = React.useCallback(async (overridePrompt?: string, overrideTemplate?: File) => {
    if (result.status === "processing") {
      return;
    }

    const trimmedPrompt = (overridePrompt ?? prompt).trim();
    if (templateFiles.length === 0) {
      toast({ title: "Template required", description: "Upload the target Excel template before generating." });
      return;
    }
    if (!trimmedPrompt) {
      toast({ title: "Prompt required", description: "Describe the mapping logic before generating." });
      return;
    }
    if (mappingMode !== "deterministic" && !apiKey.trim()) {
      toast({ title: "API key required", description: "Paste your Syntax GenAI Studio API key." });
      return;
    }

    const controller = new AbortController();
    setAbortController(controller);
    setResult({
      status: "processing",
      statusMessage:
        mappingMode === "deterministic"
          ? "Running deterministic workbook mapper locally..."
          : mappingMode === "assisted"
            ? "Asking Test Script IQ for a JSON mapping plan, then executing it deterministically..."
            : mappingMode === "sap-calm"
              ? "Planning SAP CALM field mapping, then writing into the bundled template..."
              : "Processing request with Test Script IQ...",
      responseText:
        mappingMode === "deterministic"
          ? "Running deterministic workbook mapper. The external AI agent is not being used."
          : mappingMode === "assisted"
            ? "Requesting a structured mapping plan from Test Script IQ. The app will execute the Excel writes deterministically."
            : mappingMode === "sap-calm"
              ? "Using the bundled SAP Cloud ALM Test Cases template. Test Script IQ will plan the field mapping from your clarification, then code will write the workbook."
              : "Request sent to Test Script IQ. Waiting for the final response...",
      file: null,
      preview: null
    });
    toast({
      title: "Generation started",
      description:
        mappingMode === "deterministic"
          ? "Applying coded mapping rules to the uploaded workbook."
          : mappingMode === "assisted"
            ? "Using AI for planning only; deterministic code will write the workbook."
            : mappingMode === "sap-calm"
              ? "Using the SAP CALM template and your source workbook."
              : "Sending files and prompt to Test Script IQ."
    });

    const formData = new FormData();
    formData.append("template", overrideTemplate ?? templateFiles[0].file);
    sourceFiles.forEach((source) => formData.append("sources", source.file));
    formData.append("prompt", trimmedPrompt);
    if (mappingMode !== "deterministic") {
      formData.append("apiKey", apiKey.trim());
    }
    const requestSessionId = sessionId || createSessionId();
    if (!sessionId) {
      setSessionId(requestSessionId);
      window.localStorage.setItem(SESSION_ID_STORAGE_KEY, requestSessionId);
    }
    formData.append("sessionId", requestSessionId);

    try {
      const endpoint =
        mappingMode === "deterministic"
          ? "/api/map-workbook"
          : mappingMode === "assisted" || mappingMode === "sap-calm"
            ? "/api/plan-and-map-workbook"
            : "/api/invoke-agent";
      const response = await fetch(endpoint, {
        method: "POST",
        body: formData,
        signal: controller.signal
      });
      const payload = (await response.json()) as InvokeAgentResponse;

      if (!payload.ok) {
        setResult({
          status: "error",
          responseText: payload.details ?? payload.error,
          file: null,
          preview: null,
          error: payload.error,
          details: payload.details
        });
        toast({ title: "Generation failed", description: payload.error });
        return;
      }

      const preview = payload.file ? parseWorkbookFromBase64(payload.file.base64) : null;
      setResult({
        status: "success",
        responseText: payload.responseText,
        rawResponse: payload.rawResponse,
        file: payload.file,
        preview
      });
      toast({
        title: "Generation complete",
        description:
          mappingMode === "deterministic"
            ? "Workbook populated by deterministic mapper."
            : mappingMode === "assisted"
              ? "AI mapping plan executed deterministically."
              : mappingMode === "sap-calm"
                ? "SAP CALM workbook generated."
                : payload.file
                  ? "Populated workbook detected."
                  : "Agent responded without a workbook."
      });
    } catch (error) {
      const message = error instanceof DOMException && error.name === "AbortError" ? "The request was canceled." : getErrorMessage(error);
      setResult({
        status: "error",
        responseText: message,
        file: null,
        preview: null,
        error: message
      });
      toast({ title: "Generation failed", description: message });
    } finally {
      setAbortController(null);
    }
  }, [apiKey, mappingMode, prompt, result.status, sessionId, sourceFiles, templateFiles, toast]);

  const refineWithAi = React.useCallback(
    (instruction: string, editedBase64: string | null) => {
      const trimmedInstruction = instruction.trim();
      if (!trimmedInstruction) {
        return;
      }

      const usingEditedFile = editedBase64 !== null && result.file !== null;

      const revisedPrompt = [
        prompt.trim(),
        "",
        "Revision request for the generated workbook:",
        trimmedInstruction,
        "",
        usingEditedFile
          ? "The template provided is the manually edited version from the previous run. Preserve any correct mappings and only change what is needed."
          : "Use the original template and source files again. Preserve any correct mappings from the previous run and only change what is needed.",
        result.responseText ? `Previous agent response:\n${result.responseText}` : ""
      ]
        .filter(Boolean)
        .join("\n");

      setPrompt(revisedPrompt);

      if (usingEditedFile && result.file) {
        const bytes = Uint8Array.from(atob(editedBase64), (c) => c.charCodeAt(0));
        const overrideTemplate = new File([bytes], result.file.filename, { type: result.file.mimeType });
        void submit(revisedPrompt, overrideTemplate);
      } else {
        void submit(revisedPrompt);
      }
    },
    [prompt, result.file, result.responseText, submit]
  );

  const handlePromptKeyDown = React.useCallback(
    (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
        event.preventDefault();
        void submit();
      }
    },
    [submit]
  );

  const copyResponse = React.useCallback(async () => {
    await navigator.clipboard.writeText(result.responseText);
    toast({ title: "Copied to clipboard" });
  }, [result.responseText, toast]);

  return (
    <main className="mx-auto flex min-h-screen w-full max-w-7xl flex-col gap-8 px-6 py-8">
      <header className="border-b border-border pb-6">
        <p className="text-sm font-medium uppercase tracking-[0.24em] text-accent">Syntax GenAI Studio</p>
        <h1 className="mt-3 text-4xl font-semibold tracking-normal text-foreground">Excel Mapper</h1>
        <p className="mt-2 max-w-3xl text-sm text-muted-foreground">
          Upload a target template, attach source workbooks, and run deterministic cell-by-cell mapping, with Test Script IQ available only as an optional fallback.
        </p>
      </header>

      <Alert className="border-primary/40 bg-primary/10">
        <ShieldAlert className="h-4 w-4" />
        <AlertTitle>Processing notice</AlertTitle>
        <AlertDescription>
          Deterministic mapping runs locally in this app. Files are sent to Syntax GenAI Studio only when you choose External AI Agent mode. Do not upload client-confidential data unless you have authorization.
        </AlertDescription>
      </Alert>

      <section className="grid gap-4 lg:grid-cols-2">
        <WorkbookDropzone
          title="Template Upload"
          description={mappingMode === "sap-calm" ? "SAP Cloud ALM Test Cases template loaded by SAP CALM mode." : "Target workbook to be populated."}
          files={templateFiles}
          onDropAccepted={handleTemplateDrop}
          onRemove={removeTemplate}
          showPreview
        />
        <WorkbookDropzone
          title="Source File(s)"
          description="One or more workbooks containing data to map into the template."
          files={sourceFiles}
          multiple
          onDropAccepted={handleSourceDrop}
          onRemove={removeSource}
        />
      </section>

      <Card>
        <CardHeader>
          <CardTitle>Mapping Prompt</CardTitle>
          <CardDescription>Be specific about source sheet, source columns, target sheet, target start row, and transformations.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-5">
          <Textarea
            aria-label="Mapping Prompt"
            value={prompt}
            onChange={(event) => setPrompt(event.target.value)}
            onKeyDown={handlePromptKeyDown}
            placeholder={PROMPT_PLACEHOLDER}
            rows={12}
            className="min-h-72 resize-y"
          />
          <p className="text-sm text-muted-foreground">
            Tip: include the source workbook name, sheet name, target sheet, starting row, column mappings, and any formatting or validation rules.
          </p>

          <div className="space-y-3 rounded-md border border-border bg-background/40 p-4">
            <div>
              <p className="text-sm font-medium">Mapping engine</p>
              <p className="mt-1 text-sm text-muted-foreground">
                Deterministic mode uses coded SheetJS logic for cell-by-cell mapping. AI-Assisted and SAP CALM modes let Test Script IQ interpret the prompt into JSON rules, then code writes the workbook.
              </p>
            </div>
            <div className="flex flex-col gap-2 sm:flex-row" role="group" aria-label="Mapping engine">
              <Button
                type="button"
                variant={mappingMode === "deterministic" ? "default" : "outline"}
                onClick={() => setMappingMode("deterministic")}
              >
                Deterministic Mapper
              </Button>
              <Button type="button" variant={mappingMode === "assisted" ? "default" : "outline"} onClick={() => setMappingMode("assisted")}>
                AI-Assisted Deterministic
              </Button>
              <Button type="button" variant={mappingMode === "sap-calm" ? "default" : "outline"} onClick={() => void activateSapCalmMode()}>
                SAP CALM Test Script
              </Button>
              <Button type="button" variant={mappingMode === "agent" ? "default" : "outline"} onClick={() => setMappingMode("agent")}>
                External AI Agent
              </Button>
            </div>
            {mappingMode === "deterministic" ? (
              <Alert>
                <AlertDescription>
                  Supported deterministic prompts should name a source sheet, target sheet, target start row, and either explicit column mappings like A-&gt;C or matching source/target headers.
                </AlertDescription>
              </Alert>
            ) : mappingMode === "assisted" ? (
              <Alert className="border-accent/40 bg-accent/10">
                <AlertDescription>
                  AI-Assisted mode sends workbook summaries and your prompt to Test Script IQ for a JSON mapping plan. Excel writing still happens deterministically in this app.
                </AlertDescription>
              </Alert>
            ) : mappingMode === "sap-calm" ? (
              <Alert className="border-accent/40 bg-accent/10">
                <AlertDescription>
                  SAP CALM mode loads the built-in SAP Cloud ALM Test Cases template, sends workbook summaries and your clarification prompt to Test Script IQ, then writes the mapped rows into the Test Cases sheet.
                </AlertDescription>
              </Alert>
            ) : (
              <Alert className="border-amber-400/40 bg-amber-400/10">
                <AlertDescription>
                  AI Agent mode sends files to Test Script IQ and lets the agent perform the mapping. Use this only when deterministic rules are not enough.
                </AlertDescription>
              </Alert>
            )}
          </div>

          <div className="grid gap-4 lg:grid-cols-[1fr_22rem]">
            <div className="space-y-2">
              <label htmlFor="api-key" className="text-sm font-medium">
                API Key {mappingMode === "deterministic" ? <span className="text-muted-foreground">(only needed for AI and SAP CALM modes)</span> : null}
              </label>
              <div className="flex gap-2">
                <div className="relative flex-1">
                  <KeyRound className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-muted-foreground" />
                  <Input
                    id="api-key"
                    type="password"
                    value={apiKey}
                    onChange={(event) => updateApiKey(event.target.value)}
                    placeholder="Syntax GenAI Studio API key"
                    className="pl-9"
                  />
                </div>
                <Button type="button" variant="outline" onClick={forgetApiKey}>
                  Forget
                </Button>
              </div>
            </div>

            <div className="space-y-2">
              <label htmlFor="session-id" className="text-sm font-medium">
                Session ID
              </label>
              <div className="flex gap-2">
                <Input id="session-id" value={sessionId} readOnly className="font-mono text-xs" />
                <Button type="button" variant="outline" size="icon" onClick={regenerateSession} title="Regenerate session ID" aria-label="Regenerate session ID">
                  <RefreshCw className="h-4 w-4" />
                </Button>
              </div>
            </div>
          </div>

          <div className="flex flex-col gap-3 rounded-md border border-border bg-background/40 p-4 text-sm text-muted-foreground sm:flex-row sm:items-center sm:justify-between">
            <div>
              <p>Raw file size: {formatBytes(totalRawBytes)}</p>
              <p>Estimated base64 request size: {formatBytes(estimatedRequestBytes)}</p>
            </div>
            {estimatedRequestBytes > 40 * 1024 * 1024 ? (
              <div className="flex items-center gap-2 text-amber-300">
                <AlertTriangle className="h-4 w-4" />
                Approaching the 50 MB request limit.
              </div>
            ) : null}
          </div>

          <div className="flex flex-col-reverse gap-3 sm:flex-row sm:justify-end">
            <Button type="button" variant="outline" onClick={cancelRequest} disabled={result.status !== "processing"}>
              <XCircle className="h-4 w-4" />
              Cancel
            </Button>
            <Button type="button" onClick={() => void submit()} disabled={result.status === "processing"}>
              {result.status === "processing" ? <Loader2 className="h-4 w-4 animate-spin" /> : <WandSparkles className="h-4 w-4" />}
              {result.status === "processing"
                ? "Generating..."
                : mappingMode === "deterministic"
                  ? "Run Deterministic Mapping"
                  : mappingMode === "assisted"
                    ? "Plan With AI, Execute With Code"
                    : mappingMode === "sap-calm"
                      ? "Generate SAP CALM Workbook"
                      : "Generate With AI Agent"}
            </Button>
          </div>
        </CardContent>
      </Card>

      <ResultErrorBoundary>
        <ResultPanel result={result} onCopy={copyResponse} onRefine={refineWithAi} />
      </ResultErrorBoundary>
    </main>
  );
}

async function buildUploadedWorkbook(file: File | undefined): Promise<UploadedWorkbook | null> {
  if (!file) {
    return null;
  }

  return {
    id: createSessionId(),
    file,
    preview: await parseWorkbook(file)
  };
}

function getErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : "An unknown error occurred.";
}

function ensureSapCalmPrompt(currentPrompt: string) {
  const trimmedPrompt = currentPrompt.trim();
  if (!trimmedPrompt) {
    return SAP_CALM_PROMPT;
  }
  if (/SAP Cloud ALM Test Script mode/i.test(trimmedPrompt)) {
    return currentPrompt;
  }

  return `${SAP_CALM_PROMPT}\n${trimmedPrompt}`;
}

// Future saved-prompt feature: keep the shape close to the app's eventual
// persistence API so prompt-library work can drop in without changing the page.
export type SavedPromptDraft = Omit<SavedPrompt, "id" | "updatedAt">;

function createSessionId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  const bytes = new Uint8Array(16);
  if (globalThis.crypto?.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = Math.floor(Math.random() * 256);
    }
  }

  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
