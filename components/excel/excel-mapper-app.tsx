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
import { parseWorkbook, parseWorkbookFromBase64 } from "@/lib/excel";
import type { InvokeAgentResponse, SavedPrompt, UploadedWorkbook } from "@/lib/types";
import { estimatedBase64Size, formatBytes } from "@/lib/utils";

const API_KEY_STORAGE_KEY = "excel-mapper.syntax-api-key";
const PROMPT_PLACEHOLDER =
  "Take the customer data from Source.xlsx sheet 'Customers' columns A-E and map them to the template's 'Migration Input' sheet starting row 5. Map source A to target C, source B to target D, source C to target F, and convert country names to ISO-2 codes before writing them.";

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
  const [abortController, setAbortController] = React.useState<AbortController | null>(null);
  const [result, setResult] = React.useState<ResultState>({
    status: "idle",
    responseText: "",
    file: null,
    preview: null
  });

  React.useEffect(() => {
    setSessionId(crypto.randomUUID());
    setApiKey(window.localStorage.getItem(API_KEY_STORAGE_KEY) ?? "");
  }, []);

  React.useEffect(() => {
    if (apiKey) {
      window.localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    }
  }, [apiKey]);

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

  const regenerateSession = React.useCallback(() => {
    setSessionId(crypto.randomUUID());
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

  const submit = React.useCallback(async (overridePrompt?: string) => {
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
    if (!apiKey.trim()) {
      toast({ title: "API key required", description: "Paste your Syntax GenAI Studio API key." });
      return;
    }

    const controller = new AbortController();
    setAbortController(controller);
    setResult({
      status: "processing",
      responseText: "Request sent to Test Script IQ. Waiting for the final response...",
      file: null,
      preview: null
    });
    toast({ title: "Generation started", description: "Sending files and prompt to Test Script IQ." });

    const formData = new FormData();
    formData.append("template", templateFiles[0].file);
    sourceFiles.forEach((source) => formData.append("sources", source.file));
    formData.append("prompt", trimmedPrompt);
    formData.append("apiKey", apiKey.trim());
    formData.append("sessionId", sessionId || crypto.randomUUID());

    try {
      const response = await fetch("/api/invoke-agent", {
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
        description: payload.file ? "Populated workbook detected." : "Agent responded without a workbook."
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
  }, [apiKey, prompt, result.status, sessionId, sourceFiles, templateFiles, toast]);

  const refineWithAi = React.useCallback(
    (instruction: string) => {
      const trimmedInstruction = instruction.trim();
      if (!trimmedInstruction) {
        return;
      }

      const revisedPrompt = [
        prompt.trim(),
        "",
        "Revision request for the generated workbook:",
        trimmedInstruction,
        "",
        "Use the original template and source files again. Preserve any correct mappings from the previous run and only change what is needed.",
        result.responseText ? `Previous agent response:\n${result.responseText}` : ""
      ]
        .filter(Boolean)
        .join("\n");

      setPrompt(revisedPrompt);
      void submit(revisedPrompt);
    },
    [prompt, result.responseText, submit]
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
          Upload a target template, attach source workbooks, describe the mapping, and let Test Script IQ prepare the populated file.
        </p>
      </header>

      <Alert className="border-primary/40 bg-primary/10">
        <ShieldAlert className="h-4 w-4" />
        <AlertTitle>Processing notice</AlertTitle>
        <AlertDescription>
          Files are sent to Syntax GenAI Studio for processing. Do not upload files containing client-confidential data unless you have authorization.
        </AlertDescription>
      </Alert>

      <section className="grid gap-4 lg:grid-cols-2">
        <WorkbookDropzone
          title="Template Upload"
          description="Target workbook to be populated."
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

          <div className="grid gap-4 lg:grid-cols-[1fr_22rem]">
            <div className="space-y-2">
              <label htmlFor="api-key" className="text-sm font-medium">
                API Key
              </label>
              <div className="flex gap-2">
                <div className="relative flex-1">
                  <KeyRound className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-muted-foreground" />
                  <Input
                    id="api-key"
                    type="password"
                    value={apiKey}
                    onChange={(event) => setApiKey(event.target.value)}
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
              {result.status === "processing" ? "Generating..." : "Generate Populated File"}
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
    id: crypto.randomUUID(),
    file,
    preview: await parseWorkbook(file)
  };
}

function getErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : "An unknown error occurred.";
}

// Future saved-prompt feature: keep the shape close to the app's eventual
// persistence API so prompt-library work can drop in without changing the page.
export type SavedPromptDraft = Omit<SavedPrompt, "id" | "updatedAt">;
