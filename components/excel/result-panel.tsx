"use client";

import * as React from "react";
import { Copy, Download, Loader2, WandSparkles } from "lucide-react";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Textarea } from "@/components/ui/textarea";
import { GeneratedWorkbookEditor } from "@/components/excel/generated-workbook-editor";
import { WorkbookPreviewTable } from "@/components/excel/workbook-preview-table";
import type { AgentFilePayload, WorkbookPreview } from "@/lib/types";
import { cn } from "@/lib/utils";

export type ResultState = {
  status: "idle" | "processing" | "success" | "error";
  responseText: string;
  statusMessage?: string;
  rawResponse?: unknown;
  file: AgentFilePayload | null;
  preview: WorkbookPreview | null;
  error?: string;
  details?: string;
};

type ResultPanelProps = {
  result: ResultState;
  onCopy: () => void;
  onRefine: (instruction: string, editedBase64: string | null) => void;
};

export function ResultPanel({ result, onCopy, onRefine }: ResultPanelProps) {
  const [refinement, setRefinement] = React.useState("");
  const [hasManualEdits, setHasManualEdits] = React.useState(false);
  const [editedBase64, setEditedBase64] = React.useState<string | null>(null);

  const [downloadHref, setDownloadHref] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    setEditedBase64(null);
    setHasManualEdits(false);
  }, [result.file?.base64]);

  // Create a proper Blob URL for downloading to prevent file corruption
  React.useEffect(() => {
    // Clean up previous URL
    if (downloadHref) {
      URL.revokeObjectURL(downloadHref);
    }

    if (result.file?.base64) {
      try {
        // Convert base64 to binary
        const binaryString = atob(result.file.base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }

        // Create Blob and URL
        const blob = new Blob([bytes], { type: result.file.mimeType });
        const url = URL.createObjectURL(blob);
        setDownloadHref(url);
      } catch (error) {
        console.error("Failed to create download URL:", error);
        setDownloadHref(undefined);
      }
    } else {
      setDownloadHref(undefined);
    }

    // Cleanup on unmount
    return () => {
      if (downloadHref) {
        URL.revokeObjectURL(downloadHref);
      }
    };
  }, [result.file?.base64, result.file?.mimeType]);

  if (result.status === "idle") {
    return null;
  }

  return (
    <Card>
      <CardHeader>
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <CardTitle>Result</CardTitle>
            <CardDescription>Review, edit, download, or revise the generated workbook.</CardDescription>
          </div>
          <StatusPill status={result.status} />
        </div>
      </CardHeader>
      <CardContent>
        {result.status === "processing" ? (
          <div className="mb-4 flex items-center gap-2 rounded-md border border-border bg-muted/40 p-3 text-sm text-muted-foreground">
            <Loader2 className="h-4 w-4 animate-spin text-primary" />
            {result.statusMessage ?? "Processing request..."}
          </div>
        ) : null}

        {result.status === "error" ? (
          <Alert variant="destructive" className="mb-4">
            <AlertTitle>Generation failed</AlertTitle>
            <AlertDescription>
              <p>{result.error}</p>
              {result.details ? (
                <details className="mt-3">
                  <summary className="cursor-pointer font-medium">Full error details</summary>
                  <pre className="mt-2 max-h-72 overflow-auto whitespace-pre-wrap rounded-md bg-background/70 p-3 text-xs">
                    {result.details}
                  </pre>
                </details>
              ) : null}
            </AlertDescription>
          </Alert>
        ) : null}

        {/* Quick Re-prompt Section - Always visible when there's a result */}
        {result.status === "success" || result.status === "error" ? (
          <div className="mb-4 rounded-lg border border-primary/30 bg-primary/5 p-4">
            <div className="mb-3 flex items-center gap-2">
              <WandSparkles className="h-5 w-5 text-primary" />
              <h3 className="font-semibold text-sm">Need changes? Re-prompt here</h3>
            </div>
            {hasManualEdits ? (
              <Alert className="mb-3">
                <AlertDescription className="text-xs">
                  You have manual edits in Easy View. The AI will use your edited workbook as the starting point.
                </AlertDescription>
              </Alert>
            ) : null}
            <div className="space-y-3">
              <Textarea
                value={refinement}
                onChange={(event) => setRefinement(event.target.value)}
                rows={4}
                placeholder="Paste CALM validation errors here or describe what needs to change. Example: 'Activity Target URL column has invalid values. URLs must follow format: https://sap.com. Affected rows: 11, 13, 14, 15.'"
                className="resize-none text-sm"
              />
              <div className="flex justify-end">
                <Button
                  type="button"
                  disabled={result.status === "processing" || refinement.trim().length === 0}
                  onClick={() => {
                    onRefine(refinement, editedBase64);
                    setRefinement("");
                  }}
                  className="gap-2"
                >
                  <WandSparkles className="h-4 w-4" />
                  Generate Fixed Version
                </Button>
              </div>
            </div>
          </div>
        ) : null}

        <Tabs defaultValue="workbook">
          <TabsList>
            <TabsTrigger value="workbook">Easy View</TabsTrigger>
            <TabsTrigger value="preview">Preview</TabsTrigger>
            <TabsTrigger value="response">Run Response</TabsTrigger>
            <TabsTrigger value="download">Download</TabsTrigger>
          </TabsList>
          <TabsContent value="workbook">
            {result.file ? (
              <GeneratedWorkbookEditor file={result.file} onDirtyChange={setHasManualEdits} onEditedFileChange={setEditedBase64} />
            ) : (
              <div className="rounded-md border border-dashed border-border p-6 text-sm text-muted-foreground">
                No generated workbook was detected. The agent response is still available in the Agent Response tab.
              </div>
            )}
          </TabsContent>
          <TabsContent value="preview">
            {result.preview ? (
              <div className="space-y-3">
                <p className="text-sm text-muted-foreground">
                  Previewing the first 20 rows of {result.preview.firstSheetName}.
                </p>
                <WorkbookPreviewTable preview={result.preview} />
              </div>
            ) : (
              <div className="rounded-md border border-dashed border-border p-6 text-sm text-muted-foreground">
                No populated workbook was detected in the agent response yet.
              </div>
            )}
          </TabsContent>
          <TabsContent value="response">
            <div className="mb-3 flex justify-end">
              <Button type="button" variant="outline" size="sm" onClick={onCopy}>
                <Copy className="h-4 w-4" />
                Copy
              </Button>
            </div>
            <pre className="scrollbar-thin max-h-[28rem] overflow-auto rounded-md border border-border bg-background/70 p-4 text-xs text-muted-foreground">
              {result.responseText || "No response text yet."}
            </pre>
          </TabsContent>
          <TabsContent value="download">
            {downloadHref ? (
              <a href={downloadHref} download={result.file?.filename ?? "populated-workbook.xlsx"}>
                <Button type="button">
                  <Download className="h-4 w-4" />
                  Download Populated File
                </Button>
              </a>
            ) : (
              <div className="rounded-md border border-dashed border-border p-6 text-sm text-muted-foreground">
                The response did not include a downloadable .xlsx file. Use the Run Response tab to inspect what came back.
              </div>
            )}
          </TabsContent>
        </Tabs>
      </CardContent>
    </Card>
  );
}

function StatusPill({ status }: { status: ResultState["status"] }) {
  const label = {
    idle: "Idle",
    processing: "Processing...",
    success: "Success",
    error: "Error"
  }[status];

  return (
    <span
      className={cn(
        "inline-flex w-fit items-center rounded-md border px-3 py-1 text-xs font-medium",
        status === "processing" && "border-primary/50 bg-primary/10 text-primary",
        status === "success" && "border-accent/50 bg-accent/10 text-accent",
        status === "error" && "border-destructive/50 bg-destructive/10 text-destructive",
        status === "idle" && "border-border bg-muted text-muted-foreground"
      )}
    >
      {label}
    </span>
  );
}
