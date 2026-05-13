"use client";

import { FileSpreadsheet, Trash2, UploadCloud } from "lucide-react";
import { useDropzone } from "react-dropzone";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { WorkbookPreviewTable } from "@/components/excel/workbook-preview-table";
import { MAX_FILE_SIZE_BYTES, WARN_FILE_SIZE_BYTES } from "@/lib/constants";
import type { UploadedWorkbook } from "@/lib/types";
import { cn, formatBytes, isAcceptedExcelFile } from "@/lib/utils";

type WorkbookDropzoneProps = {
  title: string;
  description: string;
  files: UploadedWorkbook[];
  multiple?: boolean;
  onDropAccepted: (files: File[]) => void;
  onRemove: (id: string) => void;
  showPreview?: boolean;
};

export function WorkbookDropzone({
  title,
  description,
  files,
  multiple = false,
  onDropAccepted,
  onRemove,
  showPreview = false
}: WorkbookDropzoneProps) {
  const { getRootProps, getInputProps, isDragActive, fileRejections } = useDropzone({
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"]
    },
    maxSize: MAX_FILE_SIZE_BYTES,
    multiple,
    onDropAccepted
  });

  return (
    <Card className="h-full">
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div
          {...getRootProps()}
          className={cn(
            "flex min-h-40 cursor-pointer flex-col items-center justify-center rounded-lg border border-dashed border-border bg-background/50 p-5 text-center transition-colors hover:bg-muted/60",
            isDragActive && "border-primary bg-primary/10"
          )}
        >
          <input {...getInputProps()} />
          <UploadCloud className="h-8 w-8 text-primary" />
          <p className="mt-3 text-sm font-medium">{isDragActive ? "Drop workbook here" : "Drag workbook here"}</p>
          <p className="mt-1 text-xs text-muted-foreground">.xlsx or .xlsm, up to 50 MB each</p>
        </div>

        {fileRejections.length > 0 ? (
          <Alert variant="destructive">
            <AlertDescription>
              {fileRejections
                .map((rejection) => `${rejection.file.name}: ${rejection.errors.map((error) => error.message).join(", ")}`)
                .join("; ")}
            </AlertDescription>
          </Alert>
        ) : null}

        {files.length === 0 ? (
          <div className="rounded-md border border-dashed border-border p-4 text-sm text-muted-foreground">
            No workbook uploaded yet.
          </div>
        ) : (
          <div className="space-y-3">
            {files.map((uploaded) => (
              <div key={uploaded.id} className="rounded-md border border-border bg-background/40 p-4">
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="h-4 w-4 shrink-0 text-accent" />
                      <p className="truncate text-sm font-medium">{uploaded.file.name}</p>
                    </div>
                    <p className="mt-1 text-xs text-muted-foreground">
                      {formatBytes(uploaded.file.size)} • {uploaded.preview.sheetNames.length} sheet
                      {uploaded.preview.sheetNames.length === 1 ? "" : "s"} • {uploaded.preview.firstSheetRowCount} rows in{" "}
                      {uploaded.preview.firstSheetName}
                    </p>
                  </div>
                  <Button
                    type="button"
                    variant="ghost"
                    size="icon"
                    aria-label={`Remove ${uploaded.file.name}`}
                    title={`Remove ${uploaded.file.name}`}
                    onClick={() => onRemove(uploaded.id)}
                  >
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </div>

                <div className="mt-3 flex flex-wrap gap-2">
                  {uploaded.preview.sheetNames.map((sheet) => (
                    <Badge key={sheet}>{sheet}</Badge>
                  ))}
                </div>

                {uploaded.warning ? (
                  <Alert className="mt-3">
                    <AlertDescription>{uploaded.warning}</AlertDescription>
                  </Alert>
                ) : null}

                {!isAcceptedExcelFile(uploaded.file) ? (
                  <Alert variant="destructive" className="mt-3">
                    <AlertDescription>Only .xlsx and .xlsm workbooks are supported.</AlertDescription>
                  </Alert>
                ) : null}

                {uploaded.file.size > WARN_FILE_SIZE_BYTES ? (
                  <Alert className="mt-3">
                    <AlertDescription>
                      This file is larger than 10 MB and may make the agent request slower.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {showPreview ? (
                  <div className="mt-4">
                    <p className="mb-2 text-xs font-medium uppercase tracking-[0.18em] text-muted-foreground">
                      First 20 rows
                    </p>
                    <WorkbookPreviewTable preview={uploaded.preview} />
                  </div>
                ) : null}
              </div>
            ))}
          </div>
        )}
      </CardContent>
    </Card>
  );
}
