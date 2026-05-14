"use client";

import * as React from "react";
import { Download, Plus, RotateCcw, Save } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { editableWorkbookToBase64, parseEditableWorkbookFromBase64 } from "@/lib/excel";
import type { AgentFilePayload, EditableWorkbook } from "@/lib/types";

type GeneratedWorkbookEditorProps = {
  file: AgentFilePayload;
  onDirtyChange?: (isDirty: boolean) => void;
  onEditedFileChange?: (base64: string) => void;
};

const MIN_VISIBLE_ROWS = 30;
const MIN_VISIBLE_COLUMNS = 12;

export function GeneratedWorkbookEditor({ file, onDirtyChange, onEditedFileChange }: GeneratedWorkbookEditorProps) {
  const [workbook, setWorkbook] = React.useState<EditableWorkbook>(() => normalizeWorkbook(parseEditableWorkbookFromBase64(file.base64)));
  const [editedBase64, setEditedBase64] = React.useState(file.base64);
  const [isDirty, setIsDirty] = React.useState(false);

  React.useEffect(() => {
    setWorkbook(normalizeWorkbook(parseEditableWorkbookFromBase64(file.base64)));
    setEditedBase64(file.base64);
    setIsDirty(false);
    onDirtyChange?.(false);
  }, [file.base64, onDirtyChange]);

  React.useEffect(() => {
    onEditedFileChange?.(editedBase64);
  }, [editedBase64, onEditedFileChange]);

  const activeSheet = workbook.sheets.find((sheet) => sheet.name === workbook.activeSheetName) ?? workbook.sheets[0];

  const updateCell = React.useCallback(
    (rowIndex: number, columnIndex: number, value: string) => {
      setWorkbook((current) => {
        const nextWorkbook = {
          ...current,
          sheets: current.sheets.map((sheet) => {
            if (sheet.name !== current.activeSheetName) {
              return sheet;
            }

            const nextRows = ensureGrid(sheet.rows, Math.max(sheet.rows.length, rowIndex + 1), Math.max(sheet.columnCount, columnIndex + 1));
            nextRows[rowIndex][columnIndex] = value;

            return {
              ...sheet,
              rows: nextRows,
              rowCount: nextRows.length,
              columnCount: Math.max(sheet.columnCount, columnIndex + 1)
            };
          })
        };
        setEditedBase64(editableWorkbookToBase64(nextWorkbook));
        setIsDirty(true);
        onDirtyChange?.(true);
        return nextWorkbook;
      });
    },
    [onDirtyChange]
  );

  const addRows = React.useCallback(() => {
    setWorkbook((current) => {
      const nextWorkbook = {
        ...current,
        sheets: current.sheets.map((sheet) => {
          if (sheet.name !== current.activeSheetName) {
            return sheet;
          }

          const nextRows = ensureGrid(sheet.rows, sheet.rows.length + 10, sheet.columnCount);
          return {
            ...sheet,
            rows: nextRows,
            rowCount: nextRows.length
          };
        })
      };
      setEditedBase64(editableWorkbookToBase64(nextWorkbook));
      setIsDirty(true);
      onDirtyChange?.(true);
      return nextWorkbook;
    });
  }, [onDirtyChange]);

  const reset = React.useCallback(() => {
    setWorkbook(normalizeWorkbook(parseEditableWorkbookFromBase64(file.base64)));
    setEditedBase64(file.base64);
    setIsDirty(false);
    onDirtyChange?.(false);
  }, [file.base64, onDirtyChange]);

  if (!activeSheet) {
    return (
      <div className="rounded-md border border-dashed border-border p-6 text-sm text-muted-foreground">
        The workbook did not contain any readable sheets.
      </div>
    );
  }

  const downloadHref = `data:${file.mimeType};base64,${editedBase64}`;

  return (
    <div className="space-y-4">
      <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div className="space-y-2">
          <div className="flex flex-wrap items-center gap-2">
            <Badge>{activeSheet.name}</Badge>
            <span className="text-xs text-muted-foreground">
              {activeSheet.rowCount} rows • {activeSheet.columnCount} columns
            </span>
            {isDirty ? <Badge className="border-accent/50 bg-accent/10 text-accent">Unsaved edits</Badge> : null}
          </div>
          <p className="text-sm text-muted-foreground">
            Click a cell to edit it. Download exports your manual changes as a new workbook.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <Button type="button" variant="outline" size="sm" onClick={addRows}>
            <Plus className="h-4 w-4" />
            Add Rows
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={reset} disabled={!isDirty}>
            <RotateCcw className="h-4 w-4" />
            Reset
          </Button>
          <a href={downloadHref} download={file.filename}>
            <Button type="button" size="sm">
              {isDirty ? <Save className="h-4 w-4" /> : <Download className="h-4 w-4" />}
              {isDirty ? "Download Edited File" : "Download File"}
            </Button>
          </a>
        </div>
      </div>

      <Tabs
        value={workbook.activeSheetName}
        onValueChange={(sheetName) => setWorkbook((current) => ({ ...current, activeSheetName: sheetName }))}
      >
        <TabsList className="max-w-full justify-start overflow-x-auto">
          {workbook.sheets.map((sheet) => (
            <TabsTrigger key={sheet.name} value={sheet.name}>
              {sheet.name}
            </TabsTrigger>
          ))}
        </TabsList>
      </Tabs>

      <div className="scrollbar-thin max-h-[34rem] overflow-auto rounded-md border border-border">
        <table className="w-full min-w-[900px] border-collapse text-left text-xs">
          <thead className="sticky top-0 z-10 bg-muted text-muted-foreground">
            <tr>
              <th className="sticky left-0 z-20 w-14 border-b border-r border-border bg-muted px-2 py-2 text-center font-medium">
                #
              </th>
              {Array.from({ length: activeSheet.columnCount }).map((_, index) => (
                <th key={index} className="min-w-32 border-b border-r border-border px-2 py-2 font-medium">
                  {columnLabel(index)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {activeSheet.rows.map((row, rowIndex) => (
              <tr key={rowIndex} className="odd:bg-background/30">
                <th className="sticky left-0 bg-card px-2 py-1 text-center font-medium text-muted-foreground">
                  {rowIndex + 1}
                </th>
                {Array.from({ length: activeSheet.columnCount }).map((_, columnIndex) => (
                  <td key={columnIndex} className="border-l border-t border-border p-0">
                    <Input
                      value={row[columnIndex] ?? ""}
                      onChange={(event) => updateCell(rowIndex, columnIndex, event.target.value)}
                      className="h-9 min-w-32 rounded-none border-0 bg-transparent px-2 text-xs focus-visible:ring-1"
                      aria-label={`${activeSheet.name} ${columnLabel(columnIndex)}${rowIndex + 1}`}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function normalizeWorkbook(workbook: EditableWorkbook): EditableWorkbook {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => {
      const rowCount = Math.max(sheet.rows.length, MIN_VISIBLE_ROWS);
      const columnCount = Math.max(sheet.columnCount, MIN_VISIBLE_COLUMNS);
      return {
        ...sheet,
        rows: ensureGrid(sheet.rows, rowCount, columnCount),
        rowCount,
        columnCount
      };
    })
  };
}

function ensureGrid(rows: string[][], rowCount: number, columnCount: number) {
  return Array.from({ length: rowCount }).map((_, rowIndex) => {
    const row = rows[rowIndex] ?? [];
    return Array.from({ length: columnCount }).map((__, columnIndex) => row[columnIndex] ?? "");
  });
}

function columnLabel(index: number) {
  let label = "";
  let current = index + 1;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}
