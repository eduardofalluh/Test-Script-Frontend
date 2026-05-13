import type { WorkbookPreview } from "@/lib/types";

type WorkbookPreviewTableProps = {
  preview: WorkbookPreview;
};

export function WorkbookPreviewTable({ preview }: WorkbookPreviewTableProps) {
  if (preview.firstSheetRows.length === 0) {
    return (
      <div className="rounded-md border border-dashed border-border p-4 text-sm text-muted-foreground">
        The first sheet is empty.
      </div>
    );
  }

  const columnCount = Math.max(...preview.firstSheetRows.map((row) => row.length), 1);

  return (
    <div className="scrollbar-thin max-h-80 overflow-auto rounded-md border border-border">
      <table className="w-full min-w-[640px] border-collapse text-left text-xs">
        <thead className="sticky top-0 bg-muted text-muted-foreground">
          <tr>
            {Array.from({ length: columnCount }).map((_, index) => (
              <th key={index} className="border-b border-r border-border px-3 py-2 font-medium">
                {columnLabel(index)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {preview.firstSheetRows.map((row, rowIndex) => (
            <tr key={rowIndex} className="odd:bg-background/30">
              {Array.from({ length: columnCount }).map((_, columnIndex) => (
                <td key={columnIndex} className="max-w-[220px] truncate border-r border-t border-border px-3 py-2">
                  {formatCell(row[columnIndex])}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function formatCell(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value);
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
