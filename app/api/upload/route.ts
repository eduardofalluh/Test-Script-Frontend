import { NextResponse } from "next/server";
import { MAX_FILE_SIZE_BYTES } from "@/lib/constants";
import { sanitizeFilename } from "@/lib/utils";

export const runtime = "nodejs";

type UploadResponse = {
  ok: true;
  files: Array<{
    reference: string;
    filename: string;
    size: number;
  }>;
  note: string;
};

export async function POST(request: Request) {
  const formData = await request.formData();
  const files = formData.getAll("files").filter((entry): entry is File => entry instanceof File);

  if (files.length === 0) {
    return NextResponse.json({ ok: false, error: "No files uploaded." }, { status: 400 });
  }

  for (const file of files) {
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".xlsm")) {
      return NextResponse.json({ ok: false, error: `${file.name} is not a supported workbook.` }, { status: 400 });
    }
    if (file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json({ ok: false, error: `${file.name} exceeds the 50 MB file limit.` }, { status: 400 });
    }
  }

  return NextResponse.json<UploadResponse>({
    ok: true,
    files: files.map((file) => ({
      reference: `ephemeral://${crypto.randomUUID()}`,
      filename: sanitizeFilename(file.name),
      size: file.size
    })),
    note: "Files are validated but not persisted. The current app sends workbooks directly to /api/invoke-agent as FormData."
  });
}
