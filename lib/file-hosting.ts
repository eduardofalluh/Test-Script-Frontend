export type HostedFileReference = {
  filename: string;
  url: string;
};

export interface TemporaryFileHost {
  put(file: File): Promise<HostedFileReference>;
}

// TODO: If Syntax GenAI Studio rejects data URLs for Excel content, replace the
// inline data URL flow with a TemporaryFileHost implementation backed by
// presigned S3, Azure Blob SAS URLs, or another short-lived HTTPS file host.
export const temporaryFileHost: TemporaryFileHost | null = null;
