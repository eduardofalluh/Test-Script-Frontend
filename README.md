# Excel Mapper

Excel Mapper is a single-user Next.js app for SAP consultants who need to map source Excel workbooks into a target Excel template with help from Syntax GenAI Studio's **Test Script IQ** agent.

The app lets you upload a target `.xlsx` or `.xlsm` template, upload one or more source workbooks, write detailed natural-language mapping instructions, invoke the external agent through a server-side Next.js API route, review the returned workbook in an editable Easy View, ask the AI for revisions, and download the populated file.

## Tech Stack

- Next.js 14 App Router
- TypeScript
- Tailwind CSS with shadcn/ui-style components
- `react-dropzone` for workbook uploads
- `@e965/xlsx` for client-side workbook previews
- Next.js API routes for the Syntax GenAI Studio proxy

## Install

```bash
npm install
```

## Run Locally

```bash
npm run dev
```

Open http://localhost:3000.

`npm run dev` clears stale `.next` output before starting. This keeps local styles from disappearing if a production build was run before restarting the development server.

For a predictable shareable local trial, use:

```bash
npm run dev:local
```

Open http://localhost:3010.

## Share With Testers

Create a local trial package:

```bash
npm run package:release
```

This creates `release/excel-mapper-local.zip` and a matching unpacked folder. Send the `.zip` file to testers. They should unzip it, install Node.js LTS from https://nodejs.org, then double-click one of these starter files:

- `START_EXCEL_MAPPER_MAC.command`
- `START_EXCEL_MAPPER_WINDOWS.bat`

The first run installs dependencies, then opens http://localhost:3010.

## API Key

Paste your Syntax GenAI Studio API key into the password field in the app. It is stored only in your browser's `localStorage` and sent to the local `/api/invoke-agent` proxy for each request. Use **Forget** to remove it from `localStorage`.

The hardcoded agent endpoint is:

```text
https://studio-api.ai.syntax-rnd.com/api/v1/agents/6d310742-9d0a-4069-8689-6c8feb61b935/invoke
```

## How It Works

1. The browser validates and previews uploaded Excel workbooks.
2. On generation, the browser submits the template, sources, prompt, API key, and session ID to `/api/invoke-agent`.
3. The API route validates file sizes and extensions, converts workbooks to base64 data URLs, builds the agent prompt, and invokes Test Script IQ.
4. The response parser looks for a returned `.xlsx` payload in common keys such as `output_file`, `file`, `result`, or any base64 value starting with the zipped workbook prefix `UEsDB`.
5. If a workbook is found, the browser opens an Easy View with sheet switching, editable cells, and an edited-file download. Otherwise, the full agent text response remains visible for diagnosis.
6. Use the AI Modify tab to describe follow-up changes. The app re-runs the agent with the original files, the original prompt, the revision request, and the prior agent response as context.

## Test Like A User

Run the browser E2E test:

```bash
npm run test:e2e
```

The test creates complex mock Excel files, starts a local mock Syntax agent, uploads the files through the UI, submits the mapping prompt, verifies the generated workbook Easy View, edits a result cell, checks API-key/session persistence after refresh, and exercises the AI Modify flow. The test uses `SYNTAX_AGENT_ENDPOINT` to avoid calling the real external agent.

## Deploy

### Vercel

Push this repository to GitHub and import it into Vercel. No environment variables are required for the current version because the agent endpoint is hardcoded and the API key is supplied by the user at runtime.

### Self-hosted

```bash
npm run build
npm run start
```

Run behind HTTPS in production so browser security and API-key handling stay sane.

## Configuration

The current constants live in `lib/constants.ts`:

- `AGENT_ENDPOINT`
- `AGENT_NAME`
- `REQUEST_TIMEOUT_MS`
- `MAX_FILE_SIZE_BYTES`
- `WARN_FILE_SIZE_BYTES`

`.env.example` includes a placeholder for making the agent endpoint configurable later.

## Known Limitations

- Only `.xlsx` and `.xlsm` files are accepted.
- Each workbook is capped at 50 MB.
- The encoded request payload is capped at roughly 50 MB.
- Uploaded files are not persisted. The app sends files directly to `/api/invoke-agent` for each generation request.
- Manual Easy View edits are exported in the browser. AI revision currently re-runs against the original uploaded template and source files rather than the manually edited workbook.
- The Syntax API request currently passes Excel files as base64 `data:` URLs in `image_url.url`, matching the available agent input shape. If the agent backend rejects non-image data URLs, replace this with short-lived presigned URLs. A `TemporaryFileHost` interface stub is already in `lib/file-hosting.ts`.
- The exact agent response shape is unknown. The parser is intentionally flexible, but it may need adjustment once real responses are observed.
- `npm audit` currently flags Next.js 14 even at the newest available 14.x release and recommends Next 16. This project stays on Next 14 to match the requested stack.

## Scripts

```bash
npm run clean
npm run dev
npm run dev:local
npm run package:release
npm run test:e2e
npm run lint
npm run typecheck
npm run build
```
