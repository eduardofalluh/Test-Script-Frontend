# Excel Mapper

Excel Mapper is a single-user Next.js app for SAP consultants who need to map source Excel workbooks into a target Excel template with deterministic, coded mapping logic. Syntax GenAI Studio's **Test Script IQ** agent can optionally interpret natural-language prompts into structured JSON mapping plans, but the app still performs the Excel writing with code.

The app lets you upload a target `.xlsx` or `.xlsm` template, upload one or more source workbooks, write detailed mapping instructions, run cell-by-cell mapping through a server-side Next.js API route, review the returned workbook in an editable Easy View, revise the output, and download the populated file. It also includes a **SAP CALM Test Script** mode that automatically loads the bundled SAP Cloud ALM Test Cases template and maps source data into that format.

## Tech Stack

- Next.js 14 App Router
- TypeScript
- Tailwind CSS with shadcn/ui-style components
- `react-dropzone` for workbook uploads
- `@e965/xlsx` for client-side workbook previews
- Next.js API routes for deterministic mapping and the optional Syntax GenAI Studio proxy

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

Windows testers must use the Windows x64 Node.js LTS installer. The 32-bit Windows Node.js build reports `process.arch` as `ia32` and cannot load the Next.js SWC compiler.

The first run installs dependencies, then opens http://localhost:3010.

## API Key

The API key is only required when you choose **AI-Assisted Deterministic** or **External AI Agent** mode. Paste your Syntax GenAI Studio API key into the password field in the app. It is stored only in your browser's `localStorage` and sent to the local API proxy for AI requests. Use **Forget** to remove it from `localStorage`.

The hardcoded agent endpoint is:

```text
https://studio-api.ai.syntax-rnd.com/api/v1/agents/6d310742-9d0a-4069-8689-6c8feb61b935/invoke
```

## How It Works

1. The browser validates and previews uploaded Excel workbooks.
2. In default **Deterministic Mapper** mode, the browser submits the template, sources, and prompt to `/api/map-workbook`.
3. The deterministic route parses explicit instructions such as source sheet, target sheet, starting row, and A->C style column mappings, then copies cells with SheetJS on the server.
4. If matching headers are available, the deterministic mapper can infer common mappings such as `Name` -> `Customer Name` and `Contact Email` -> `Email`.
5. If requested, simple coded transformations such as country name to ISO-2 conversion are applied.
6. **AI-Assisted Deterministic** mode sends workbook summaries and the user's prompt to `/api/plan-and-map-workbook`. Test Script IQ returns JSON mapping specs only; the app validates and executes those specs with deterministic SheetJS logic.
7. **SAP CALM Test Script** mode loads `public/templates/sap-cloud-alm-test-cases-template.xlsx`, pre-fills SAP Cloud ALM mapping guidance, asks the user for clarification in the prompt, clears the template's sample rows, and writes into the `Test Cases` sheet starting at row 2.
8. **External AI Agent** mode still uses `/api/invoke-agent`, converts workbooks to base64 data URLs, builds the agent prompt, and lets Test Script IQ return a workbook.
9. If a workbook is found, the browser opens an Easy View with sheet switching, editable cells, and an edited-file download. Otherwise, the full run response remains visible for diagnosis.
10. Use the Revise tab to describe follow-up changes. The app re-runs the currently selected mapping engine.

## Test Like A User

Run the browser E2E test:

```bash
npm run test:e2e
```

The test creates complex mock Excel files, starts a local mock Syntax agent, uploads the files through the UI, submits the mapping prompt, verifies deterministic mapping without invoking the mock agent, checks a larger header-inference workbook, verifies SAP CALM template mode, verifies AI-assisted planning produces only a JSON mapping plan before deterministic execution, edits a result cell, checks API-key/session persistence after refresh, and confirms the optional AI fallback can still call the mock agent. The test uses `SYNTAX_AGENT_ENDPOINT` to avoid calling the real external agent.

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
- Uploaded files are not persisted. Deterministic mode sends files only to the local Next.js route. AI-Assisted Deterministic mode sends workbook summaries and prompt text to Test Script IQ for planning. External AI Agent mode sends full files to `/api/invoke-agent`, which proxies to Syntax GenAI Studio.
- Deterministic mapping currently handles explicit sheet/start-row/column mappings, common header-based mappings, source row notes, validation notes, constant values, clearing sample target rows, and country-name-to-ISO2 conversion.
- Ambiguous mapping instructions should either be rewritten with explicit rules or run through AI-Assisted Deterministic mode so the LLM can translate the prompt into executable JSON specs.
- Manual Easy View edits are exported in the browser. Revision requests can use the edited workbook as the next template.
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
