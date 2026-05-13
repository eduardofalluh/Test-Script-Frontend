import { createServer, type Server } from "node:http";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { expect, test } from "@playwright/test";
import * as XLSX from "@e965/xlsx";

const MOCK_AGENT_PORT = 4011;

let mockServer: Server;
let invocationCount = 0;
let latestAgentRequest: unknown = null;

test.beforeAll(async () => {
  mockServer = createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/invoke") {
      response.writeHead(404);
      response.end("Not found");
      return;
    }

    const body = await readRequestBody(request);
    latestAgentRequest = JSON.parse(body);
    invocationCount += 1;

    response.writeHead(200, { "Content-Type": "application/json" });
    response.end(
      JSON.stringify({
        message: `Mock Test Script IQ response ${invocationCount}: populated workbook created with validation summary.`,
        output_file: createGeneratedWorkbookBase64(invocationCount),
        filename: "populated-customer-template.xlsx"
      })
    );
  });

  await new Promise<void>((resolve) => {
    mockServer.listen(MOCK_AGENT_PORT, "127.0.0.1", resolve);
  });
});

test.afterAll(async () => {
  await new Promise<void>((resolve, reject) => {
    mockServer.close((error) => (error ? reject(error) : resolve()));
  });
});

test("uploads complex Excel files, invokes mock agent, edits result, and persists local credentials", async ({
  page
}, testInfo) => {
  const templatePath = testInfo.outputPath("Target_Template.xlsx");
  const sourcePath = testInfo.outputPath("Complex_Source_Data.xlsx");
  mkdirSync(dirname(templatePath), { recursive: true });
  writeFileSync(templatePath, createTemplateWorkbookBuffer());
  writeFileSync(sourcePath, createComplexSourceWorkbookBuffer());

  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Excel Mapper" })).toBeVisible();

  const sessionField = page.getByRole("textbox", { name: "Session ID" });
  await expect(sessionField).not.toHaveValue("");
  const firstSessionId = await sessionField.inputValue();

  const apiKeyField = page.getByLabel("API Key");
  await apiKeyField.fill("mock-local-api-key");
  await page.reload();
  await expect(page.getByLabel("API Key")).toHaveValue("mock-local-api-key");
  await expect(page.getByRole("textbox", { name: "Session ID" })).toHaveValue(firstSessionId);

  const fileInputs = page.locator("input[type=file]");
  await fileInputs.nth(0).setInputFiles(templatePath);
  await fileInputs.nth(1).setInputFiles(sourcePath);

  const app = page.getByRole("main");
  await expect(app.getByText("Target_Template.xlsx")).toBeVisible();
  await expect(app.getByText("Complex_Source_Data.xlsx")).toBeVisible();
  await expect(app.getByText("Migration Input", { exact: true })).toBeVisible();
  await expect(app.getByText("Customers", { exact: true })).toBeVisible();

  await page.getByLabel("Mapping Prompt").fill(
    [
      "Map Complex_Source_Data.xlsx sheet Customers rows into Target_Template.xlsx sheet Migration Input starting row 5.",
      "Use Customer ID, Name, Country, Currency, Risk Score, and Contact Email.",
      "Also summarize Orders and Region Mapping sheets in the validation output."
    ].join(" ")
  );
  await page.getByRole("button", { name: "Generate Populated File" }).click();

  await expect(page.getByText("Success")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Easy View" })).toHaveAttribute("data-state", "active");
  await expect(page.getByRole("tab", { name: "Migration Input" })).toBeVisible();
  await expect(page.locator('input[value="CUST-001"]')).toBeVisible();

  const editableCustomerName = page.getByLabel("Migration Input C5");
  await editableCustomerName.fill("Acme Canada Edited");
  await expect(page.getByText("Unsaved edits")).toBeVisible();
  await expect(page.getByRole("button", { name: "Download Edited File" })).toBeVisible();

  await page.getByRole("tab", { name: "AI Modify" }).click();
  await page.getByLabel("Tell the agent what to change").fill("Change country values to ISO-2 codes and keep edited customer names.");
  await page.getByRole("button", { name: "Ask AI to Revise" }).click();
  await page.getByRole("tab", { name: "Agent Response" }).click();
  await expect(page.getByText("Mock Test Script IQ response 2")).toBeVisible();

  expect(invocationCount).toBe(2);
  expect(JSON.stringify(latestAgentRequest)).toContain("Revision request for the generated workbook");
  expect(JSON.stringify(latestAgentRequest)).toContain("data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,");
});

function createTemplateWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  const migrationInput = [
    ["SAP Migration Template", "", "", "", "", "", "", ""],
    ["Object", "Customer Master", "", "", "", "", "", ""],
    ["Instructions", "Populate from row 5", "", "", "", "", "", ""],
    ["Legacy ID", "Source Row", "Customer Name", "Country", "Currency", "Risk Score", "Email", "Validation Notes"],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""]
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(migrationInput), "Migration Input");
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Field", "Required", "Rule"],
      ["Customer Name", "Yes", "Trim whitespace"],
      ["Country", "Yes", "Use ISO-2 when requested"],
      ["Risk Score", "No", "0-100"]
    ]),
    "Field Rules"
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

function createComplexSourceWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Customer ID", "Name", "Country", "Currency", "Risk Score", "Contact Email", "Created On"],
      ["CUST-001", "Acme Canada", "Canada", "CAD", 12, "ops@acme.example", "2026-01-15"],
      ["CUST-002", "Nordic Parts", "Sweden", "SEK", 35, "finance@nordic.example", "2026-02-01"],
      ["CUST-003", "Sao Paulo Metals", "Brazil", "BRL", 54, "tax@spmetals.example", "2026-02-20"]
    ]),
    "Customers"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Order ID", "Customer ID", "Amount", "Currency", "Open"],
      ["SO-1001", "CUST-001", 12500.5, "CAD", true],
      ["SO-1002", "CUST-002", 8450, "SEK", false],
      ["SO-1003", "CUST-003", 18500.75, "BRL", true]
    ]),
    "Orders"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Country", "Region", "SAP Sales Org"],
      ["Canada", "NA", "CA01"],
      ["Sweden", "EMEA", "SE01"],
      ["Brazil", "LATAM", "BR01"]
    ]),
    "Region Mapping"
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

function createGeneratedWorkbookBase64(runNumber: number) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["SAP Migration Template", "", "", "", "", "", "", ""],
      ["Object", "Customer Master", "", "", "", "", "", ""],
      ["Instructions", "Populated by mock Test Script IQ", "", "", "", "", "", ""],
      ["Legacy ID", "Source Row", "Customer Name", "Country", "Currency", "Risk Score", "Email", "Validation Notes"],
      ["CUST-001", 2, runNumber === 1 ? "Acme Canada" : "Acme Canada Revised", "Canada", "CAD", 12, "ops@acme.example", "Valid"],
      ["CUST-002", 3, "Nordic Parts", "Sweden", "SEK", 35, "finance@nordic.example", "Valid"],
      ["CUST-003", 4, "Sao Paulo Metals", "Brazil", "BRL", 54, "tax@spmetals.example", "Valid"]
    ]),
    "Migration Input"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Check", "Status"],
      ["Template sheet found", "Pass"],
      ["Source sheets processed", "Pass"],
      ["Run number", runNumber]
    ]),
    "Validation Summary"
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "base64" }) as string;
}

function readRequestBody(request: NodeJS.ReadableStream) {
  return new Promise<string>((resolve, reject) => {
    const chunks: Buffer[] = [];
    request.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    request.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    request.on("error", reject);
  });
}
