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

test.beforeEach(() => {
  invocationCount = 0;
  latestAgentRequest = null;
});

test("uploads complex Excel files, maps deterministically, edits result, and keeps AI fallback available", async ({
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
      "Map Complex_Source_Data.xlsx source sheet 'Customers' rows into Target_Template.xlsx target sheet 'Migration Input' starting row 5.",
      "Map A->A, B->C, C->D, D->E, E->F, F->G.",
      "Convert country names to ISO-2 codes."
    ].join(" ")
  );
  await page.getByRole("button", { name: "Run Deterministic Mapping" }).click();

  await expect(page.getByText("Success")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Easy View" })).toHaveAttribute("data-state", "active");
  await expect(page.getByRole("tab", { name: "Migration Input" })).toBeVisible();
  await expect(page.locator('input[value="CUST-001"]')).toBeVisible();
  await expect(page.locator('input[value="CA"]')).toBeVisible();
  expect(invocationCount).toBe(0);

  const editableCustomerName = page.getByLabel("Migration Input C5");
  await editableCustomerName.fill("Acme Canada Edited");
  await expect(page.getByText("Unsaved edits")).toBeVisible();
  await expect(page.getByRole("button", { name: "Download Edited File" })).toBeVisible();

  await page.getByRole("button", { name: "External AI Agent" }).click();
  await page.getByRole("button", { name: "Generate With AI Agent" }).click();
  await page.getByRole("tab", { name: "Run Response" }).click();
  await expect(page.getByText("Mock Test Script IQ response 1")).toBeVisible();

  expect(invocationCount).toBe(1);
  expect(JSON.stringify(latestAgentRequest)).toContain("data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,");
});

test("maps a larger multi-sheet workbook with header inference and no API key", async ({ page }, testInfo) => {
  const templatePath = testInfo.outputPath("Large_Target_Template.xlsx");
  const sourcePath = testInfo.outputPath("Large_Source_Data.xlsx");
  mkdirSync(dirname(templatePath), { recursive: true });
  writeFileSync(templatePath, createTemplateWorkbookBuffer());
  writeFileSync(sourcePath, createLargeSourceWorkbookBuffer(80));

  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Excel Mapper" })).toBeVisible();
  await expect(page.getByLabel("API Key")).toHaveValue("");

  const fileInputs = page.locator("input[type=file]");
  await fileInputs.nth(0).setInputFiles(templatePath);
  await fileInputs.nth(1).setInputFiles(sourcePath);

  await page.getByLabel("Mapping Prompt").fill(
    [
      "Map source sheet 'Customers' into target sheet 'Migration Input' starting row 5.",
      "Use matching source and target headers, include source row and validation notes, and convert country names to ISO-2 codes."
    ].join(" ")
  );
  await page.getByRole("button", { name: "Run Deterministic Mapping" }).click();

  await expect(page.getByText("Success")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Migration Input" })).toBeVisible();
  await expect(page.locator('input[value="CUST-001"]')).toBeVisible();
  await expect(page.locator('input[value="CUST-042"]')).toBeVisible();
  await expect(page.locator('input[value="Customer 42"]')).toBeVisible();
  await expect(page.getByLabel("Migration Input D46")).toHaveValue("SE");
  await expect(page.locator('input[value="Mapped deterministically"]').first()).toBeVisible();
  await expect(page.getByRole("button", { name: "Download File" })).toBeVisible();

  await page.getByRole("tab", { name: "Run Response" }).click();
  await expect(page.getByText("Rows mapped: 80")).toBeVisible();
  await expect(page.getByText("Deterministic mapper completed without using the external AI agent.")).toBeVisible();
  expect(invocationCount).toBe(0);
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

function createLargeSourceWorkbookBuffer(rowCount: number) {
  const workbook = XLSX.utils.book_new();
  const countries = [
    ["Canada", "CAD"],
    ["Sweden", "SEK"],
    ["Brazil", "BRL"],
    ["United States", "USD"],
    ["Germany", "EUR"]
  ];
  const customers = [["Customer ID", "Name", "Country", "Currency", "Risk Score", "Contact Email", "Created On"]];

  for (let index = 1; index <= rowCount; index += 1) {
    const [country, currency] = countries[(index - 1) % countries.length];
    customers.push([
      `CUST-${String(index).padStart(3, "0")}`,
      `Customer ${index}`,
      country,
      currency,
      String((index * 7) % 100),
      `customer${index}@example.com`,
      `2026-03-${String((index % 28) + 1).padStart(2, "0")}`
    ]);
  }

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(customers), "Customers");
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Sales Org", "Country", "Company Code"],
      ["CA01", "Canada", "1000"],
      ["SE01", "Sweden", "2000"],
      ["BR01", "Brazil", "3000"],
      ["US01", "United States", "4000"],
      ["DE01", "Germany", "5000"]
    ]),
    "Reference Data"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Ignored ID", "Description"],
      ["X-001", "This sheet should not be mapped"],
      ["X-002", "Prompt selects Customers explicitly"]
    ]),
    "Ignore Me"
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
