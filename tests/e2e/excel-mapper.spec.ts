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
    const requestText = JSON.stringify(latestAgentRequest);

    if (requestText.includes("You are a mapping planner")) {
      const mappingPlan = requestText.includes("Customer Master - Export")
        ? {
            sourceSheetName: "Customer Master - Export",
            targetSheetName: "Migration Input",
            targetStartRow: 8,
            mappings: [
              { generated: "validation-note", targetHeader: "Validation Notes" },
              { sourceHeader: "Primary Email", targetHeader: "Email" },
              { sourceHeader: "Risk Rating", targetHeader: "Risk Score" },
              { sourceHeader: "Local Currency", targetHeader: "Currency" },
              { sourceHeader: "Country Name", targetHeader: "Country" },
              { sourceHeader: "Legal Name", targetHeader: "Customer Name" },
              { generated: "source-row", targetHeader: "Source Row" },
              { sourceHeader: "Customer Number", targetHeader: "Legacy ID" }
            ],
            transformations: {
              convertCountryToIso2: true
            }
          }
        : {
            sourceSheetName: "Customers",
            targetSheetName: "Migration Input",
            targetStartRow: 5,
            mappings: [
              { sourceHeader: "Customer ID", targetHeader: "Legacy ID" },
              { generated: "source-row", targetHeader: "Source Row" },
              { sourceHeader: "Name", targetHeader: "Customer Name" },
              { sourceHeader: "Country", targetHeader: "Country" },
              { sourceHeader: "Currency", targetHeader: "Currency" },
              { sourceHeader: "Risk Score", targetHeader: "Risk Score" },
              { sourceHeader: "Contact Email", targetHeader: "Email" },
              { generated: "validation-note", targetHeader: "Validation Notes" }
            ],
            transformations: {
              convertCountryToIso2: true
            }
          };

      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(
        JSON.stringify({
          message: JSON.stringify(mappingPlan)
        })
      );
      return;
    }

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

test("uses LLM only to plan mappings, then executes Excel writes deterministically", async ({ page }, testInfo) => {
  const templatePath = testInfo.outputPath("Planner_Target_Template.xlsx");
  const sourcePath = testInfo.outputPath("Planner_Source_Data.xlsx");
  mkdirSync(dirname(templatePath), { recursive: true });
  writeFileSync(templatePath, createTemplateWorkbookBuffer());
  writeFileSync(sourcePath, createComplexSourceWorkbookBuffer());

  await page.goto("/");
  await page.getByLabel("API Key").fill("mock-local-api-key");

  const fileInputs = page.locator("input[type=file]");
  await fileInputs.nth(0).setInputFiles(templatePath);
  await fileInputs.nth(1).setInputFiles(sourcePath);

  await page.getByLabel("Mapping Prompt").fill(
    "Please map the Customers sheet into the Migration Input sheet starting at row 5. Use the customer id, source row, name, country as ISO code, currency, risk score, email, and add validation notes."
  );
  await page.getByRole("button", { name: "AI-Assisted Deterministic" }).click();
  await page.getByRole("button", { name: "Plan With AI, Execute With Code" }).click();

  await expect(page.getByText("Success")).toBeVisible();
  await expect(page.getByLabel("Migration Input A5")).toHaveValue("CUST-001");
  await expect(page.getByLabel("Migration Input D5")).toHaveValue("CA");
  await expect(page.locator('input[value="Mapped deterministically"]').first()).toBeVisible();
  await page.getByRole("tab", { name: "Run Response" }).click();
  await expect(page.getByText("AI-assisted mapping plan created by Test Script IQ. Workbook populated by deterministic mapper.")).toBeVisible();
  await expect(page.getByText("Mapping plan:")).toBeVisible();

  expect(invocationCount).toBe(1);
  expect(JSON.stringify(latestAgentRequest)).toContain("You are a mapping planner");
  expect(JSON.stringify(latestAgentRequest)).not.toContain("data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,");
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

test("handles an intricate natural-language AI-assisted prompt and verifies downloaded workbook cells", async ({
  page
}, testInfo) => {
  const templatePath = testInfo.outputPath("Intricate_Target_Template.xlsx");
  const sourcePath = testInfo.outputPath("Intricate_Source_Data.xlsx");
  mkdirSync(dirname(templatePath), { recursive: true });
  writeFileSync(templatePath, createIntricateTemplateWorkbookBuffer());
  writeFileSync(sourcePath, createIntricateSourceWorkbookBuffer());

  await page.goto("/");
  await page.getByLabel("API Key").fill("mock-local-api-key");

  const fileInputs = page.locator("input[type=file]");
  await fileInputs.nth(0).setInputFiles(templatePath);
  await fileInputs.nth(1).setInputFiles(sourcePath);

  await page.getByLabel("Mapping Prompt").fill(
    [
      "Build the SAP customer onboarding load from the workbook named Intricate_Source_Data.xlsx.",
      "The source data is on sheet 'Customer Master - Export'. Ignore the Orders, Notes, and Reference sheets.",
      "Populate the target template sheet 'Migration Input' beginning on row 8, because rows 1-7 are instructions and headers.",
      "The target columns are intentionally not in the same order as the source.",
      "Use Customer Number as Legacy ID, Legal Name as Customer Name, Country Name as ISO-2 Country, Local Currency as Currency, Risk Rating as Risk Score, and Primary Email as Email.",
      "Also write the original source Excel row number and add a validation note for each populated row."
    ].join(" ")
  );
  await page.getByRole("button", { name: "AI-Assisted Deterministic" }).click();
  await page.getByRole("button", { name: "Plan With AI, Execute With Code" }).click();

  await expect(page.getByText("Success")).toBeVisible();
  await expect(page.getByLabel("Migration Input H8")).toHaveValue("IC-1001");
  await expect(page.getByLabel("Migration Input F8")).toHaveValue("Northwind Canada Ltd");
  await expect(page.getByLabel("Migration Input E8")).toHaveValue("CA");
  await expect(page.getByLabel("Migration Input B9")).toHaveValue("ap@contoso-se.example");
  await expect(page.getByLabel("Migration Input A10")).toHaveValue("Mapped deterministically");

  const workbook = await downloadWorkbook(page);
  const sheet = workbook.Sheets["Migration Input"];
  expect(cellValue(sheet, "H8")).toBe("IC-1001");
  expect(cellValue(sheet, "G8")).toBe(2);
  expect(cellValue(sheet, "F8")).toBe("Northwind Canada Ltd");
  expect(cellValue(sheet, "E8")).toBe("CA");
  expect(cellValue(sheet, "D8")).toBe("CAD");
  expect(cellValue(sheet, "C8")).toBe(18);
  expect(cellValue(sheet, "B8")).toBe("payables@northwind.example");
  expect(cellValue(sheet, "A8")).toBe("Mapped deterministically");
  expect(cellValue(sheet, "H9")).toBe("IC-1002");
  expect(cellValue(sheet, "E9")).toBe("SE");
  expect(cellValue(sheet, "B9")).toBe("ap@contoso-se.example");
  expect(cellValue(sheet, "H10")).toBe("IC-1003");
  expect(cellValue(sheet, "E10")).toBe("BR");

  await page.getByRole("tab", { name: "Run Response" }).click();
  await expect(page.getByText("targetStartRow\": 8")).toBeVisible();
  await expect(page.locator("pre").filter({ hasText: "Customer Master - Export" })).toBeVisible();
  expect(invocationCount).toBe(1);
  expect(JSON.stringify(latestAgentRequest)).toContain("Customer Master - Export");
  expect(JSON.stringify(latestAgentRequest)).not.toContain("data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,");
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

function createIntricateTemplateWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["SAP Data Migration Template", "", "", "", "", "", "", ""],
      ["Object", "Customer onboarding", "", "", "", "", "", ""],
      ["Instruction", "Rows 1-7 are protected template notes", "", "", "", "", "", ""],
      ["Instruction", "Start loading customer records at row 8", "", "", "", "", "", ""],
      ["Instruction", "Country must be ISO-2", "", "", "", "", "", ""],
      ["Instruction", "Column order differs from source", "", "", "", "", "", ""],
      ["Validation Notes", "Email", "Risk Score", "Currency", "Country", "Customer Name", "Source Row", "Legacy ID"],
      ["", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", ""]
    ]),
    "Migration Input"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Target Header", "Rule"],
      ["Legacy ID", "Required"],
      ["Country", "ISO-2"],
      ["Validation Notes", "Generated"]
    ]),
    "Template Rules"
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

function createIntricateSourceWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Customer Number", "Legal Name", "Country Name", "Local Currency", "Risk Rating", "Primary Email", "Lifecycle Status", "Created On"],
      ["IC-1001", "Northwind Canada Ltd", "Canada", "CAD", 18, "payables@northwind.example", "Active", "2026-04-01"],
      ["IC-1002", "Contoso Sweden AB", "Sweden", "SEK", 44, "ap@contoso-se.example", "Active", "2026-04-02"],
      ["IC-1003", "Fabrikam Brasil Ltda", "Brazil", "BRL", 62, "finance@fabrikam-br.example", "Review", "2026-04-03"]
    ]),
    "Customer Master - Export"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Order", "Customer Number", "Amount"],
      ["SO-2001", "IC-1001", 3000],
      ["SO-2002", "IC-1002", 4500]
    ]),
    "Orders"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Note ID", "Text"],
      ["N-1", "Should not be mapped"]
    ]),
    "Notes"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Country Name", "ISO-2"],
      ["Canada", "CA"],
      ["Sweden", "SE"],
      ["Brazil", "BR"]
    ]),
    "Reference"
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

async function downloadWorkbook(page: import("@playwright/test").Page) {
  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download File" }).click();
  const download = await downloadPromise;
  const path = await download.path();
  if (!path) {
    throw new Error("Downloaded workbook path was not available.");
  }
  return XLSX.readFile(path);
}

function cellValue(sheet: XLSX.WorkSheet, cell: string) {
  return sheet[cell]?.v;
}

function readRequestBody(request: NodeJS.ReadableStream) {
  return new Promise<string>((resolve, reject) => {
    const chunks: Buffer[] = [];
    request.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    request.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    request.on("error", reject);
  });
}
