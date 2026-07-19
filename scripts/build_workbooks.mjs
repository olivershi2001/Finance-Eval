import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixtureDir = path.join(root, "fixtures");
const outputDir = path.join(root, "outputs", "customer_cube_eval");
const expectedPath = path.join(root, "scenarios", "customer_cube_expected.json");

await fs.mkdir(fixtureDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const products = ["Platform", "Analytics", "Security", "Support", "Automation"];
const rawRows = [];
for (let customerNumber = 1; customerNumber <= 50; customerNumber += 1) {
  const customer = `Customer ${String(customerNumber).padStart(3, "0")}`;
  const baseProductIndex = (customerNumber - 1) % products.length;
  const baseProduct = products[baseProductIndex];
  const secondaryProduct = products[
    (baseProductIndex + (customerNumber % 3) + 1) % products.length
  ];
  const baseAmount = 80 + ((customerNumber - 1) % 10) * 20;
  const startOffset = ((customerNumber - 1) * 7) % 33;
  const splitOpeningTransaction = customerNumber % 10 === 0;

  // Sixteen consecutive base-product months create deterministic customer
  // starts, expansions, contractions, recoveries, and eventual churn.
  for (let activeMonth = 0; activeMonth < 16; activeMonth += 1) {
    const monthIndex = startOffset + activeMonth;
    const year = 2022 + Math.floor(monthIndex / 12);
    const month = monthIndex % 12;
    const monthDate = new Date(Date.UTC(year, month, 1));
    const multiplier = activeMonth <= 4
      ? 1.0
      : activeMonth <= 9
        ? 1.25
        : activeMonth <= 12
          ? 0.9
          : 1.1;
    const amount = Math.round(baseAmount * multiplier);
    if (splitOpeningTransaction && activeMonth === 0) {
      // Split rows test aggregation of duplicate customer/month/product keys.
      const firstPiece = Math.round(amount * 0.4);
      rawRows.push([customer, monthDate, firstPiece, baseProduct]);
      rawRows.push([customer, monthDate, amount - firstPiece, baseProduct]);
    } else {
      rawRows.push([customer, monthDate, amount, baseProduct]);
    }
  }

  // Intermittent second-product purchases create cross-sell and product-level
  // downsell without forcing the whole customer to churn.
  const secondaryMonths = splitOpeningTransaction ? [4, 7, 10] : [4, 7, 10, 13];
  for (const activeMonth of secondaryMonths) {
    const monthIndex = startOffset + activeMonth;
    const year = 2022 + Math.floor(monthIndex / 12);
    const month = monthIndex % 12;
    const monthDate = new Date(Date.UTC(year, month, 1));
    const amount = 25 + (customerNumber % 8) * 10;
    rawRows.push([customer, monthDate, amount, secondaryProduct]);
  }
}

rawRows.sort((left, right) => (
  left[1] - right[1]
  || left[0].localeCompare(right[0])
  || left[3].localeCompare(right[3])
  || left[2] - right[2]
));

if (rawRows.length !== 1000) {
  throw new Error(`Expected exactly 1000 dummy rows, received ${rawRows.length}`);
}

async function buildRawWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Raw Data");
  sheet.showGridLines = false;
  sheet.getRange("A1:D1").values = [["Customer name", "Month", "Amount", "Product"]];
  sheet.getRangeByIndexes(1, 0, rawRows.length, 4).values = rawRows;
  sheet.getRange("A1:D1").format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF" },
    rowHeight: 24,
  };
  sheet.getRange(`B2:B${rawRows.length + 1}`).format.numberFormat = "mmm yyyy";
  sheet.getRange(`C2:C${rawRows.length + 1}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
  sheet.getRange(`A1:D${rawRows.length + 1}`).format.font = { name: "Aptos", size: 10 };
  sheet.getRange("A1:A1").format.columnWidth = 20;
  sheet.getRange("B1:B1").format.columnWidth = 13;
  sheet.getRange("C1:C1").format.columnWidth = 14;
  sheet.getRange("D1:D1").format.columnWidth = 16;
  sheet.freezePanes.freezeRows(1);
  const table = sheet.tables.add(`A1:D${rawRows.length + 1}`, true, "RawRevenueTable");
  table.style = "TableStyleMedium2";
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(fixtureDir, "customer_cube_raw.xlsx"));
  const preview = await workbook.render({ sheetName: "Raw Data", range: "A1:D32", scale: 1.5, format: "png" });
  await fs.writeFile(path.join(outputDir, "qa_raw_data.png"), new Uint8Array(await preview.arrayBuffer()));
}

const outputHeaders = [
  "Period",
  "Slice type",
  "Slice value",
  "Opening revenue",
  "New",
  "Upsell",
  "Downsell",
  "Cross-sell",
  "Churn",
  "Closing revenue",
  "Gross retention",
  "Net retention",
  "TTM gross retention",
  "Reconciliation check",
];

function outputValues(rows) {
  return rows.map((row) => [
    row.period,
    row.slice_type,
    row.slice_value,
    row.opening_revenue,
    row.new,
    row.upsell,
    row.downsell,
    row.cross_sell,
    row.churn,
    row.closing_revenue,
    row.gross_retention,
    row.net_retention,
    row.ttm_gross_retention,
    null,
  ]);
}

function styleOutputSheet(sheet, rows, title) {
  const lastRow = rows.length + 4;
  sheet.showGridLines = false;
  sheet.getRange("A1:N1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:N1").format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", size: 16, name: "Aptos Display" },
    rowHeight: 30,
    verticalAlignment: "center",
  };
  sheet.getRange("A2:N2").merge();
  sheet.getRange("A2").values = [["Frozen gold output • monthly revenue is summed into calendar periods • losses are positive magnitudes"]];
  sheet.getRange("A2:N2").format = {
    fill: "#D9EAF7",
    font: { italic: true, color: "#404040", size: 10 },
    rowHeight: 22,
  };
  sheet.getRange("A4:N4").values = [outputHeaders];
  sheet.getRange("A4:N4").format = {
    fill: "#4472C4",
    font: { bold: true, color: "#FFFFFF", size: 9 },
    wrapText: true,
    rowHeight: 30,
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRangeByIndexes(4, 0, rows.length, 14).values = outputValues(rows);
  sheet.getRange(`D5:J${lastRow}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
  sheet.getRange(`K5:M${lastRow}`).format.numberFormat = "0.0%;[Red](0.0%);-";
  sheet.getRange(`N5`).formulas = [["=J5-(D5+E5+F5+H5-G5-I5)"]];
  sheet.getRange(`N5:N${lastRow}`).fillDown();
  sheet.getRange(`N5:N${lastRow}`).format.numberFormat = "$#,##0.00;[Red]($#,##0.00);-";
  sheet.getRange(`A4:N${lastRow}`).format.font = { name: "Aptos", size: 9 };
  sheet.getRange(`A5:C${lastRow}`).format.horizontalAlignment = "left";
  sheet.getRange(`D5:N${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`N5:N${lastRow}`).conditionalFormats.add("cellIs", {
    operator: "equal",
    formula: 0,
    format: { fill: "#E2F0D9", font: { color: "#006100" } },
  });
  sheet.getRange(`N5:N${lastRow}`).conditionalFormats.add("cellIs", {
    operator: "notEqual",
    formula: 0,
    format: { fill: "#FCE4D6", font: { color: "#9C0006" } },
  });
  const widths = [14, 13, 17, 16, 11, 11, 11, 12, 11, 16, 14, 14, 18, 20];
  widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width);
  sheet.freezePanes.freezeRows(4);
  // The renderer can show transparent cells as black on large annual tables,
  // so the annual reference uses explicit banding instead of a table style.
  if (title.startsWith("Quarterly")) {
    const table = sheet.tables.add(`A4:N${lastRow}`, true, "QuarterlyGoldTable");
    table.style = "TableStyleMedium2";
  }
  // Explicit band fills avoid renderer-dependent transparency on zero/null cells.
  for (let row = 5; row <= lastRow; row += 1) {
    sheet.getRange(`A${row}:N${row}`).format.fill = row % 2 ? "#DDEBF7" : "#FFFFFF";
    sheet.getRange(`A${row}:N${row}`).format.font = { name: "Aptos", size: 9, color: "#000000" };
  }
}

async function buildReferenceWorkbook(expectedRows) {
  const workbook = Workbook.create();
  const quarterly = workbook.worksheets.add("Quarterly");
  const annual = workbook.worksheets.add("Annual");
  const definitions = workbook.worksheets.add("Metric Definitions");
  const rubric = workbook.worksheets.add("Scoring Rubric");

  const quarterlyRows = expectedRows.filter((row) => row.grain === "quarterly");
  const annualRows = expectedRows.filter((row) => row.grain === "annual");
  styleOutputSheet(quarterly, quarterlyRows, "Quarterly Customer Cube — Reference Output");
  styleOutputSheet(annual, annualRows, "Annual Customer Cube — Reference Output");

  definitions.showGridLines = false;
  definitions.getRange("A1:D1").merge();
  definitions.getRange("A1").values = [["Customer Cube Metric Definitions"]];
  definitions.getRange("A1:D1").format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    rowHeight: 30,
  };
  definitions.getRange("A3:D3").values = [["Metric", "Rule", "Sign", "Scored"]];
  definitions.getRange("A4:D13").values = [
    ["Opening revenue", "Prior comparable period revenue at customer × product", "Positive", "Yes"],
    ["New", "Customer had zero prior-period revenue; all current revenue is new", "Positive", "Yes"],
    ["Upsell", "Existing customer and product; current revenue exceeds prior revenue", "Positive", "Yes"],
    ["Downsell", "Existing customer remains active; product revenue decreases or disappears", "Positive loss", "Yes"],
    ["Cross-sell", "Existing customer begins a product absent in the prior period", "Positive", "Yes"],
    ["Churn", "Customer total revenue falls to zero; all prior product revenue churns", "Positive loss", "Yes"],
    ["Closing revenue", "Opening + New + Upsell + Cross-sell − Downsell − Churn", "Positive", "Yes"],
    ["Gross retention", "(Opening − Downsell − Churn) / Opening; capped at 100%", "Rate", "Yes"],
    ["Net retention", "(Opening + Upsell + Cross-sell − Downsell − Churn) / Opening", "Rate", "Yes"],
    ["TTM gross retention", "Gross retention comparing rolling 12 months with the preceding 12 months", "Rate", "Yes"],
  ];
  definitions.getRange("A3:D3").format = { fill: "#4472C4", font: { bold: true, color: "#FFFFFF" } };
  definitions.getRange("A3:D13").format.font = { name: "Aptos", size: 10 };
  definitions.getRange("A4:D13").format.wrapText = true;
  definitions.getRange("A1:A13").format.columnWidth = 22;
  definitions.getRange("B1:B13").format.columnWidth = 72;
  definitions.getRange("C1:C13").format.columnWidth = 16;
  definitions.getRange("D1:D13").format.columnWidth = 12;
  definitions.getRange("A4:D13").format.rowHeight = 34;
  definitions.freezePanes.freezeRows(3);

  rubric.showGridLines = false;
  rubric.getRange("A1:D1").merge();
  rubric.getRange("A1").values = [["Evaluation Scoring Rubric"]];
  rubric.getRange("A1:D1").format = { fill: "#17365D", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 30 };
  rubric.getRange("A3:D3").values = [["Section", "Weight", "What it tests", "Failure examples"]];
  rubric.getRange("A4:D8").values = [
    ["Structure", 0.10, "Valid grains, cuts, numeric values, unique row keys", "Duplicate keys, invalid labels, non-finite values"],
    ["Coverage", 0.20, "All expected period × cut rows, with no invented rows", "Missing Annual tab equivalent or product cuts"],
    ["Movement accuracy", 0.40, "Opening/closing revenue and five movement categories", "Cross-sell mislabeled as New; churn mislabeled as Downsell"],
    ["Retention accuracy", 0.20, "GRR, NRR, and rolling TTM GRR", "New revenue included in NRR; wrong TTM window"],
    ["Reconciliation", 0.10, "Revenue bridge and retention formulas tie", "Closing revenue does not equal the movement bridge"],
  ];
  rubric.getRange("A10:D10").merge();
  rubric.getRange("A10").values = [["Default passing score: 95.0 / 100. Dollar tolerance: $0.01. Rates use 1e-6 relative tolerance."]];
  rubric.getRange("A3:D3").format = { fill: "#4472C4", font: { bold: true, color: "#FFFFFF" } };
  rubric.getRange("A4:D8").format.font = { name: "Aptos", size: 10 };
  rubric.getRange("B4:B8").format.numberFormat = "0%";
  rubric.getRange("A4:D10").format.wrapText = true;
  rubric.getRange("A1:A10").format.columnWidth = 22;
  rubric.getRange("B1:B10").format.columnWidth = 12;
  rubric.getRange("C1:C10").format.columnWidth = 54;
  rubric.getRange("D1:D10").format.columnWidth = 54;
  rubric.getRange("A4:D8").format.rowHeight = 38;
  rubric.getRange("A10:D10").format = { fill: "#FFF2CC", font: { bold: true, color: "#7F6000" }, rowHeight: 28 };

  const inspections = [];
  for (const [sheetName, range] of [["Quarterly", "A1:N28"], ["Annual", "A1:N28"], ["Metric Definitions", "A1:D13"], ["Scoring Rubric", "A1:D10"]]) {
    const inspected = await workbook.inspect({ kind: "table", range: `${sheetName}!${range}`, include: "values,formulas", tableMaxRows: 28, tableMaxCols: 14 });
    inspections.push(inspected.ndjson);
    const preview = await workbook.render({ sheetName, range, scale: 1.25, format: "png" });
    const safeName = sheetName.toLowerCase().replaceAll(" ", "_");
    await fs.writeFile(path.join(outputDir, `qa_${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });
  await fs.writeFile(path.join(outputDir, "qa_inspection.txt"), `${inspections.join("\n")}\n${errors.ndjson}\n`);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, "customer_cube_reference_output.xlsx"));
}

await buildRawWorkbook();
try {
  const expected = JSON.parse(await fs.readFile(expectedPath, "utf8"));
  await buildReferenceWorkbook(expected.rows);
} catch (error) {
  if (error?.code !== "ENOENT") throw error;
}
