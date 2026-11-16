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
  const summary = workbook.worksheets.add("Executive Summary");
  const quarterly = workbook.worksheets.add("Quarterly");
  const annual = workbook.worksheets.add("Annual");
  const dataQuality = workbook.worksheets.add("Data Quality");
  const definitions = workbook.worksheets.add("Metric Definitions");
  const rubric = workbook.worksheets.add("Scoring Rubric");
  const source = workbook.worksheets.add("Source Data");

  const quarterlyRows = expectedRows.filter((row) => row.grain === "quarterly");
  const annualRows = expectedRows.filter((row) => row.grain === "annual");
  styleOutputSheet(quarterly, quarterlyRows, "Quarterly Customer Cube — Reference Output");
  styleOutputSheet(annual, annualRows, "Annual Customer Cube — Reference Output");

  const latestQuarterOverall = quarterlyRows.filter(
    (row) => row.slice_type === "overall",
  ).at(-1);
  const latestQuarterProducts = quarterlyRows.filter(
    (row) => row.period === latestQuarterOverall.period && row.slice_type === "product",
  );
  const latestQuarterCustomers = quarterlyRows.filter(
    (row) => row.period === latestQuarterOverall.period && row.slice_type === "customer",
  );
  const rowNumber = (rows, row) => rows.indexOf(row) + 5;
  const overallRow = rowNumber(quarterlyRows, latestQuarterOverall);
  const lowestGrrProduct = latestQuarterProducts
    .filter((row) => row.gross_retention !== null)
    .sort((left, right) => left.gross_retention - right.gross_retention)[0];
  const largestChurnCustomer = [...latestQuarterCustomers]
    .sort((left, right) => right.churn - left.churn)[0];
  const topCustomers = [...latestQuarterCustomers]
    .sort((left, right) => right.closing_revenue - left.closing_revenue)
    .slice(0, 5);
  const lowestGrrRow = rowNumber(quarterlyRows, lowestGrrProduct);
  const largestChurnRow = rowNumber(quarterlyRows, largestChurnCustomer);
  const topCustomerCells = topCustomers
    .map((row) => `'Quarterly'!J${rowNumber(quarterlyRows, row)}`)
    .join(",");

  summary.showGridLines = false;
  summary.getRange("A1:H1").merge();
  summary.getRange("A1").values = [["Customer Cube — Executive Summary"]];
  summary.getRange("A1:H1").format = {
    fill: "#18324A",
    font: { bold: true, color: "#FFFFFF", size: 18, name: "Aptos Display" },
    rowHeight: 32,
  };
  summary.getRange("A2:H2").merge();
  summary.getRange("A2").values = [[
    `Latest period: ${latestQuarterOverall.period} • Source: 1,000 customer-product-month rows • Calendar-period revenue basis`,
  ]];
  summary.getRange("A2:H2").format = {
    fill: "#F2F4F7",
    font: { italic: true, color: "#667085", size: 10 },
    rowHeight: 22,
  };
  summary.getRange("A4:H4").values = [[
    "Period", latestQuarterOverall.period,
    "Closing revenue", null,
    "Gross retention", null,
    "Net retention", null,
  ]];
  summary.getRange("D4").formulas = [[`='Quarterly'!J${overallRow}`]];
  summary.getRange("F4").formulas = [[`='Quarterly'!K${overallRow}`]];
  summary.getRange("H4").formulas = [[`='Quarterly'!L${overallRow}`]];
  summary.getRange("A5:H5").values = [[
    "New", null,
    "Churn", null,
    "TTM gross retention", null,
    "Top-five concentration", null,
  ]];
  summary.getRange("B5").formulas = [[`='Quarterly'!E${overallRow}`]];
  summary.getRange("D5").formulas = [[`='Quarterly'!I${overallRow}`]];
  summary.getRange("F5").formulas = [[`='Quarterly'!M${overallRow}`]];
  summary.getRange("H5").formulas = [[`=SUM(${topCustomerCells})/'Quarterly'!J${overallRow}`]];
  summary.getRange("A4:H5").format = {
    borders: { preset: "all", style: "thin", color: "#D0D5DD" },
    rowHeight: 28,
  };
  summary.getRange("A4:A5").format = { fill: "#E8F1F7", font: { bold: true, color: "#18324A" } };
  summary.getRange("C4:C5").format = { fill: "#E8F1F7", font: { bold: true, color: "#18324A" } };
  summary.getRange("E4:E5").format = { fill: "#E8F1F7", font: { bold: true, color: "#18324A" } };
  summary.getRange("G4:G5").format = { fill: "#E8F1F7", font: { bold: true, color: "#18324A" } };
  summary.getRange("B4:B4").format.font = { bold: true, color: "#000000" };
  summary.getRange("D4:D5").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  summary.getRange("B5:B5").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  summary.getRange("F4:H5").format.numberFormat = "0.0%;[Red](0.0%);-";

  summary.getRange("A7:H7").merge();
  summary.getRange("A7").values = [["Four Quantified Review Insights"]];
  summary.getRange("A7:H7").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, rowHeight: 22 };
  const insightFormulas = [
    `="Latest-quarter closing revenue is "&TEXT(D4,"$#,##0")&", with GRR of "&TEXT(F4,"0.0%")&" and NRR of "&TEXT(H4,"0.0%")&"."`,
    `="Lowest product GRR is "&TEXT('Quarterly'!K${lowestGrrRow},"0.0%")&" for "&'Quarterly'!C${lowestGrrRow}&", indicating the clearest product-level retention pressure."`,
    `="Largest customer churn in the latest quarter is "&TEXT('Quarterly'!I${largestChurnRow},"$#,##0")&" from "&'Quarterly'!C${largestChurnRow}&"."`,
    `="Top-five customer concentration is "&TEXT(H5,"0.0%")&" of latest-quarter closing revenue."`,
  ];
  for (let index = 0; index < insightFormulas.length; index += 1) {
    const row = 8 + index;
    summary.getRange(`A${row}`).values = [[index + 1]];
    summary.getRange(`B${row}:H${row}`).merge();
    summary.getRange(`B${row}`).formulas = [[insightFormulas[index]]];
    summary.getRange(`A${row}:H${row}`).format = {
      borders: { preset: "outside", style: "thin", color: "#D0D5DD" },
      wrapText: true,
      rowHeight: 32,
      verticalAlignment: "center",
    };
    summary.getRange(`A${row}`).format = { fill: "#E8F1F7", font: { bold: true, color: "#18324A" }, horizontalAlignment: "center" };
  }

  summary.getRange("A13:H13").merge();
  summary.getRange("A13").values = [["Suggested Diligence Follow-ups"]];
  summary.getRange("A13:H13").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, rowHeight: 22 };
  const followUps = [
    `Investigate renewal, pricing, and usage drivers for ${lowestGrrProduct.slice_value}, the product with the lowest latest-quarter GRR.`,
    `Review the loss history and replacement pipeline for ${largestChurnCustomer.slice_value}, the largest latest-quarter churn event.`,
    "Separate durable same-product upsell from cross-sell when assessing the quality and repeatability of NRR.",
    "Review top-five customer contract terms and renewal timing to assess concentration and downside risk.",
  ];
  for (let index = 0; index < followUps.length; index += 1) {
    const row = 14 + index;
    summary.getRange(`A${row}`).values = [[index + 1]];
    summary.getRange(`B${row}:H${row}`).merge();
    summary.getRange(`B${row}`).values = [[followUps[index]]];
    summary.getRange(`A${row}:H${row}`).format = { wrapText: true, rowHeight: 30, verticalAlignment: "center" };
    summary.getRange(`A${row}`).format = { font: { bold: true, color: "#0F766E" }, horizontalAlignment: "center" };
  }
  summary.getRange("A19:H20").merge();
  summary.getRange("A19").values = [[
    "Human-review boundary: the workbook supports analysis and diligence prioritization; final investment interpretation and external use require human approval.",
  ]];
  summary.getRange("A19:H20").format = { fill: "#FFF2CC", font: { italic: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
  [16, 18, 16, 18, 18, 18, 20, 20].forEach(
    (width, index) => summary.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width,
  );
  summary.freezePanes.freezeRows(2);

  source.showGridLines = false;
  source.getRange("A1:F1").values = [[
    "Customer name", "Month", "Amount", "Product", "Aggregation key", "Additional duplicate key flag",
  ]];
  source.getRangeByIndexes(1, 0, rawRows.length, 4).values = rawRows;
  source.getRange("E2").formulas = [["=A2&\"|\"&TEXT(B2,\"yyyymm\")&\"|\"&D2"]];
  source.getRange(`E2:E${rawRows.length + 1}`).fillDown();
  source.getRange("F2").formulas = [["=IF(COUNTIF($E$2:E2,E2)>1,1,0)"]];
  source.getRange(`F2:F${rawRows.length + 1}`).fillDown();
  source.getRange("A1:F1").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF" }, rowHeight: 26, wrapText: true };
  source.getRange(`B2:B${rawRows.length + 1}`).format.numberFormat = "mmm yyyy";
  source.getRange(`C2:C${rawRows.length + 1}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
  source.getRange(`A1:F${rawRows.length + 1}`).format.font = { name: "Aptos", size: 9 };
  [20, 13, 14, 16, 42, 20].forEach(
    (width, index) => source.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width,
  );
  const sourceTable = source.tables.add(`A1:F${rawRows.length + 1}`, true, "ReferenceSourceTable");
  sourceTable.style = "TableStyleMedium2";
  source.freezePanes.freezeRows(1);

  dataQuality.showGridLines = false;
  dataQuality.getRange("A1:F1").merge();
  dataQuality.getRange("A1").values = [["Source Data Quality Review"]];
  dataQuality.getRange("A1:F1").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 30 };
  dataQuality.getRange("A2:F2").merge();
  dataQuality.getRange("A2").values = [["Checks are calculated from the Source Data tab. Duplicate keys are valid split transactions and are aggregated by sum."]];
  dataQuality.getRange("A2:F2").format = { fill: "#F2F4F7", font: { italic: true, color: "#667085" }, wrapText: true, rowHeight: 26 };
  dataQuality.getRange("A4:F4").values = [["Check", "Count", "Status", "Treatment", "Evidence", "Reviewer note"]];
  dataQuality.getRange("A5:A11").values = [
    ["Input rows"],
    ["Additional duplicate customer-month-product keys"],
    ["Blank customer names"],
    ["Blank product names"],
    ["Negative amounts"],
    ["Zero amounts"],
    ["Blank months"],
  ];
  dataQuality.getRange("B5:B11").formulas = [
    [`=COUNTA('Source Data'!A2:A${rawRows.length + 1})`],
    [`=SUM('Source Data'!F2:F${rawRows.length + 1})`],
    [`=COUNTBLANK('Source Data'!A2:A${rawRows.length + 1})`],
    [`=COUNTBLANK('Source Data'!D2:D${rawRows.length + 1})`],
    [`=COUNTIF('Source Data'!C2:C${rawRows.length + 1},"<0")`],
    [`=COUNTIF('Source Data'!C2:C${rawRows.length + 1},0)`],
    [`=COUNTBLANK('Source Data'!B2:B${rawRows.length + 1})`],
  ];
  dataQuality.getRange("C5:C11").formulas = [
    ["=\"INFO\""],
    ["=IF(B6=0,\"PASS\",\"REVIEW\")"],
    ["=IF(B7=0,\"PASS\",\"FAIL\")"],
    ["=IF(B8=0,\"PASS\",\"FAIL\")"],
    ["=IF(B9=0,\"PASS\",\"FAIL\")"],
    ["=IF(B10=0,\"PASS\",\"REVIEW\")"],
    ["=IF(B11=0,\"PASS\",\"FAIL\")"],
  ];
  dataQuality.getRange("D5:F11").values = [
    ["Control total", "Source Data!A2:A1001", "Expected 1,000 rows"],
    ["Aggregate identical keys by sum", "Source Data!E:F", "Five split-transaction keys are intentional"],
    ["Reject blank customer labels", "Source Data!A:A", "None observed"],
    ["Reject blank product labels", "Source Data!D:D", "None observed"],
    ["Reject negative values for this movement convention", "Source Data!C:C", "None observed"],
    ["Retain zeroes but exclude them from active revenue", "Source Data!C:C", "None observed"],
    ["Require a valid calendar month", "Source Data!B:B", "None observed"],
  ];
  dataQuality.getRange("A4:F4").format = { fill: "#4472C4", font: { bold: true, color: "#FFFFFF" }, rowHeight: 28, wrapText: true };
  dataQuality.getRange("A5:F11").format = { borders: { preset: "inside", style: "thin", color: "#D0D5DD" }, wrapText: true, rowHeight: 32, verticalAlignment: "center" };
  dataQuality.getRange("B5:B11").format.numberFormat = "#,##0";
  dataQuality.getRange("C5:C11").conditionalFormats.add("containsText", { text: "PASS", format: { fill: "#E2F0D9", font: { color: "#006100", bold: true } } });
  dataQuality.getRange("C5:C11").conditionalFormats.add("containsText", { text: "REVIEW", format: { fill: "#FFF2CC", font: { color: "#7F6000", bold: true } } });
  dataQuality.getRange("C5:C11").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: "#FCE4D6", font: { color: "#9C0006", bold: true } } });
  [34, 12, 13, 42, 24, 34].forEach(
    (width, index) => dataQuality.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width,
  );
  dataQuality.freezePanes.freezeRows(4);

  definitions.showGridLines = false;
  definitions.getRange("A1:D1").merge();
  definitions.getRange("A1").values = [["Customer Cube Metric Definitions"]];
  definitions.getRange("A1:D1").format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    rowHeight: 30,
  };
  definitions.getRange("A3:D3").values = [["Metric", "Rule", "Sign", "Scored"]];
  definitions.getRange("A4:D14").values = [
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
    ["Reactivation treatment", "A returning customer is reported under New because Reactivation is not a separate requested output metric; the convention must be disclosed", "Positive", "Yes"],
  ];
  definitions.getRange("A3:D3").format = { fill: "#4472C4", font: { bold: true, color: "#FFFFFF" } };
  definitions.getRange("A3:D14").format.font = { name: "Aptos", size: 10 };
  definitions.getRange("A4:D14").format.wrapText = true;
  definitions.getRange("A1:A14").format.columnWidth = 22;
  definitions.getRange("B1:B14").format.columnWidth = 72;
  definitions.getRange("C1:C14").format.columnWidth = 16;
  definitions.getRange("D1:D14").format.columnWidth = 12;
  definitions.getRange("A4:D14").format.rowHeight = 38;
  definitions.freezePanes.freezeRows(3);

  rubric.showGridLines = false;
  rubric.getRange("A1:I1").merge();
  rubric.getRange("A1").values = [["Customer Cube Finance Review Rubric"]];
  rubric.getRange("A1:I1").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 30 };
  rubric.getRange("A2:I2").merge();
  rubric.getRange("A2").values = [["0–4 anchored scale • weighted score = rating ÷ 4 × weight • critical failure overrides aggregate score"]];
  rubric.getRange("A2:I2").format = { fill: "#F2F4F7", font: { italic: true, color: "#667085" }, rowHeight: 22 };
  rubric.getRange("A4:I4").values = [["ID", "Dimension", "Grader", "Weight", "Critical?", "What it tests", "0 - Unacceptable", "2 - Partial", "4 - Excellent"]];
  rubric.getRange("A5:I11").values = [
    ["C1", "Numerical correctness", "Deterministic", 0.30, "Yes", "Bridges, GRR/NRR/TTM GRR, cuts, concentration, and tie-outs", "Material outputs are wrong or a material bridge does not tie", "Most totals are correct but a material classification, cut, or retention calculation is wrong", "All scored values match recalculated evidence and every material bridge ties"],
    ["C2", "Movement methodology", "Hybrid", 0.15, "Yes", "Customer-level New, Upsell, Downsell, Cross-sell, Churn, product migration, and reactivation convention", "Confuses core movement definitions or classifies only at product-row level", "General method is correct but an edge case or convention is mishandled", "Method is explicit, customer-level, and correct across observed edge cases"],
    ["C3", "Data-quality detection", "Hybrid", 0.10, "No", "Duplicates, missing periods, blanks, invalid amounts, and inconsistent records", "Ignores material data issues", "Finds obvious issues but misses, mishandles, or fails to quantify others", "Finds material issues, quantifies impact, and applies defensible treatment"],
    ["C4", "Traceability", "SME", 0.15, "Yes", "Connection from claims and outputs to source rows, definitions, and calculations", "Claims cannot be reconciled or sources are invented", "Key outputs are reproducible but some claims or conventions lack evidence", "Every material number and claim is traceable to a calculation and source slice"],
    ["C5", "Analytical judgment", "SME", 0.15, "No", "Prioritized, quantified retention, movement, concentration, and customer-quality insights", "Conclusions are generic, unsupported, or misleading", "Observations are useful but have limited quantification or prioritization", "Insights are quantified, prioritized, and anticipate senior-review questions"],
    ["C6", "Practical usefulness", "SME", 0.10, "No", "Reviewer efficiency and concrete diligence follow-ups", "Output requires complete rework", "Useful analyst draft but lacks clear, actionable next steps", "Immediately useful to a deal team with clear next diligence actions"],
    ["C7", "Formatting and presentation", "SME", 0.05, "No", "Number formats, units, labels, layout, source notes, and visible caveats", "Misleading units or periods, clipped content, or unusable tables", "Readable but inconsistent formats, spacing, or labels need cleanup", "Consistently formatted, clearly labeled, traceable, and senior-review-ready"],
  ];
  rubric.getRange("A13:I13").merge();
  rubric.getRange("A13").values = [["Pass threshold: 85 / 100 • Review threshold: 70 / 100 • Human approval required for final investment interpretation and external use"]];
  rubric.getRange("A15:I15").merge();
  rubric.getRange("A15").values = [["Critical Failures"]];
  rubric.getRange("B16:H16").merge();
  rubric.getRange("A16").values = [["ID"]];
  rubric.getRange("B16").values = [["Failure condition"]];
  rubric.getRange("I16").values = [["Severity"]];
  const criticalFailures = [
    ["CF1", "Any material revenue bridge fails to reconcile", "Critical"],
    ["CF2", "Fabricated customer, metric, source, or external fact", "Critical"],
    ["CF3", "Product migration is misclassified as customer churn or expansion with material impact", "Critical"],
    ["CF4", "ARR, revenue, bookings, or billings are materially conflated", "Critical"],
    ["CF5", "Material source data is silently omitted", "Critical"],
    ["CF6", "Scale, period, or a material caveat is misleading or hidden", "Critical"],
  ];
  for (let index = 0; index < criticalFailures.length; index += 1) {
    const row = 17 + index;
    rubric.getRange(`B${row}:H${row}`).merge();
    rubric.getRange(`A${row}`).values = [[criticalFailures[index][0]]];
    rubric.getRange(`B${row}`).values = [[criticalFailures[index][1]]];
    rubric.getRange(`I${row}`).values = [[criticalFailures[index][2]]];
  }
  rubric.getRange("A24:I24").merge();
  rubric.getRange("A24").values = [["Dataset Variants"]];
  rubric.getRange("B25:C25").merge();
  rubric.getRange("D25:H25").merge();
  rubric.getRange("A25").values = [["ID"]];
  rubric.getRange("B25").values = [["Variant"]];
  rubric.getRange("D25").values = [["Expected behavior"]];
  rubric.getRange("I25").values = [["Severity"]];
  const datasetVariants = [
    ["T1", "Product migration with unchanged customer revenue", "Avoids artificial customer churn or expansion", "Critical"],
    ["T2", "Churn followed by reactivation", "Applies and discloses the declared reactivation convention", "High"],
    ["T3", "Customer rename without stable ID", "Flags the identity limitation and does not invent linkage", "High"],
    ["T4", "Duplicate customer-month-product key", "Flags and aggregates valid split transactions without double counting", "High"],
    ["T5", "Missing month", "Flags the gap and avoids automatic churn or reactivation", "High"],
    ["T6", "Zero opening revenue cohort", "Reports retention as null or not applicable", "Medium"],
    ["T7", "Mixed units or period labels", "Normalizes scale and avoids a magnitude or timing misread", "Critical"],
    ["T8", "Truncated churn-volatility chart", "Uses a defensible scale and labels the range", "High"],
  ];
  for (let index = 0; index < datasetVariants.length; index += 1) {
    const row = 26 + index;
    rubric.getRange(`B${row}:C${row}`).merge();
    rubric.getRange(`D${row}:H${row}`).merge();
    rubric.getRange(`A${row}`).values = [[datasetVariants[index][0]]];
    rubric.getRange(`B${row}`).values = [[datasetVariants[index][1]]];
    rubric.getRange(`D${row}`).values = [[datasetVariants[index][2]]];
    rubric.getRange(`I${row}`).values = [[datasetVariants[index][3]]];
  }
  rubric.getRange("A4:I4").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF" }, wrapText: true, rowHeight: 30 };
  rubric.getRange("A5:I11").format = { font: { name: "Aptos", size: 9 }, wrapText: true, rowHeight: 54, verticalAlignment: "top" };
  rubric.getRange("D5:D11").format.numberFormat = "0%";
  rubric.getRange("A13:I13").format = { fill: "#FFF2CC", font: { bold: true, color: "#7F6000" }, rowHeight: 28 };
  rubric.getRange("A15:I15").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, rowHeight: 22 };
  rubric.getRange("A16:I16").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF" } };
  rubric.getRange("A17:I22").format = { wrapText: true, rowHeight: 26, verticalAlignment: "center" };
  rubric.getRange("A24:I24").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, rowHeight: 22 };
  rubric.getRange("A25:I25").format = { fill: "#18324A", font: { bold: true, color: "#FFFFFF" } };
  rubric.getRange("A26:I33").format = { wrapText: true, rowHeight: 32, verticalAlignment: "center" };
  [9, 24, 16, 11, 11, 38, 42, 42, 42].forEach(
    (width, index) => rubric.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width,
  );
  rubric.freezePanes.freezeRows(4);

  const inspections = [];
  for (const [sheetName, range] of [
    ["Executive Summary", "A1:H20"],
    ["Quarterly", "A1:N28"],
    ["Annual", "A1:N28"],
    ["Data Quality", "A1:F11"],
    ["Metric Definitions", "A1:D14"],
    ["Scoring Rubric", "A1:I33"],
    ["Source Data", "A1:F32"],
  ]) {
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
