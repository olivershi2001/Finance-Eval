import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [payloadPath, outputPath, qaDir] = process.argv.slice(2);
if (!payloadPath || !outputPath || !qaDir) {
  throw new Error(
    "Usage: node build_model_candidate.mjs <payload.json> <output.xlsx> <qa-dir>",
  );
}

const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const candidate = payload.candidate;
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(qaDir, { recursive: true });

const colors = {
  navy: "#17365D",
  blue: "#4472C4",
  paleBlue: "#D9EAF7",
  teal: "#0F766E",
  paleTeal: "#DDEBF7",
  gold: "#FFF2CC",
  green: "#E2F0D9",
  greenText: "#006100",
  red: "#FCE4D6",
  redText: "#9C0006",
  gray: "#F2F2F2",
  darkGray: "#404040",
  white: "#FFFFFF",
  black: "#000000",
};

const currencyFormat = payload.currency_format || "$#,##0;[Red]($#,##0);-";
const percentFormat = "0.0%;[Red](0.0%);-";
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
  "Bridge check",
];

function safeText(value) {
  if (typeof value === "string" && value.startsWith("=")) return `'${value}`;
  return value;
}

function rowValues(rows) {
  return rows.map((row) => [
    safeText(row.period),
    safeText(row.slice_type),
    safeText(row.slice_value),
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

function titleBand(sheet, range, title, subtitle) {
  sheet.showGridLines = false;
  sheet.getRange(range).merge();
  const topLeft = range.split(":")[0];
  sheet.getRange(topLeft).values = [[safeText(title)]];
  sheet.getRange(range).format = {
    fill: colors.navy,
    font: {
      bold: true,
      color: colors.white,
      size: 16,
      name: "Aptos Display",
    },
    rowHeight: 30,
    verticalAlignment: "center",
  };
  const endColumn = range.split(":")[1].replace(/[0-9]/g, "");
  sheet.getRange(`A2:${endColumn}2`).merge();
  sheet.getRange("A2").values = [[safeText(subtitle)]];
  sheet.getRange(`A2:${endColumn}2`).format = {
    fill: colors.paleBlue,
    font: { italic: true, color: colors.darkGray, size: 10, name: "Aptos" },
    rowHeight: 22,
    wrapText: true,
  };
}

function styleCubeSheet(sheet, rows, title, tableName) {
  titleBand(
    sheet,
    "A1:N1",
    title,
    `${payload.output_origin || "Model-generated output"} • ${candidate.currency_unit} • ${candidate.period_basis} • losses are positive magnitudes`,
  );
  sheet.getRange("A4:N4").values = [outputHeaders];
  sheet.getRange("A4:N4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white, size: 9, name: "Aptos" },
    wrapText: true,
    rowHeight: 30,
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  if (rows.length > 0) {
    const lastRow = rows.length + 4;
    sheet.getRangeByIndexes(4, 0, rows.length, 14).values = rowValues(rows);
    sheet.getRange(`N5`).formulas = [["=J5-(D5+E5+F5+H5-G5-I5)"]];
    sheet.getRange(`N5:N${lastRow}`).fillDown();
    sheet.getRange(`A5:C${lastRow}`).format.horizontalAlignment = "left";
    sheet.getRange(`D5:N${lastRow}`).format.horizontalAlignment = "right";
    sheet.getRange(`D5:J${lastRow}`).format.numberFormat = currencyFormat;
    sheet.getRange(`K5:M${lastRow}`).format.numberFormat = percentFormat;
    sheet.getRange(`N5:N${lastRow}`).format.numberFormat =
      "$#,##0.00;[Red]($#,##0.00);-";
    sheet.getRange(`N5:N${lastRow}`).conditionalFormats.add("cellIs", {
      operator: "equal",
      formula: 0,
      format: { fill: colors.green, font: { color: colors.greenText } },
    });
    sheet.getRange(`N5:N${lastRow}`).conditionalFormats.add("cellIs", {
      operator: "notEqual",
      formula: 0,
      format: { fill: colors.red, font: { color: colors.redText } },
    });
    for (let row = 5; row <= lastRow; row += 1) {
      sheet.getRange(`A${row}:N${row}`).format.fill =
        row % 2 ? colors.paleTeal : colors.white;
      sheet.getRange(`A${row}:N${row}`).format.font = {
        name: "Aptos",
        size: 9,
        color: colors.black,
      };
    }
    const table = sheet.tables.add(`A4:N${lastRow}`, true, tableName);
    table.style = "TableStyleMedium2";
  } else {
    sheet.getRange("A5:N6").merge();
    sheet.getRange("A5").values = [["No rows were returned by the candidate model."]];
    sheet.getRange("A5:N6").format = {
      fill: colors.red,
      font: { bold: true, color: colors.redText, name: "Aptos" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
    };
  }
  [14, 13, 18, 16, 11, 11, 11, 12, 11, 16, 14, 14, 18, 18].forEach(
    (width, index) => {
      sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
    },
  );
  sheet.freezePanes.freezeRows(4);
}

function writeSectionHeader(sheet, row, title) {
  sheet.getRange(`A${row}:H${row}`).merge();
  sheet.getRange(`A${row}`).values = [[safeText(title)]];
  sheet.getRange(`A${row}:H${row}`).format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white, name: "Aptos" },
    rowHeight: 22,
  };
}

function latestOverallRow(rows) {
  let latestIndex = -1;
  for (let index = 0; index < rows.length; index += 1) {
    if (rows[index].slice_type === "overall") latestIndex = index;
  }
  return latestIndex < 0 ? null : latestIndex + 5;
}

function buildExecutiveSummary(sheet) {
  titleBand(
    sheet,
    "A1:H1",
    candidate.workbook_title || "Customer Cube — Model Output",
    `Candidate: ${payload.candidate_model.model || "unspecified"} • Generated ${payload.generated_at} • Contract ${payload.candidate_valid ? "valid" : "invalid"}`,
  );
  sheet.getRange("A4:H4").format.fill = colors.white;
  const latestRow = latestOverallRow(candidate.quarterly_rows);
  const cards = [
    ["Closing revenue", "J"],
    ["New", "E"],
    ["Churn", "I"],
    ["GRR", "K"],
    ["NRR", "L"],
    ["TTM GRR", "M"],
  ];
  cards.forEach(([label, column], index) => {
    const cardColumn = index + 1;
    const cell = sheet.getCell(3, cardColumn);
    cell.values = [[label]];
    cell.format = {
      fill: colors.gray,
      font: { bold: true, color: colors.darkGray, size: 9 },
      horizontalAlignment: "center",
    };
    const valueCell = sheet.getCell(4, cardColumn);
    if (latestRow) {
      valueCell.formulas = [[
        `=IF(ISBLANK('Quarterly'!${column}${latestRow}),\"N/A\",'Quarterly'!${column}${latestRow})`,
      ]];
    } else {
      valueCell.values = [["N/A"]];
    }
    valueCell.format = {
      fill: colors.white,
      font: { bold: true, color: colors.black, size: 13 },
      horizontalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: "#BFBFBF" },
    };
    valueCell.format.numberFormat = ["K", "L", "M"].includes(column)
      ? percentFormat
      : currencyFormat;
  });
  sheet.getRange("B4:G5").format.rowHeight = 24;

  writeSectionHeader(sheet, 7, "Executive assessment");
  sheet.getRange("A8:H10").merge();
  sheet.getRange("A8").values = [[safeText(candidate.summary)]];
  sheet.getRange("A8:H10").format = {
    fill: colors.white,
    font: { name: "Aptos", size: 10, color: colors.black },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "outside", style: "thin", color: "#D9D9D9" },
  };

  let row = 12;
  writeSectionHeader(sheet, row, "Prioritized insights");
  row += 1;
  const insights = candidate.insights.length
    ? candidate.insights
    : ["No insights were returned by the candidate model."];
  for (let index = 0; index < insights.length; index += 1) {
    sheet.getRange(`A${row}`).values = [[index + 1]];
    sheet.getRange(`B${row}:H${row}`).merge();
    sheet.getRange(`B${row}`).values = [[safeText(insights[index])]];
    sheet.getRange(`A${row}:H${row}`).format = {
      fill: index % 2 ? colors.white : colors.paleTeal,
      font: { name: "Aptos", size: 10, color: colors.black },
      wrapText: true,
      rowHeight: 32,
      verticalAlignment: "center",
    };
    row += 1;
  }
  row += 1;
  writeSectionHeader(sheet, row, "Diligence questions");
  row += 1;
  const questions = candidate.diligence_questions.length
    ? candidate.diligence_questions
    : ["No diligence questions were returned by the candidate model."];
  for (let index = 0; index < questions.length; index += 1) {
    sheet.getRange(`A${row}`).values = [[index + 1]];
    sheet.getRange(`B${row}:H${row}`).merge();
    sheet.getRange(`B${row}`).values = [[safeText(questions[index])]];
    sheet.getRange(`A${row}:H${row}`).format = {
      fill: index % 2 ? colors.white : colors.gold,
      font: { name: "Aptos", size: 10, color: colors.black },
      wrapText: true,
      rowHeight: 32,
      verticalAlignment: "center",
    };
    row += 1;
  }
  [5, 18, 18, 18, 18, 18, 18, 18].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(2);
  return row;
}

function buildDataQuality(sheet) {
  titleBand(
    sheet,
    "A1:E1",
    "Data Quality",
    payload.quality_origin || "Issues reported by the candidate model; affected rows are model-generated claims and must be reviewed.",
  );
  const headers = ["Issue", "Severity", "Affected rows", "Impact", "Treatment"];
  sheet.getRange("A4:E4").values = [headers];
  sheet.getRange("A4:E4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    rowHeight: 24,
    horizontalAlignment: "center",
  };
  const rows = candidate.data_quality;
  if (rows.length) {
    sheet.getRangeByIndexes(4, 0, rows.length, 5).values = rows.map((item) => [
      safeText(item.issue),
      safeText(item.severity),
      item.affected_rows,
      safeText(item.impact),
      safeText(item.treatment),
    ]);
    const lastRow = rows.length + 4;
    sheet.getRange(`A5:E${lastRow}`).format = {
      font: { name: "Aptos", size: 10 },
      wrapText: true,
      verticalAlignment: "top",
    };
    rows.forEach((item, index) => {
      const estimatedLines = Math.max(
        Math.ceil(item.issue.length / 28),
        Math.ceil(item.impact.length / 42),
        Math.ceil(item.treatment.length / 48),
      );
      sheet.getRange(`A${index + 5}:E${index + 5}`).format.rowHeight = Math.min(
        126,
        Math.max(42, estimatedLines * 18),
      );
    });
    sheet.getRange(`B5:B${lastRow}`).conditionalFormats.add("containsText", {
      text: "Critical",
      format: { fill: colors.red, font: { color: colors.redText, bold: true } },
    });
    sheet.getRange(`B5:B${lastRow}`).conditionalFormats.add("containsText", {
      text: "Review",
      format: { fill: colors.gold, font: { color: "#7F6000", bold: true } },
    });
    const table = sheet.tables.add(`A4:E${lastRow}`, true, "DataQualityTable");
    table.style = "TableStyleMedium2";
  } else {
    sheet.getRange("A5:E6").merge();
    sheet.getRange("A5").values = [["No data-quality findings were returned."]];
    sheet.getRange("A5:E6").format = {
      fill: colors.gold,
      font: { bold: true, color: "#7F6000" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
    };
  }
  [28, 13, 14, 42, 48].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
}

function buildAssumptions(sheet) {
  titleBand(
    sheet,
    "A1:C1",
    "Assumptions and Conventions",
    "Model-stated assumptions are shown verbatim and should be checked against source data and methodology.",
  );
  sheet.getRange("A4:C4").values = [["#", "Assumption or convention", "Review status"]];
  sheet.getRange("A4:C4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    rowHeight: 24,
  };
  const assumptions = candidate.assumptions.length
    ? candidate.assumptions
    : ["No assumptions were returned by the candidate model."];
  const values = assumptions.map((item, index) => [index + 1, safeText(item), "Review"]);
  sheet.getRangeByIndexes(4, 0, values.length, 3).values = values;
  const lastRow = values.length + 4;
  sheet.getRange(`A5:C${lastRow}`).format = {
    font: { name: "Aptos", size: 10 },
    wrapText: true,
    rowHeight: 36,
    verticalAlignment: "center",
  };
  sheet.getRange(`C5:C${lastRow}`).format.fill = colors.gold;
  [6, 80, 16].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
}

function buildSourceData(sheet) {
  titleBand(
    sheet,
    "A1:D1",
    "Source Data",
    `${payload.source_file.path} • SHA-256 ${payload.source_file.sha256} • original sheet ${payload.source_sheet}`,
  );
  const rows = payload.source_rows.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      if (
        rowIndex > 0 &&
        columnIndex === 1 &&
        typeof value === "string" &&
        /^\d{4}-\d{2}-\d{2}/.test(value)
      ) {
        return new Date(value);
      }
      return safeText(value);
    }),
  );
  if (!rows.length) return;
  sheet.getRangeByIndexes(3, 0, rows.length, rows[0].length).values = rows;
  const lastRow = rows.length + 3;
  sheet.getRange("A4:D4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    rowHeight: 24,
  };
  sheet.getRange(`A5:D${lastRow}`).format.font = { name: "Aptos", size: 9 };
  sheet.getRange(`B5:B${lastRow}`).format.numberFormat = "mmm yyyy";
  sheet.getRange(`C5:C${lastRow}`).format.numberFormat = currencyFormat;
  const table = sheet.tables.add(`A4:D${lastRow}`, true, "SourceRevenueTable");
  table.style = "TableStyleMedium2";
  [22, 14, 15, 18].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
}

function buildChecks(sheet) {
  titleBand(
    sheet,
    "A1:F1",
    "Model Output Checks",
    "Visible automated checks support review; they do not replace the independent LLM judge.",
  );
  sheet.getRange("A4:F4").values = [[
    "Check",
    "Actual",
    "Expected",
    "Difference",
    "Status",
    "Notes",
  ]];
  sheet.getRange("A4:F4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
    rowHeight: 24,
  };
  const qLast = candidate.quarterly_rows.length + 4;
  const aLast = candidate.annual_rows.length + 4;
  const checkRows = [
    ["Structured response contract", payload.candidate_valid ? 1 : 0, 1, null, null, payload.validation_error || "Schema-valid response"],
    ["Quarterly output rows", candidate.quarterly_rows.length, "> 0", null, null, "Includes all model-returned cuts"],
    ["Annual output rows", candidate.annual_rows.length, "> 0", null, null, "Includes all model-returned cuts"],
    ["Quarterly bridge failures", null, 0, null, null, "Non-zero bridge checks"],
    ["Annual bridge failures", null, 0, null, null, "Non-zero bridge checks"],
  ];
  sheet.getRange("A5:F9").values = checkRows;
  if (candidate.quarterly_rows.length) {
    sheet.getRange("B8").formulas = [[`=COUNTIF(Quarterly!N5:N${qLast},"<>0")`]];
  }
  if (candidate.annual_rows.length) {
    sheet.getRange("B9").formulas = [[`=COUNTIF(Annual!N5:N${aLast},"<>0")`]];
  }
  sheet.getRange("D5:D9").formulas = [
    ["=B5-C5"],
    ["=IF(B6>0,0,1)"],
    ["=IF(B7>0,0,1)"],
    ["=B8-C8"],
    ["=B9-C9"],
  ];
  sheet.getRange("E5:E9").formulas = [
    ["=IF(D5=0,\"OK\",\"REVIEW\")"],
    ["=IF(D6=0,\"OK\",\"REVIEW\")"],
    ["=IF(D7=0,\"OK\",\"REVIEW\")"],
    ["=IF(D8=0,\"OK\",\"REVIEW\")"],
    ["=IF(D9=0,\"OK\",\"REVIEW\")"],
  ];
  sheet.getRange("A5:F9").format = {
    font: { name: "Aptos", size: 10 },
    wrapText: true,
    rowHeight: 32,
    verticalAlignment: "center",
  };
  sheet.getRange("E5:E9").conditionalFormats.add("containsText", {
    text: "OK",
    format: { fill: colors.green, font: { color: colors.greenText, bold: true } },
  });
  sheet.getRange("E5:E9").conditionalFormats.add("containsText", {
    text: "REVIEW",
    format: { fill: colors.red, font: { color: colors.redText, bold: true } },
  });
  [30, 16, 16, 14, 14, 44].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
}

function buildModelResponse(sheet) {
  titleBand(
    sheet,
    "A1:B1",
    "Raw Model Response",
    "Shown because the response could not be converted into the required workbook contract.",
  );
  sheet.getRange("A4:B4").values = [["Chunk", "Raw response"]];
  sheet.getRange("A4:B4").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.white },
  };
  const text = payload.raw_response || "No response text was available.";
  const chunks = [];
  for (let index = 0; index < text.length; index += 30000) {
    chunks.push([chunks.length + 1, safeText(text.slice(index, index + 30000))]);
  }
  sheet.getRangeByIndexes(4, 0, chunks.length, 2).values = chunks;
  sheet.getRange(`A5:B${chunks.length + 4}`).format = {
    font: { name: "Aptos", size: 9 },
    wrapText: true,
    verticalAlignment: "top",
    rowHeight: 120,
  };
  sheet.getRange("A1").format.columnWidth = 10;
  sheet.getRange("B1").format.columnWidth = 120;
  sheet.freezePanes.freezeRows(4);
}

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Executive Summary");
const quarterly = workbook.worksheets.add("Quarterly");
const annual = workbook.worksheets.add("Annual");
const dataQuality = workbook.worksheets.add("Data Quality");
const assumptions = workbook.worksheets.add("Assumptions");
const checks = workbook.worksheets.add("Checks");
const source = workbook.worksheets.add("Source Data");
const modelResponse = payload.candidate_valid
  ? null
  : workbook.worksheets.add("Model Response");

const summaryLastRow = buildExecutiveSummary(summary);
styleCubeSheet(
  quarterly,
  candidate.quarterly_rows,
  payload.output_origin ? "Quarterly Customer Cube — " + payload.output_origin : "Quarterly Customer Cube — Candidate Model Output",
  "QuarterlyCandidateTable",
);
styleCubeSheet(
  annual,
  candidate.annual_rows,
  payload.output_origin ? "Annual Customer Cube — " + payload.output_origin : "Annual Customer Cube — Candidate Model Output",
  "AnnualCandidateTable",
);
buildDataQuality(dataQuality);
buildAssumptions(assumptions);
buildChecks(checks);
buildSourceData(source);
if (modelResponse) buildModelResponse(modelResponse);

workbook.recalculate();

const qaTargets = [
  ["Executive Summary", `A1:H${Math.min(summaryLastRow, 35)}`],
  ["Quarterly", "A1:N28"],
  ["Annual", "A1:N28"],
  ["Data Quality", "A1:E20"],
  ["Assumptions", "A1:C20"],
  ["Checks", "A1:F12"],
  ["Source Data", "A1:D32"],
];
if (modelResponse) qaTargets.push(["Model Response", "A1:B12"]);

const inspections = [];
for (const [sheetName, range] of qaTargets) {
  const inspected = await workbook.inspect({
    kind: "table",
    range: `${sheetName}!${range}`,
    include: "values,formulas",
    tableMaxRows: 35,
    tableMaxCols: 14,
  });
  inspections.push(inspected.ndjson);
  const preview = await workbook.render({
    sheetName,
    range,
    scale: 1.2,
    format: "png",
  });
  const safeName = sheetName.toLowerCase().replaceAll(" ", "_");
  await fs.writeFile(
    path.join(qaDir, `${safeName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!",
  options: { useRegex: true, maxResults: 500 },
  summary: "final formula error scan",
});
await fs.writeFile(
  path.join(qaDir, "inspection.txt"),
  `${inspections.join("\n")}\n${errors.ndjson}\n`,
);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
