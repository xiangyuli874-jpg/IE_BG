import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "E:/AI/gongshibiao/outputs/worktime_logic_test/工时测量表_山积逻辑测试版.xlsx";
const renderDir = "E:/AI/gongshibiao/.codex_tmp/render_check";
await fs.mkdir(renderDir, { recursive: true });

const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const check = await workbook.inspect({
  kind: "table",
  range: "测试版_山积逻辑!A1:N29",
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 20,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const sheet of workbook.worksheets.items) {
  const blob = await workbook.render({
    sheetName: sheet.name,
    range: sheet.name === "测试版_山积逻辑" ? "A1:U29" : "A1:U42",
    scale: 1,
  });
  console.log(sheet.name, typeof blob, Object.keys(blob ?? {}));
}

console.log("rendered");
