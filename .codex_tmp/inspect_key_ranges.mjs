import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "E:/AI/gongshibiao/工时测量表+员工负荷山积表.xlsx";
const outputDir = "E:/AI/gongshibiao/.codex_tmp";
const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const ranges = [
  "工时测量记录表!A1:AZ35",
  "标准工时表!A1:Q40",
  "工时分析!A1:U40",
  "平衡率及动作变异分析图!A1:U40",
];

const results = {};
for (const range of ranges) {
  const inspected = await workbook.inspect({
    kind: "table",
    range,
    include: "values,formulas",
    tableMaxRows: 60,
    tableMaxCols: 60,
  });
  results[range] = inspected.ndjson;
}

await fs.writeFile(path.join(outputDir, "key_ranges.json"), JSON.stringify(results, null, 2), "utf8");
console.log(JSON.stringify(results, null, 2));
