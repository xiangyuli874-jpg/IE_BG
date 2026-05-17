import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "E:/AI/gongshibiao/工时测量表+员工负荷山积表.xlsx";
const outputDir = "E:/AI/gongshibiao/.codex_tmp";

const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheetSummaries = [];
for (const sheet of workbook.worksheets.items) {
  const used = await workbook.inspect({
    kind: "table",
    range: `${sheet.name}!A1:AZ120`,
    include: "values,formulas",
    tableMaxRows: 120,
    tableMaxCols: 52,
  });
  sheetSummaries.push({
    name: sheet.name,
    preview: used.ndjson,
  });
}

const formulaHits = await workbook.inspect({
  kind: "match",
  searchTerm: "=",
  options: { useRegex: false, maxResults: 5000, searchIn: "formulas" },
  summary: "formula cells",
});

const chartKeywordHits = await workbook.inspect({
  kind: "match",
  searchTerm: "山积|负荷|工时|节拍|CT|UPH",
  options: { useRegex: true, maxResults: 500 },
  summary: "keywords",
});

await fs.writeFile(
  path.join(outputDir, "workbook_summary.json"),
  JSON.stringify(
    {
      sheets: sheetSummaries.map((s) => ({
        name: s.name,
        preview: s.preview,
      })),
      formulaHits: formulaHits.ndjson,
      keywordHits: chartKeywordHits.ndjson,
    },
    null,
    2,
  ),
  "utf8",
);

console.log(
  JSON.stringify(
    {
      sheetNames: sheetSummaries.map((s) => s.name),
      formulaHitExcerpt: formulaHits.ndjson.split("\n").slice(0, 40),
      keywordHitExcerpt: chartKeywordHits.ndjson.split("\n").slice(0, 80),
    },
    null,
    2,
  ),
);
