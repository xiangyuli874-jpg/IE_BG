import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { buildWorkbookBuffer } from "../netlify/functions/_shared/excel";
import { createSession } from "../netlify/functions/_shared/session";

const modelName = "XQG100-测试";
const session = createSession("a-line", "ordinary-washer-dryer", modelName);
const groupCounts = session.groups.map((group) => group.processes.length).join("/");
if (groupCounts !== "19/24/36/27") {
  throw new Error("default washer-dryer template counts changed");
}
session.parameters.plannedOutput = 420;
session.groups[0].processes.slice(0, 12).forEach((process, index) => {
  process.samples = [12 + index, 12.5 + index, 13 + index, null, null];
  process.people = 1;
});

const result = await buildWorkbookBuffer(session, "group", session.groups[0].id);
const outputPath = path.resolve(".codex_tmp", "verify_export.xlsx");
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, result.buffer);

const zip = await JSZip.loadAsync(result.buffer);
const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
const chartXml = await zip.file("xl/charts/chart1.xml")?.async("string");
const tableXml = await zip.file("xl/tables/table1.xml")?.async("string");
const traceXml = await zip.file("xl/worksheets/sheet6.xml")?.async("string");
const lastRow = result.processCount + 2;

if (!workbookXml?.includes("现场测量记录")) {
  throw new Error("trace sheet missing");
}
if (!chartXml?.includes(`工序测量!$C$3:$C$${lastRow}`)) {
  throw new Error("chart category range was not expanded");
}
if (!tableXml?.includes(`ref="B2:S${lastRow}"`)) {
  throw new Error("tblProcess range was not expanded");
}
if (result.processCount !== 12) {
  throw new Error("blank process names with samples were not exported");
}
if (!result.fileName.includes(modelName) || !result.fileName.includes("工序工时表")) {
  throw new Error("export file name does not include manual model name");
}
if (!traceXml?.includes(modelName)) {
  throw new Error("trace sheet does not include manual model name");
}

console.log(JSON.stringify({ outputPath, fileName: result.fileName, processCount: result.processCount }, null, 2));
