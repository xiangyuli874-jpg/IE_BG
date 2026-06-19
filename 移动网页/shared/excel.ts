import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { findLine } from "../src/data/presets.js";
import type { ExportScope, WorktimeGroup, WorktimeProcess, WorktimeSession } from "../src/types.js";

const TEMPLATE_FILE = "工时测量与负荷山积自动扩展模板_v16.xlsx";
const TEMPLATE_PARTS = ["outputs", "worktime_new_template", TEMPLATE_FILE];

interface ExportProcess {
  group: WorktimeGroup;
  process: WorktimeProcess;
}

interface WorkbookBuildResult {
  buffer: Buffer;
  fileName: string;
  processCount: number;
}

function displayModelName(session: WorktimeSession) {
  return session.modelName?.trim() || "洗烘结构";
}

const formulas = {
  processNo: "ROW()-ROW(tblProcess[#Headers])",
  average:
    'IF(COUNT(tblProcess[[#This Row],[测量1]]:tblProcess[[#This Row],[测量5]])=0,"",AVERAGE(tblProcess[[#This Row],[测量1]]:tblProcess[[#This Row],[测量5]]))',
  standard:
    'IF(tblProcess[[#This Row],[平均实测(s)]]="","",tblProcess[[#This Row],[平均实测(s)]]*(1+参数设置!$B$11)*(1+参数设置!$B$6))',
  load:
    'IF(OR(tblProcess[[#This Row],[标准工时(s)]]="",tblProcess[[#This Row],[人员]]=""),"",tblProcess[[#This Row],[标准工时(s)]]/tblProcess[[#This Row],[人员]])',
  takt: 'IF(tblProcess[[#This Row],[工序号]]="","",参数设置!$B$12)',
  status:
    'IF(OR(tblProcess[[#This Row],[负荷ST(s)]]="",tblProcess[[#This Row],[TT基准(s)]]="",tblProcess[[#This Row],[TT基准(s)]]=0),"",IF(tblProcess[[#This Row],[负荷ST(s)]]/tblProcess[[#This Row],[TT基准(s)]]>1,"超节拍",IF(tblProcess[[#This Row],[负荷ST(s)]]/tblProcess[[#This Row],[TT基准(s)]]<0.7,"负荷低","正常")))',
  maxCt:
    'IF(COUNT(tblProcess[[#This Row],[测量1]]:tblProcess[[#This Row],[测量5]])=0,"",MAX(tblProcess[[#This Row],[测量1]]:tblProcess[[#This Row],[测量5]]))',
  variation:
    'IF(OR(tblProcess[[#This Row],[最大CT(s)]]="",tblProcess[[#This Row],[TT基准(s)]]="",tblProcess[[#This Row],[TT基准(s)]]=0),"",IF(tblProcess[[#This Row],[最大CT(s)]]/tblProcess[[#This Row],[TT基准(s)]]>1,"高",IF(tblProcess[[#This Row],[最大CT(s)]]/tblProcess[[#This Row],[TT基准(s)]]>=0.95,"中","低")))',
  normalBar:
    'IF(OR(tblProcess[[#This Row],[负荷ST(s)]]="",tblProcess[[#This Row],[TT基准(s)]]=""),0,IF(tblProcess[[#This Row],[负荷ST(s)]]<=tblProcess[[#This Row],[TT基准(s)]],tblProcess[[#This Row],[负荷ST(s)]],0))',
  overBar:
    'IF(OR(tblProcess[[#This Row],[负荷ST(s)]]="",tblProcess[[#This Row],[TT基准(s)]]=""),0,IF(tblProcess[[#This Row],[负荷ST(s)]]>tblProcess[[#This Row],[TT基准(s)]],tblProcess[[#This Row],[负荷ST(s)]],0))'
};

function templateCandidates() {
  const here = path.dirname(fileURLToPath(import.meta.url));
  return [
    path.resolve(process.cwd(), ...TEMPLATE_PARTS),
    path.resolve(here, "../", ...TEMPLATE_PARTS),
    path.resolve(here, ...TEMPLATE_PARTS)
  ];
}

async function loadTemplate() {
  const templatePath = templateCandidates().find((candidate) => existsSync(candidate));
  if (!templatePath) {
    throw new Error("未找到 v16 Excel 模板文件");
  }
  return fs.readFile(templatePath);
}

function escapeXml(value: unknown) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function xmlAttr(source: string, attr: string) {
  return source.match(new RegExp(`${attr}="([^"]*)"`))?.[1] ?? "";
}

function columnName(index: number) {
  let name = "";
  while (index > 0) {
    const mod = (index - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    index = Math.floor((index - mod) / 26);
  }
  return name;
}

function cellXml(ref: string, value: string | number | null | undefined, style = "") {
  const styleAttr = style ? ` s="${style}"` : "";
  if (value === null || value === undefined || value === "") {
    return `<c r="${ref}"${styleAttr}/>`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`;
  }
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function formulaCellXml(ref: string, formula: string, style = "") {
  const styleAttr = style ? ` s="${style}"` : "";
  return `<c r="${ref}"${styleAttr}><f>${escapeXml(formula)}</f></c>`;
}

function styleForCell(sheetXml: string, ref: string) {
  const match = sheetXml.match(new RegExp(`<c r="${ref}"([^>]*)`));
  return match?.[1].match(/s="([^"]*)"/)?.[1] ?? "";
}

function replaceOrAddCell(rowXml: string, ref: string, replacement: string) {
  const cellPattern = new RegExp(`<c r="${ref}"(?: [^>]*)?(?:/>|>[\\s\\S]*?</c>)`);
  if (cellPattern.test(rowXml)) {
    return rowXml.replace(cellPattern, replacement);
  }
  return rowXml.replace("</row>", `${replacement}</row>`);
}

function updateCell(sheetXml: string, row: number, col: string, replacement: string) {
  const rowPattern = new RegExp(`(<row[^>]*r="${row}"[^>]*>)([\\s\\S]*?)(</row>)`);
  const match = sheetXml.match(rowPattern);
  const ref = `${col}${row}`;
  if (!match) {
    const newRow = `<row r="${row}">${replacement}</row>`;
    return sheetXml.replace("</sheetData>", `${newRow}</sheetData>`);
  }
  const nextRow = replaceOrAddCell(match[0], ref, replacement);
  return sheetXml.replace(match[0], nextRow);
}

function rowFromValues(rowIndex: number, values: Array<string | number | null | undefined>) {
  const cells = values.map((value, index) => cellXml(`${columnName(index + 1)}${rowIndex}`, value)).join("");
  return `<row r="${rowIndex}">${cells}</row>`;
}

async function readZipXml(zip: JSZip, part: string) {
  const file = zip.file(part);
  if (!file) throw new Error(`Excel 内部文件缺失：${part}`);
  return file.async("string");
}

function writeZipXml(zip: JSZip, part: string, xml: string) {
  zip.file(part, xml);
}

function normalizeTarget(target: string) {
  const cleaned = target.replace(/^\/+/, "");
  return cleaned.startsWith("xl/") ? cleaned : path.posix.join("xl", cleaned);
}

async function worksheetPath(zip: JSZip, sheetName: string) {
  const workbookXml = await readZipXml(zip, "xl/workbook.xml");
  const relsXml = await readZipXml(zip, "xl/_rels/workbook.xml.rels");
  const sheetMatch = [...workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)].find((match) => xmlAttr(match[1], "name") === sheetName);
  if (!sheetMatch) throw new Error(`找不到工作表：${sheetName}`);
  const relId = xmlAttr(sheetMatch[1], "r:id");
  const relMatch = [...relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)].find((match) => xmlAttr(match[1], "Id") === relId);
  if (!relMatch) throw new Error(`找不到工作表关系：${sheetName}`);
  return normalizeTarget(xmlAttr(relMatch[1], "Target"));
}

function patchDimension(sheetXml: string, ref: string) {
  return sheetXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${ref}"/>`);
}

function patchProcessSheet(originalXml: string, rows: ExportProcess[]) {
  const lastRow = Math.max(rows.length, 1) + 2;
  const row1 = originalXml.match(/<row[^>]*r="1"[\s\S]*?<\/row>/)?.[0] ?? "";
  const row2 = originalXml.match(/<row[^>]*r="2"[\s\S]*?<\/row>/)?.[0] ?? "";
  const styles = Object.fromEntries(
    ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"].map((col) => [col, styleForCell(originalXml, `${col}3`)])
  ) as Record<string, string>;

  const dataRows = Array.from({ length: Math.max(rows.length, 1) }, (_, index) => {
    const rowNumber = index + 3;
    const item = rows[index];
    const process = item?.process;
    const groupName = item?.group.name ?? "";
    const samples = process?.samples ?? [null, null, null, null, null];
    const remarkParts = [groupName, process?.remark].filter(Boolean);
    const cells = [
      formulaCellXml(`B${rowNumber}`, formulas.processNo, styles.B),
      cellXml(`C${rowNumber}`, process?.name ?? "", styles.C),
      cellXml(`D${rowNumber}`, samples[0], styles.D),
      cellXml(`E${rowNumber}`, samples[1], styles.E),
      cellXml(`F${rowNumber}`, samples[2], styles.F),
      cellXml(`G${rowNumber}`, samples[3], styles.G),
      cellXml(`H${rowNumber}`, samples[4], styles.H),
      process?.people && process.people > 0
        ? cellXml(`I${rowNumber}`, process.people, styles.I)
        : formulaCellXml(`I${rowNumber}`, 'IF(OR(tblProcess[[#This Row],[工序名称]]<>"",COUNT(tblProcess[[#This Row],[测量1]]:tblProcess[[#This Row],[测量5]])>0),1,"")', styles.I),
      formulaCellXml(`J${rowNumber}`, formulas.average, styles.J),
      formulaCellXml(`K${rowNumber}`, formulas.standard, styles.K),
      formulaCellXml(`L${rowNumber}`, formulas.load, styles.L),
      formulaCellXml(`M${rowNumber}`, formulas.takt, styles.M),
      formulaCellXml(`N${rowNumber}`, formulas.status, styles.N),
      formulaCellXml(`O${rowNumber}`, formulas.maxCt, styles.O),
      formulaCellXml(`P${rowNumber}`, formulas.variation, styles.P),
      cellXml(`Q${rowNumber}`, remarkParts.join("｜"), styles.Q),
      formulaCellXml(`R${rowNumber}`, formulas.normalBar, styles.R),
      formulaCellXml(`S${rowNumber}`, formulas.overBar, styles.S)
    ];
    return `<row r="${rowNumber}" spans="2:19">${cells.join("")}</row>`;
  }).join("");

  return patchDimension(
    originalXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${row1}${row2}${dataRows}</sheetData>`),
    `B1:S${lastRow}`
  );
}

function patchParameterSheet(sheetXml: string, session: WorktimeSession) {
  const values: Array<[number, number | null]> = [
    [3, session.parameters.plannedOutput],
    [4, session.parameters.shiftHours],
    [5, session.parameters.utilization],
    [6, session.parameters.allowance],
    [7, session.parameters.skill],
    [8, session.parameters.effort],
    [9, session.parameters.environment],
    [10, session.parameters.consistency]
  ];
  return values.reduce((xml, [row, value]) => {
    const style = styleForCell(xml, `B${row}`);
    return updateCell(xml, row, "B", cellXml(`B${row}`, value, style));
  }, sheetXml);
}

async function patchTable(zip: JSZip, lastRow: number) {
  const tablePart = Object.keys(zip.files).find((part) => part.startsWith("xl/tables/") && part.endsWith(".xml"));
  if (!tablePart) throw new Error("找不到 tblProcess 表格定义");
  const xml = await readZipXml(zip, tablePart);
  if (!xml.includes('name="tblProcess"')) return;
  const ref = `B2:S${lastRow}`;
  writeZipXml(
    zip,
    tablePart,
    xml
      .replace(/ref="B2:S\d+"/g, `ref="${ref}"`)
      .replace(/<autoFilter ref="B2:S\d+"\/>/g, `<autoFilter ref="${ref}"/>`)
  );
}

async function patchCharts(zip: JSZip, lastRow: number) {
  const chartParts = Object.keys(zip.files).filter((part) => part.startsWith("xl/charts/") && part.endsWith(".xml"));
  await Promise.all(
    chartParts.map(async (part) => {
      const xml = await readZipXml(zip, part);
      const patched = xml
        .replace(/工序测量!\$C\$3:\$C\$\d+/g, `工序测量!$C$3:$C$${lastRow}`)
        .replace(/工序测量!\$R\$3:\$R\$\d+/g, `工序测量!$R$3:$R$${lastRow}`)
        .replace(/工序测量!\$S\$3:\$S\$\d+/g, `工序测量!$S$3:$S$${lastRow}`)
        .replace(/工序测量!\$M\$3:\$M\$\d+/g, `工序测量!$M$3:$M$${lastRow}`);
      writeZipXml(zip, part, patched);
    })
  );
}

function buildTraceSheetXml(session: WorktimeSession, rows: ExportProcess[], exportTitle: string) {
  const line = findLine(session.lineId);
  const header = ["班组", "工序序号", "工序名称", "测量1", "测量2", "测量3", "测量4", "测量5", "人员", "备注"];
  const dataRows = rows.map((item, index) =>
    rowFromValues(index + 8, [
      item.group.name,
      item.process.stationNo ?? item.process.order,
      item.process.name,
      ...item.process.samples,
      item.process.people,
      item.process.remark
    ])
  );
  const sheetRows = [
    rowFromValues(1, ["现场测量记录"]),
    rowFromValues(3, ["线体", line?.name ?? session.lineId, "机型", displayModelName(session)]),
    rowFromValues(4, ["导出范围", exportTitle, "导出时间", new Date().toISOString()]),
    rowFromValues(5, ["计划单班产量", session.parameters.plannedOutput, "班次工时", session.parameters.shiftHours]),
    rowFromValues(6, ["稼动率", session.parameters.utilization, "宽放率", session.parameters.allowance]),
    rowFromValues(7, header),
    ...dataRows
  ].join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:J${Math.max(8, rows.length + 7)}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="16"/>
  <cols>
    <col min="1" max="1" width="14" customWidth="1"/>
    <col min="2" max="2" width="12" customWidth="1"/>
    <col min="3" max="3" width="34" customWidth="1"/>
    <col min="4" max="9" width="10" customWidth="1"/>
    <col min="10" max="10" width="28" customWidth="1"/>
  </cols>
  <sheetData>${sheetRows}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
}

async function addTraceSheet(zip: JSZip, session: WorktimeSession, rows: ExportProcess[], exportTitle: string) {
  const workbookXml = await readZipXml(zip, "xl/workbook.xml");
  const relsXml = await readZipXml(zip, "xl/_rels/workbook.xml.rels");
  const contentTypesXml = await readZipXml(zip, "[Content_Types].xml");
  const sheetIds = [...workbookXml.matchAll(/sheetId="(\d+)"/g)].map((match) => Number(match[1]));
  const relIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]));
  const sheetTargets = [...relsXml.matchAll(/Target="worksheets\/sheet(\d+)\.xml"/g)].map((match) => Number(match[1]));
  const nextSheetId = Math.max(0, ...sheetIds) + 1;
  const nextRelId = `rId${Math.max(0, ...relIds) + 1}`;
  const nextSheetNumber = Math.max(0, ...sheetTargets) + 1;
  const sheetPart = `xl/worksheets/sheet${nextSheetNumber}.xml`;

  writeZipXml(zip, sheetPart, buildTraceSheetXml(session, rows, exportTitle));
  writeZipXml(
    zip,
    "xl/workbook.xml",
    workbookXml.replace(
      "</sheets>",
      `<sheet name="现场测量记录" sheetId="${nextSheetId}" r:id="${nextRelId}"/></sheets>`
    )
  );
  writeZipXml(
    zip,
    "xl/_rels/workbook.xml.rels",
    relsXml.replace(
      "</Relationships>",
      `<Relationship Id="${nextRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nextSheetNumber}.xml"/></Relationships>`
    )
  );
  writeZipXml(
    zip,
    "[Content_Types].xml",
    contentTypesXml.replace(
      "</Types>",
      `<Override PartName="/${sheetPart}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
    )
  );
}

async function forceWorkbookRecalc(zip: JSZip) {
  const workbookXml = await readZipXml(zip, "xl/workbook.xml");
  const calcPr = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  const patched = workbookXml.includes("<calcPr")
    ? workbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/, calcPr)
    : workbookXml.replace("</workbook>", `${calcPr}</workbook>`);
  writeZipXml(zip, "xl/workbook.xml", patched);
}

function hasProcessContent(process: WorktimeProcess) {
  return Boolean(process.name.trim() || process.remark.trim() || process.samples.some((sample) => sample !== null));
}

function collectRows(session: WorktimeSession, scope: ExportScope, groupId?: string) {
  const selectedGroupId = groupId || session.currentGroupId;
  const groups = scope === "group"
    ? session.groups.filter((group) => group.id === selectedGroupId)
    : [...session.groups].sort((a, b) => a.order - b.order);
  if (scope === "group" && groups.length === 0) {
    throw new Error("当前班组不存在");
  }

  const seen = new Set<string>();
  const rows: ExportProcess[] = [];
  for (const group of groups) {
    const processes = [...group.processes].sort((a, b) => a.order - b.order);
    for (const process of processes) {
      if (!hasProcessContent(process)) continue;
      const key = process.sourceKey || `${group.name}|${process.stationNo ?? ""}|${process.name}`;
      if (scope === "all" && seen.has(key)) continue;
      seen.add(key);
      rows.push({ group, process });
    }
  }
  return rows;
}

function formatTimestamp(date = new Date()) {
  const pad = (value: number) => String(value).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
}

function cleanFileName(value: string) {
  return value.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "");
}

function exportFileName(session: WorktimeSession, scope: ExportScope, exportTitle: string) {
  const line = findLine(session.lineId);
  return cleanFileName(`${line?.name ?? session.lineId}_${displayModelName(session)}_${exportTitle}_工序工时表_${formatTimestamp()}.xlsx`);
}

export async function buildWorkbookBuffer(session: WorktimeSession, scope: ExportScope, groupId?: string): Promise<WorkbookBuildResult> {
  const rows = collectRows(session, scope, groupId);
  const selectedGroup = session.groups.find((group) => group.id === (groupId || session.currentGroupId));
  const exportTitle = scope === "all" ? "整机" : selectedGroup?.name ?? "当前班组";
  const template = await loadTemplate();
  const zip = await JSZip.loadAsync(template);
  const processSheetPart = await worksheetPath(zip, "工序测量");
  const parameterSheetPart = await worksheetPath(zip, "参数设置");
  const processSheetXml = await readZipXml(zip, processSheetPart);
  const parameterSheetXml = await readZipXml(zip, parameterSheetPart);
  const lastRow = Math.max(rows.length, 1) + 2;

  writeZipXml(zip, processSheetPart, patchProcessSheet(processSheetXml, rows));
  writeZipXml(zip, parameterSheetPart, patchParameterSheet(parameterSheetXml, session));
  await patchTable(zip, lastRow);
  await patchCharts(zip, lastRow);
  await addTraceSheet(zip, session, rows, exportTitle);
  await forceWorkbookRecalc(zip);

  const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return {
    buffer,
    fileName: exportFileName(session, scope, exportTitle),
    processCount: rows.length
  };
}
