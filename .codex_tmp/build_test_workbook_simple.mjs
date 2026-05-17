import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "E:/AI/gongshibiao/工时测量表+员工负荷山积表.xlsx";
const outputDir = "E:/AI/gongshibiao/outputs/worktime_logic_test";
const outputPath = path.join(outputDir, "工时测量表_山积逻辑测试版.xlsx");

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheetName = "测试版_山积逻辑";

const prior = workbook.worksheets.items.find((sheet) => sheet.name === sheetName);
if (prior) workbook.worksheets.delete(prior);

const sheet = workbook.worksheets.add(sheetName);

sheet.getRange("A1:I1").merge();
sheet.getRange("A1").values = [["测试版：山积图与线平衡逻辑拆解"]];
sheet.getRange("A3:I4").values = [
  ["用途", "这个测试页不改原表，只把最终山积图真正依赖的数据抽出来，便于看清：标准工时表 -> 图表辅助区 -> 山积图。", null, null, null, null, null, null, null],
  ["说明", "原图当前依赖的 80 个工站槽位来自《平衡率及动作变异分析图》B:CC 的横向辅助区；这里先抽取前 20 个槽位做验证。", null, null, null, null, null, null, null],
];
sheet.getRange("B3:I3").merge();
sheet.getRange("B4:I4").merge();

sheet.getRange("A6:I6").values = [[
  "工站槽位",
  "工站",
  "人员",
  "工站总ST(S)",
  "TT(S)",
  "TAV(S)",
  "负荷率",
  "状态",
  "备注",
]];

const fixedValues = [];
const formulas = [];
for (let i = 1; i <= 20; i += 1) {
  const row = 6 + i;
  fixedValues.push([i, null, null, null, null, null, null, null, ""]);
  formulas.push([
    `=IFERROR(INDEX('平衡率及动作变异分析图'!$B$29:$CC$29,1,A${row}),"")`,
    `=IFERROR(INDEX('平衡率及动作变异分析图'!$B$30:$CC$30,1,A${row}),"")`,
    `=IFERROR(INDEX('平衡率及动作变异分析图'!$B$39:$CC$39,1,A${row}),"")`,
    `=IFERROR(INDEX('平衡率及动作变异分析图'!$B$41:$CC$41,1,A${row}),"")`,
    `=IFERROR(INDEX('平衡率及动作变异分析图'!$B$42:$CC$42,1,A${row}),"")`,
    `=IFERROR(D${row}/E${row},"")`,
    `=IF(G${row}="","",IF(G${row}>1,"超节拍",IF(G${row}>=0.9,"接近上限","正常")))`,
  ]);
}
sheet.getRange("A7:I26").values = fixedValues;
sheet.getRange("B7:H26").formulas = formulas;

sheet.getRange("K2:N2").values = [["关键指标", "数值", "来源", "备注"]];
sheet.getRange("K3:N6").values = [
  ["节拍时间TT", null, "CF3", "来自平衡图辅助区"],
  ["平均ST", null, "CF4", "来自工站ST平均值"],
  ["瓶颈ST", null, "CF5", "来自工站ST最大值"],
  ["线平衡率", null, "CF7", "来自工站合计 / 瓶颈产能"],
];
sheet.getRange("L3:L6").formulas = [
  [`=IFERROR('平衡率及动作变异分析图'!CF3,"")`],
  [`=IFERROR('平衡率及动作变异分析图'!CF4,"")`],
  [`=IFERROR('平衡率及动作变异分析图'!CF5,"")`],
  [`=IFERROR('平衡率及动作变异分析图'!CF7,"")`],
];

sheet.getRange("K9:N9").values = [["工站槽位", "ST", "TT", "TAV"]];
const helperFormulas = [];
for (let i = 1; i <= 20; i += 1) {
  const sourceRow = 6 + i;
  helperFormulas.push([
    `=A${sourceRow}`,
    `=D${sourceRow}`,
    `=E${sourceRow}`,
    `=F${sourceRow}`,
  ]);
}
sheet.getRange("K10:N29").formulas = helperFormulas;

sheet.getRange("P2:U5").values = [
  ["字段", "含义", "原始落点", "测试页落点", "继续扩展时建议", null],
  ["工站总ST", "单工站动作ST合计", "图表页 B39:CC39", "D列", "正式模板可转成标准纵表", null],
  ["TT", "目标节拍", "图表页 B41:CC41", "E列", "可改成参数输入驱动", null],
  ["TAV", "平均参考线", "图表页 B42:CC42", "F列", "可保留为比较基线", null],
];
sheet.getRange("P8:U11").merge();
sheet.getRange("P8").values = [[
  "这里预留为下一版测试图区域。当前已经在 K9:N29 做好标准化图表源数据，适合继续做 ST / TT / TAV 对照图。",
]];

sheet.getRange("D7:F26").numberFormat = "0.00";
sheet.getRange("G7:G26").numberFormat = "0.0%";
sheet.getRange("L3:L5").numberFormat = "0.00";
sheet.getRange("L6").numberFormat = "0.0%";
sheet.getRange("L10:N29").numberFormat = "0.00";

sheet.getRange("A:A").columnWidth = 12;
sheet.getRange("B:B").columnWidth = 16;
sheet.getRange("C:C").columnWidth = 10;
sheet.getRange("D:F").columnWidth = 14;
sheet.getRange("G:H").columnWidth = 12;
sheet.getRange("I:I").columnWidth = 16;
sheet.getRange("K:N").columnWidth = 14;
sheet.getRange("P:U").columnWidth = 18;

const check = await workbook.inspect({
  kind: "table",
  range: `${sheetName}!A1:N29`,
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 20,
});
console.log(check.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
