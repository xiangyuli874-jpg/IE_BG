import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
const result = await workbook.help("chart");
console.log(result.ndjson ?? JSON.stringify(result, null, 2));
