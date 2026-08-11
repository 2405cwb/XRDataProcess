import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "I:/动态/甘肃比对最终平整度.xlsx";
const outputDir = "G:/job/工作COD/2025/公路/XRDataProcess/XRDataProcess/.codex-tmp-iri/rendered";
await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(source);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheets = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 12000 });
console.log("SHEETS");
console.log(sheets.ndjson);

const overview = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 30000,
  tableMaxRows: 15,
  tableMaxCols: 20,
  tableMaxCellChars: 120,
});
console.log("OVERVIEW");
console.log(overview.ndjson);

const formulas = await workbook.inspect({
  kind: "formula",
  maxChars: 30000,
  options: { maxResults: 500 },
});
console.log("FORMULAS");
console.log(formulas.ndjson);

for (let i = 0; i < workbook.worksheets.items.length; i++) {
  const sheet = workbook.worksheets.getItemAt(i);
  const used = sheet.getUsedRange();
  console.log(`USED ${sheet.name}`, JSON.stringify({ address: used?.address, values: used?.values?.slice(0, 12) }));
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 0.7, format: "png" });
  const safe = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(`${outputDir}/${String(i + 1).padStart(2, "0")}_${safe}.png`, new Uint8Array(await preview.arrayBuffer()));
}
