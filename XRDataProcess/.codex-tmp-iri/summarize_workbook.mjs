import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("I:/动态/甘肃比对最终平整度.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);

for (const sheetName of ["左", "右"]) {
  const sheet = workbook.worksheets.getItem(sheetName);
  const headers = sheet.getRange("A3:G3").values[0];
  const rows = sheet.getRange("A4:P183").values.filter((r) => r.slice(0, 7).every((v) => typeof v === "number"));
  const mean = (arr) => arr.reduce((a, b) => a + b, 0) / arr.length;
  const vehicleMeans = headers.map((_, c) => mean(rows.map((r) => r[c])));
  const meanAbsDevs = headers.map((_, c) => mean(rows.map((r) => r[9 + c])));
  const gtOne = headers.map((_, c) => rows.filter((r) => r[9 + c] > 1).length);
  const consistency = mean(rows.map((r) => r[8]));
  console.log(JSON.stringify({ sheetName, rows: rows.length, headers, vehicleMeans, meanAbsDevs, gtOne, consistency }));
}
