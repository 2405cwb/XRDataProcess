import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("I:/动态/甘肃比对最终平整度.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);

for (const sheetName of ["左", "右"]) {
  const sheet = workbook.worksheets.getItem(sheetName);
  for (const rangeName of ["A3:W8", "A175:W187"]) {
    const range = sheet.getRange(rangeName);
    console.log(JSON.stringify({
      sheet: sheetName,
      range: rangeName,
      values: range.values,
      formulas: range.formulas,
    }));
  }
}
