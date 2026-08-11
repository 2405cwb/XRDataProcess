import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const input = JSON.parse(await fs.readFile("./tmp/docs/kb_output_compare_data.json", "utf8"));
const outPath = "C:\\Users\\cwb\\Desktop\\job\\01二维公路软件\\平整度验证\\新算法平整度栗庙路真值验证\\两种K_B出表算法平整度对比.xlsx";
const previewDir = "./tmp/docs/xlsx_preview_kb_output_compare";
await fs.mkdir(previewDir, { recursive: true });

const wb = Workbook.create();
const summary = wb.worksheets.add("对比汇总");
const detail = wb.worksheets.add("逐区间明细");
const cross = wb.worksheets.add("速度_区间交叉汇总");
for (const s of [summary, detail, cross]) s.showGridLines = false;

const navy = "#17365D";
const blue = "#1F4E78";
const lightBlue = "#D9EAF7";
const teal = "#0F766E";
const green = "#E2F0D9";
const red = "#FCE4D6";
const gray = "#F2F2F2";
const border = { preset: "all", style: "thin", color: "#D9E2F3" };
const headerFmt = { fill: blue, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: border };
const sectionFmt = { fill: lightBlue, font: { bold: true, color: navy }, verticalAlignment: "center", borders: border };

function title(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range.split(":")[0]).values = [[text]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 16 }, horizontalAlignment: "center", verticalAlignment: "center" };
}
function writeStatTable(sheet, startRow, titleText, rows, firstHeader, firstKey) {
  sheet.getRange(`A${startRow}:K${startRow}`).merge();
  sheet.getRange(`A${startRow}`).values = [[titleText]];
  sheet.getRange(`A${startRow}:K${startRow}`).format = sectionFmt;
  const headers = [firstHeader, "区间数", "A 平均IRI", "B 平均IRI", "平均差值(B-A)", "平均绝对差", "平均相对变化", "平均绝对相对变化", "最大绝对差", "最大绝对相对变化", "B高/B低"];
  sheet.getRange(`A${startRow + 1}:K${startRow + 1}`).values = [headers];
  sheet.getRange(`A${startRow + 1}:K${startRow + 1}`).format = headerFmt;
  const data = rows.map(x => [
    x[firstKey] ?? x.label, x.count, x.meanA, x.meanB, x.meanDelta, x.meanAbsDelta,
    x.meanRelative, x.meanAbsRelative, x.maxAbsDelta, x.maxAbsRelative, `${x.higher}/${x.lower}`,
  ]);
  const endRow = startRow + 1 + data.length;
  sheet.getRange(`A${startRow + 2}:K${endRow}`).values = data;
  sheet.getRange(`A${startRow + 2}:K${endRow}`).format = { borders: border, verticalAlignment: "center" };
  sheet.getRange(`B${startRow + 2}:B${endRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`C${startRow + 2}:I${endRow}`).format.numberFormat = "0.0000";
  sheet.getRange(`G${startRow + 2}:H${endRow}`).format.numberFormat = "0.00%";
  sheet.getRange(`J${startRow + 2}:J${endRow}`).format.numberFormat = "0.00%";
  sheet.getRange(`E${startRow + 2}:E${endRow}`).conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0, format: { fill: green, font: { color: "#006100" } } });
  sheet.getRange(`E${startRow + 2}:E${endRow}`).conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0, format: { fill: red, font: { color: "#9C0006" } } });
}

title(summary, "A1:K1", "两种 K/B 出表算法平整度对比");
summary.getRange("A2:K2").merge();
summary.getRange("A2").values = [["A：v2.2.5.9 软件出表（乘 K 平均值）   |   B：内部软件（乘 K+B）   |   统计对象：两侧均存在的左侧 IRI 百米区间"]];
summary.getRange("A2:K2").format = { fill: gray, font: { italic: true, color: "#404040" }, horizontalAlignment: "left" };
summary.getRange("A4:K4").merge();
summary.getRange("A4").values = [["总体核对"]];
summary.getRange("A4:K4").format = sectionFmt;
summary.getRange("A5:D5").values = [["匹配工程报表", "匹配百米区间", "A 源数据行", "B 源数据行"]];
summary.getRange("A5:D5").format = headerFmt;
summary.getRange("A6:D6").values = [[18, input.matched, input.rowCountA, input.rowCountB]];
summary.getRange("A6:D6").format = { fill: "#FFFFFF", font: { bold: true, size: 13, color: navy }, horizontalAlignment: "center", borders: border, numberFormat: "#,##0" };
summary.getRange("F5:K5").values = [["A 平均IRI", "B 平均IRI", "平均差值(B-A)", "平均绝对差", "平均相对变化", "最大绝对相对变化"]];
summary.getRange("F5:K5").format = headerFmt;
const o = input.overall;
summary.getRange("F6:K6").values = [[o.meanA, o.meanB, o.meanDelta, o.meanAbsDelta, o.meanRelative, o.maxAbsRelative]];
summary.getRange("F6:K6").format = { fill: "#FFFFFF", font: { bold: true, size: 13, color: navy }, horizontalAlignment: "center", borders: border };
summary.getRange("F6:I6").format.numberFormat = "0.0000";
summary.getRange("J6:K6").format.numberFormat = "0.00%";

writeStatTable(summary, 9, "按速度汇总（全部匹配区间，含 600m 后末段）", input.bySpeed, "速度", "label");
writeStatTable(summary, 16, "稳定区间按速度汇总（200-600m）", input.stableBySpeed, "速度", "label");
writeStatTable(summary, 23, "按百米区间汇总（全部速度，末段单列）", input.bySegment, "百米区间", "label");
summary.getRange("A33:K33").merge();
summary.getRange("A33").values = [["说明：差值 = B − A；相对变化 = (B − A) ÷ A。正值表示“乘 K+B”结果高于“乘 K 平均值”；全部匹配区间包含 600m 后各工程的末段。"]];
summary.getRange("A33:K33").format = { fill: gray, font: { italic: true, color: "#595959" }, wrapText: true };
summary.getRange("A1:K33").format.font = { name: "Microsoft YaHei" };
for (const [col, width] of Object.entries({ A: 17, B: 10, C: 14, D: 14, E: 17, F: 15, G: 16, H: 20, I: 15, J: 20, K: 14 })) summary.getRange(`${col}:${col}`).format.columnWidth = width;
summary.getRange("1:1").format.rowHeight = 26;
summary.getRange("2:2").format.rowHeight = 22;
summary.getRange("33:33").format.rowHeight = 30;
summary.freezePanes.freezeRows(2);

title(detail, "A1:K1", "逐百米区间明细对比");
detail.getRange("A2:K2").merge();
detail.getRange("A2").values = [["A：乘 K 平均值；B：乘 K+B。差值与相对变化使用工作簿公式，筛选器可按速度、工程、里程区间定位。"]];
detail.getRange("A2:K2").format = { fill: gray, font: { italic: true, color: "#404040" } };
const headers = ["速度分组(km/h)", "工程", "起始桩号(m)", "结束桩号(m)", "百米区间", "报表车速(km/h)", "A 左IRI(m/km)", "B 左IRI(m/km)", "差值 B-A(m/km)", "相对变化", "方向"];
detail.getRange("A4:K4").values = [headers];
detail.getRange("A4:K4").format = headerFmt;
const rowValues = input.detail.map(x => [x.speedGroup, x.project, x.start, x.end, x.interval, x.speed, x.iriA, x.iriB, null, null, null]);
const detailEnd = 4 + rowValues.length;
detail.getRange(`A5:H${detailEnd}`).values = rowValues.map(r => r.slice(0, 8));
detail.getRange(`A5:H${detailEnd}`).format = { borders: border, verticalAlignment: "center" };
detail.getRange(`I5`).formulas = [["=H5-G5"]];
detail.getRange(`I5:I${detailEnd}`).fillDown();
detail.getRange(`J5`).formulas = [["=IF(G5=0,\"\",I5/G5)"]];
detail.getRange(`J5:J${detailEnd}`).fillDown();
detail.getRange(`K5`).formulas = [["=IF(I5>0,\"B高\",IF(I5<0,\"B低\",\"相同\"))"]];
detail.getRange(`K5:K${detailEnd}`).fillDown();
detail.getRange(`I5:K${detailEnd}`).format = { borders: border, verticalAlignment: "center" };
detail.getRange(`A5:F${detailEnd}`).format.numberFormat = "#,##0";
detail.getRange(`G5:I${detailEnd}`).format.numberFormat = "0.0000";
detail.getRange(`J5:J${detailEnd}`).format.numberFormat = "0.00%";
detail.getRange(`I5:I${detailEnd}`).conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0, format: { fill: green, font: { color: "#006100" } } });
detail.getRange(`I5:I${detailEnd}`).conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0, format: { fill: red, font: { color: "#9C0006" } } });
detail.getRange(`J5:J${detailEnd}`).conditionalFormats.add("dataBar", { color: "#5B9BD5", gradient: true });
const tbl = detail.tables.add(`A4:K${detailEnd}`, true, "K_B_Compare_Detail");
tbl.style = "TableStyleMedium2";
tbl.showBandedColumns = false;
detail.freezePanes.freezeRows(4);
detail.freezePanes.freezeColumns(2);
for (const [col, width] of Object.entries({ A: 14, B: 13, C: 13, D: 13, E: 15, F: 15, G: 16, H: 16, I: 17, J: 13, K: 11 })) detail.getRange(`${col}:${col}`).format.columnWidth = width;
detail.getRange("1:1").format.rowHeight = 26;
detail.getRange("2:2").format.rowHeight = 22;

title(cross, "A1:K1", "速度 × 百米区间交叉汇总");
cross.getRange("A2:K2").merge();
cross.getRange("A2").values = [["用于判断差异是否集中在特定速度或路段；差值 = B − A。"]];
cross.getRange("A2:K2").format = { fill: gray, font: { italic: true, color: "#404040" } };
const crossHeaders = ["速度(km/h)", "百米区间", "区间数", "A 平均IRI", "B 平均IRI", "平均差值", "平均绝对差", "平均相对变化", "平均绝对相对变化", "最大绝对差", "最大绝对相对变化"];
cross.getRange("A4:K4").values = [crossHeaders];
cross.getRange("A4:K4").format = headerFmt;
const cvals = input.bySpeedSegment.map(x => [x.speed, x.interval, x.count, x.meanA, x.meanB, x.meanDelta, x.meanAbsDelta, x.meanRelative, x.meanAbsRelative, x.maxAbsDelta, x.maxAbsRelative]);
cross.getRange(`A5:K${4 + cvals.length}`).values = cvals;
cross.getRange(`A5:K${4 + cvals.length}`).format = { borders: border, verticalAlignment: "center" };
cross.getRange(`A5:C${4 + cvals.length}`).format.numberFormat = "#,##0";
cross.getRange(`D5:G${4 + cvals.length}`).format.numberFormat = "0.0000";
cross.getRange(`H5:I${4 + cvals.length}`).format.numberFormat = "0.00%";
cross.getRange(`J5:J${4 + cvals.length}`).format.numberFormat = "0.0000";
cross.getRange(`K5:K${4 + cvals.length}`).format.numberFormat = "0.00%";
cross.getRange(`F5:F${4 + cvals.length}`).conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0, format: { fill: green, font: { color: "#006100" } } });
cross.getRange(`F5:F${4 + cvals.length}`).conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0, format: { fill: red, font: { color: "#9C0006" } } });
const crossTable = cross.tables.add(`A4:K${4 + cvals.length}`, true, "K_B_Cross_Summary");
crossTable.style = "TableStyleMedium2";
cross.freezePanes.freezeRows(4);
for (const [col, width] of Object.entries({ A: 14, B: 16, C: 10, D: 14, E: 14, F: 14, G: 15, H: 16, I: 20, J: 15, K: 20 })) cross.getRange(`${col}:${col}`).format.columnWidth = width;
cross.getRange("1:1").format.rowHeight = 26;

const check1 = await wb.inspect({ kind: "table", range: "对比汇总!A1:K33", include: "values,formulas", tableMaxRows: 33, tableMaxCols: 11 });
console.log(check1.ndjson);
const check2 = await wb.inspect({ kind: "table", range: "逐区间明细!A1:K12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 11 });
console.log(check2.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);

for (const [sheetName, range, file] of [["对比汇总", "A1:K33", "summary.png"], ["逐区间明细", "A1:K22", "detail.png"], ["速度_区间交叉汇总", "A1:K25", "cross.png"]]) {
  const img = await wb.render({ sheetName, range, scale: 1.25, format: "png" });
  await fs.writeFile(`${previewDir}/${file}`, new Uint8Array(await img.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outPath);
console.log(`SAVED ${outPath}`);
