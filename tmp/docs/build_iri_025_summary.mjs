import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const repo = process.cwd();
const sourcePath = path.join(repo, "tmp", "docs", "iri_truth_kb_analysis.json");
const outputDir = path.join(repo, "output", "doc");
const outputPath = path.join(outputDir, "栗庙路_0.25m平整度结果与真值对比.xlsx");
const previewDir = path.join(repo, "tmp", "docs", "xlsx_preview_025");
const data = JSON.parse(await fs.readFile(sourcePath, "utf8"));
const order = ["200-300", "300-400", "400-500", "500-600"];
const samples = [...data.samples].sort((a, b) =>
  a.group.localeCompare(b.group) || a.project.localeCompare(b.project) || order.indexOf(a.interval) - order.indexOf(b.interval)
);

const metrics = (rows) => {
  const errors = rows.map((r) => r.raw_iri - r.truth);
  const rel = rows.map((r, i) => errors[i] / r.truth);
  const mean = (x) => x.reduce((a, b) => a + b, 0) / x.length;
  return {
    count: rows.length,
    iriMean: mean(rows.map((r) => r.raw_iri)),
    bias: mean(errors),
    mae: mean(errors.map(Math.abs)),
    rmse: Math.sqrt(mean(errors.map((x) => x * x))),
    mape: mean(rel.map(Math.abs)),
    maxRel: Math.max(...rel.map(Math.abs)),
    within5: rel.filter((x) => Math.abs(x) <= 0.05).length,
  };
};

const wb = Workbook.create();
const summary = wb.worksheets.add("汇总");
const byInterval = wb.worksheets.add("分区间统计");
const detail = wb.worksheets.add("明细72段");
for (const sheet of [summary, byInterval, detail]) sheet.showGridLines = false;

const titleStyle = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF", size: 16 }, horizontalAlignment: "center", verticalAlignment: "center" };
const sectionStyle = { fill: "#D9EAF7", font: { bold: true, color: "#1F1F1F" }, horizontalAlignment: "center", verticalAlignment: "center" };
const headerStyle = { fill: "#5B9BD5", font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
const tableBorder = { preset: "inside", style: "thin", color: "#D9E2F3" };

// Summary sheet
summary.mergeCells("A1:J1");
summary.getRange("A1").values = [["栗庙路 0.25m 平整度结果与真值对比汇总"]];
summary.getRange("A1:J1").format = titleStyle;
summary.getRange("A1:J1").format.rowHeight = 30;
summary.mergeCells("A2:J2");
summary.getRange("A2").values = [["数据：18个工程、72个100m区间；IRI为当前0.25m重出结果；真值为200–300=1.25、300–400=1.27、400–500=1.43、500–600=1.75。"]];
summary.getRange("A2:J2").format = { fill: "#EAF2F8", font: { color: "#404040", italic: true }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A2:J2").format.rowHeight = 28;
summary.getRange("A4:J4").values = [["速度组", "样本数", "IRI均值", "平均偏差", "MAE", "RMSE", "平均相对误差", "最大相对误差", "≤5%样本数", "结论"]];
summary.getRange("A4:J4").format = headerStyle;
const groups = ["30", "50", "70"];
const summaryRows = groups.map((g) => {
  const m = metrics(samples.filter((s) => s.group === g));
  return [g, m.count, m.iriMean, m.bias, m.mae, m.rmse, m.mape, m.maxRel, `${m.within5}/${m.count}`, m.mape < 0.05 ? "平均误差<5%" : "平均误差≥5%"];
});
const total = metrics(samples);
summaryRows.push(["全部", total.count, total.iriMean, total.bias, total.mae, total.rmse, total.mape, total.maxRel, `${total.within5}/${total.count}`, "全部样本汇总"]);
summary.getRange("A5:J8").values = summaryRows;
summary.getRange("A4:J8").format.borders = tableBorder;
summary.getRange("B5:B8").format.numberFormat = "0";
summary.getRange("C5:F8").format.numberFormat = "0.0000";
summary.getRange("G5:H8").format.numberFormat = "0.00%";
summary.getRange("A8:J8").format = { fill: "#E2F0D9", font: { bold: true } };
summary.getRange("G5:H7").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0.05, format: { fill: "#FCE4D6", font: { color: "#C00000", bold: true } } });
summary.getRange("A10:J10").merge();
summary.getRange("A10").values = [["判读：本工作簿统计的是未带入新K、B的0.25m原始结果与真值的偏差；详细区间与逐工程偏差见后续工作表。K、B拟合可达性请参见对应Markdown报告。"]];
summary.getRange("A10:J10").format = { fill: "#FFF2CC", wrapText: true, verticalAlignment: "center" };
summary.getRange("A10:J10").format.rowHeight = 34;
summary.getRange("L3:M6").values = [["速度组", "平均相对误差"], ...groups.map((g) => {
  const m = metrics(samples.filter((s) => s.group === g));
  return [g, m.mape];
})];
summary.getRange("L3:M6").format.numberFormat = "0.00%";
const chart = summary.charts.add("bar", summary.getRange("L3:M6"));
chart.title = "各速度组平均相对误差";
chart.hasLegend = false;
chart.setPosition("L8", "S23");

// Interval sheet
byInterval.mergeCells("A1:J1");
byInterval.getRange("A1").values = [["0.25m结果：按速度组与真值区间统计"]];
byInterval.getRange("A1:J1").format = titleStyle;
byInterval.getRange("A3:J3").values = [["速度组", "区间(m)", "真值IRI", "IRI均值", "最小IRI", "最大IRI", "平均偏差", "MAE", "平均相对误差", "≤5%样本数"]];
byInterval.getRange("A3:J3").format = headerStyle;
const intervalRows = [];
for (const g of groups) {
  for (const interval of order) {
    const rows = samples.filter((s) => s.group === g && s.interval === interval);
    const m = metrics(rows);
    intervalRows.push([g, interval, rows[0].truth, m.iriMean, Math.min(...rows.map((r) => r.raw_iri)), Math.max(...rows.map((r) => r.raw_iri)), m.bias, m.mae, m.mape, `${m.within5}/${m.count}`]);
  }
}
byInterval.getRange("A4:J15").values = intervalRows;
byInterval.getRange("A3:J15").format.borders = tableBorder;
byInterval.getRange("C4:H15").format.numberFormat = "0.0000";
byInterval.getRange("I4:I15").format.numberFormat = "0.00%";
byInterval.getRange("I4:I15").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0.05, format: { fill: "#FCE4D6", font: { color: "#C00000" } } });
byInterval.freezePanes.freezeRows(3);

// Detail sheet
detail.mergeCells("A1:I1");
detail.getRange("A1").values = [["0.25m重出结果：72个100m区间明细"]];
detail.getRange("A1:I1").format = titleStyle;
detail.getRange("A3:I3").values = [["速度组", "工程", "起始(m)", "结束(m)", "表中左IRI", "平均车速(km/h)", "真值IRI", "偏差", "相对误差"]];
detail.getRange("A3:I3").format = headerStyle;
detail.getRange("A4:G75").values = samples.map((s) => [s.group, s.project.split("__", 1)[0], Number(s.interval.split("-")[0]), Number(s.interval.split("-")[1]), s.raw_iri, s.speed, s.truth]);
detail.getRange("H4").formulas = [["=E4-G4"]];
detail.getRange("H4:H75").fillDown();
detail.getRange("I4").formulas = [["=H4/G4"]];
detail.getRange("I4:I75").fillDown();
detail.getRange("A3:I75").format.borders = tableBorder;
detail.getRange("C4:D75").format.numberFormat = "0";
detail.getRange("E4:H75").format.numberFormat = "0.00000";
detail.getRange("I4:I75").format.numberFormat = "0.00%";
detail.getRange("I4:I75").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0.05, format: { fill: "#FCE4D6", font: { color: "#C00000" } } });
detail.getRange("I4:I75").conditionalFormats.add("cellIs", { operator: "lessThan", formula: -0.05, format: { fill: "#FFF2CC", font: { color: "#9C5700" } } });
detail.tables.add("A3:I75", true, "Iri025Detail");
detail.freezePanes.freezeRows(3);

for (const sheet of [summary, byInterval, detail]) {
  sheet.getUsedRange().format.autofitColumns();
  sheet.getUsedRange().format.autofitRows();
}
summary.getRange("A1:J10").format.columnWidth = 14;
summary.getRange("J4:J8").format.columnWidth = 18;
byInterval.getRange("A1:J15").format.columnWidth = 14;
detail.getRange("A1:I75").format.columnWidth = 14;
detail.getRange("B4:B75").format.columnWidth = 12;

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });
for (const [name, range] of [["汇总", "A1:J10"], ["分区间统计", "A1:J15"], ["明细72段", "A1:I20"]]) {
  const image = await wb.render({ sheetName: name, range, scale: 1.5, format: "png" });
  await fs.writeFile(path.join(previewDir, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
}
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);
console.log(JSON.stringify({ outputPath, previewDir }));
