import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const repo = process.cwd();
const source = JSON.parse(await fs.readFile(path.join(repo, "tmp", "docs", "iri_truth_kb_analysis.json"), "utf8"));
const outDir = path.join(repo, "output", "doc");
const outFile = path.join(outDir, "栗庙路_0.25m原始数据_KB两种优化方案.xlsx");
const previewDir = path.join(repo, "tmp", "docs", "xlsx_preview_kb_optimization");
const order = ["200-300", "300-400", "400-500", "500-600"];
const groups = ["30", "50", "70"];
const samples = [...source.samples].sort((a, b) => a.group.localeCompare(b.group) || a.project.localeCompare(b.project) || order.indexOf(a.interval) - order.indexOf(b.interval));

function minMape(rows) {
  const x = rows.map((r) => r.raw_iri), y = rows.map((r) => r.truth), candidates = [[0, 1.35]];
  for (let i = 0; i < x.length; i++) for (let j = i + 1; j < x.length; j++) {
    const d = x[i] - x[j];
    if (Math.abs(d) > 1e-12) { const k = (y[i] - y[j]) / d; if (k >= 0) candidates.push([k, y[i] - k * x[i]]); }
  }
  const score = ([k, b]) => y.reduce((sum, yy, i) => sum + Math.abs((k * x[i] + b - yy) / yy), 0) / y.length;
  const [k, b] = candidates.reduce((best, q) => score(q) < score(best) ? q : best);
  const rel = y.map((yy, i) => (k * x[i] + b - yy) / yy);
  return { k, b, mean: score([k, b]), max: Math.max(...rel.map(Math.abs)) };
}

function minimax(rows) {
  const x = rows.map((r) => r.raw_iri), y = rows.map((r) => r.truth), candidates = [];
  for (let i = 0; i < x.length; i++) for (let j = i + 1; j < x.length; j++) for (let h = j + 1; h < x.length; h++) {
    for (const si of [-1, 1]) for (const sj of [-1, 1]) for (const sh of [-1, 1]) {
      const a = [[x[i], 1, -si*y[i]], [x[j], 1, -sj*y[j]], [x[h], 1, -sh*y[h]]];
      const det = a[0][0]*(a[1][1]*a[2][2]-a[1][2]*a[2][1])-a[0][1]*(a[1][0]*a[2][2]-a[1][2]*a[2][0])+a[0][2]*(a[1][0]*a[2][1]-a[1][1]*a[2][0]);
      if (Math.abs(det) < 1e-12) continue;
      const solve = (m, v) => { const d=m[0][0]*(m[1][1]*m[2][2]-m[1][2]*m[2][1])-m[0][1]*(m[1][0]*m[2][2]-m[1][2]*m[2][0])+m[0][2]*(m[1][0]*m[2][1]-m[1][1]*m[2][0]); const dx=v[0]*(m[1][1]*m[2][2]-m[1][2]*m[2][1])-m[0][1]*(v[1]*m[2][2]-m[1][2]*v[2])+m[0][2]*(v[1]*m[2][1]-m[1][1]*v[2]); const dy=m[0][0]*(v[1]*m[2][2]-m[1][2]*v[2])-v[0]*(m[1][0]*m[2][2]-m[1][2]*m[2][0])+m[0][2]*(m[1][0]*v[2]-v[1]*m[2][0]); const dz=m[0][0]*(m[1][1]*v[2]-v[1]*m[2][1])-m[0][1]*(m[1][0]*v[2]-v[1]*m[2][0])+v[0]*(m[1][0]*m[2][1]-m[1][1]*m[2][0]); return [dx/d,dy/d,dz/d]; };
      const q = solve(a, [y[i], y[j], y[h]]); if (q[0] >= 0 && q[2] >= -1e-10) candidates.push(q);
    }
  }
  const score = ([k,b]) => Math.max(...y.map((yy,i) => Math.abs((k*x[i]+b-yy)/yy)));
  const q = candidates.reduce((best, q) => score(q) < score(best) ? q : best);
  const rel = y.map((yy,i) => (q[0]*x[i]+q[1]-yy)/yy);
  return { k:q[0], b:q[1], mean:rel.reduce((s,v)=>s+Math.abs(v),0)/rel.length, max:Math.max(...rel.map(Math.abs)) };
}

const models = Object.fromEntries(groups.map((g) => { const rows = samples.filter((s) => s.group === g); return [g, { mape:minMape(rows), minimax:minimax(rows), speedMin:Math.min(...rows.map(r=>r.speed)), speedMax:Math.max(...rows.map(r=>r.speed)) }]; }));

const wb = Workbook.create();
const overview = wb.worksheets.add("方案汇总");
const params = wb.worksheets.add("K_B参数与建议");
const detail = wb.worksheets.add("72段预测明细");
for (const sh of [overview, params, detail]) sh.showGridLines = false;
const title = { fill:"#1F4E78", font:{bold:true,color:"#FFFFFF",size:16}, horizontalAlignment:"center", verticalAlignment:"center" };
const header = { fill:"#5B9BD5", font:{bold:true,color:"#FFFFFF"}, horizontalAlignment:"center", verticalAlignment:"center", wrapText:true };
const border = { preset:"inside", style:"thin", color:"#D9E2F3" };

// Overview
overview.mergeCells("A1:L1"); overview.getRange("A1").values=[["栗庙路0.25m原始平整度：K、B两种优化方案"]]; overview.getRange("A1:L1").format=title; overview.getRange("A1:L1").format.rowHeight=30;
overview.mergeCells("A2:L2"); overview.getRange("A2").values=[["数据来源：2026-07-30 14:30重出的“原始平整度结果”18份表；每速度组24条、合计72条100m数据。真值：1.25、1.27、1.43、1.75。"]]; overview.getRange("A2:L2").format={fill:"#EAF2F8",wrapText:true,font:{italic:true,color:"#404040"}};
overview.getRange("A4:L4").values=[["速度组","实际速度范围","最小平均绝对误差K","最小平均绝对误差B","该方案平均绝对误差","该方案最大绝对误差","最小最大绝对误差K","最小最大绝对误差B","该方案平均绝对误差","该方案最大绝对误差","建议","5%绝对误差结论"]]; overview.getRange("A4:L4").format=header; overview.getRange("A4:L4").format.rowHeight=40;
const overviewRows=groups.map((g)=>{ const m=models[g]; const advice=g==="70"?"平均绝对误差优先：用最小平均方案":"单一K/B无法达到平均绝对误差5%，不建议固化"; return [g,`${m.speedMin.toFixed(2)}–${m.speedMax.toFixed(2)}`,m.mape.k,m.mape.b,m.mape.mean,m.mape.max,m.minimax.k,m.minimax.b,m.minimax.mean,m.minimax.max,advice,g==="70"?"平均绝对值可<5%；单段不可":"平均、单段绝对值均不可<5%"]});
overview.getRange("A5:L7").values=overviewRows; overview.getRange("A4:L7").format.borders=border; overview.getRange("C5:J7").format.numberFormat="0.000000"; overview.getRange("E5:F7").format.numberFormat="0.00%"; overview.getRange("I5:J7").format.numberFormat="0.00%"; overview.getRange("K5:L7").format.wrapText=true;
overview.getRange("E5:F7").conditionalFormats.add("cellIs",{operator:"greaterThan",formula:0.05,format:{fill:"#FCE4D6",font:{color:"#C00000",bold:true}}}); overview.getRange("I5:J7").conditionalFormats.add("cellIs",{operator:"greaterThan",formula:0.05,format:{fill:"#FCE4D6",font:{color:"#C00000",bold:true}}});
overview.mergeCells("A9:L10"); overview.getRange("A9").values=[["绝对误差定义：ABS(IRI预测−IRI真值)/IRI真值。若考核平均绝对误差，采用“最小平均绝对误差”列的K、B；30、50速度的理论下限分别为6.04%、5.09%，不能承诺小于5%。若要求每个100m绝对误差都≤5%，采用“最小最大绝对误差”列；但三组最优最大绝对误差仍为11.74%、11.20%、8.24%，因此均不可达。"]]; overview.getRange("A9:L10").format={fill:"#FFF2CC",wrapText:true,verticalAlignment:"center"}; overview.getRange("A9:L10").format.rowHeight=28;
overview.getRange("N3:O6").values=[["速度组","最小平均绝对误差"],...groups.map(g=>[g,models[g].mape.mean])]; overview.getRange("N3:O6").format.numberFormat="0.00%"; const chart=overview.charts.add("bar",overview.getRange("N3:O6")); chart.title="最小可达平均绝对相对误差"; chart.hasLegend=false; chart.setPosition("N8","U23");

// Parameters and guidance
params.mergeCells("A1:F1"); params.getRange("A1").values=[["K、B参数及使用建议"]]; params.getRange("A1:F1").format=title;
params.getRange("A3:F3").values=[["目标","速度组","K","B","平均绝对相对误差","最大绝对相对误差"]]; params.getRange("A3:F3").format=header;
const pr=[]; for(const g of groups){const m=models[g].mape;pr.push(["最小平均绝对误差",g,m.k,m.b,m.mean,m.max]);} for(const g of groups){const m=models[g].minimax;pr.push(["最小最大绝对误差",g,m.k,m.b,m.mean,m.max]);}
params.getRange("A4:F9").values=pr; params.getRange("A3:F9").format.borders=border; params.getRange("C4:D9").format.numberFormat="0.000000"; params.getRange("E4:F9").format.numberFormat="0.00%";
params.mergeCells("A11:F14"); params.getRange("A11").values=[["建议：\n1）70速度：若只考核平均绝对相对误差，推荐最小平均方案 K=1.286898、B=-0.413919，平均绝对误差3.51%；但最大绝对误差仍为8.80%。\n2）50速度：最小平均方案 K=0.834415、B=0.281737，平均绝对误差理论最低仍为5.09%，不可承诺达标。\n3）30速度：最小平均方案 K=0.940095、B=0.277096，平均绝对误差理论最低6.04%，不建议以单一K、B作为最终方案。"]]; params.getRange("A11:F14").format={fill:"#FFF2CC",wrapText:true,verticalAlignment:"top"}; params.getRange("A11:F14").format.rowHeight=35;

// Detail with formulas linked to parameter cells
detail.mergeCells("A1:M1"); detail.getRange("A1").values=[["0.25m原始IRI：两种K、B方案逐段预测与偏差"]]; detail.getRange("A1:M1").format=title;
detail.getRange("A3:M3").values=[["速度组","工程","区间(m)","原始IRI","车速(km/h)","真值IRI","最小平均方案预测","偏差","绝对相对误差","最小最大方案预测","偏差","绝对相对误差","备注"]]; detail.getRange("A3:M3").format=header;
detail.getRange("A4:F75").values=samples.map(s=>[s.group,s.project.split("__",1)[0],s.interval,s.raw_iri,s.speed,s.truth]);
const mapeFormula="=D4*IF(A4=\"30\",'K_B参数与建议'!$C$4,IF(A4=\"50\",'K_B参数与建议'!$C$5,'K_B参数与建议'!$C$6))+IF(A4=\"30\",'K_B参数与建议'!$D$4,IF(A4=\"50\",'K_B参数与建议'!$D$5,'K_B参数与建议'!$D$6))";
const maxFormula="=D4*IF(A4=\"30\",'K_B参数与建议'!$C$7,IF(A4=\"50\",'K_B参数与建议'!$C$8,'K_B参数与建议'!$C$9))+IF(A4=\"30\",'K_B参数与建议'!$D$7,IF(A4=\"50\",'K_B参数与建议'!$D$8,'K_B参数与建议'!$D$9))";
detail.getRange("G4").formulas=[[mapeFormula]]; detail.getRange("G4:G75").fillDown(); detail.getRange("H4").formulas=[["=G4-F4"]]; detail.getRange("H4:H75").fillDown(); detail.getRange("I4").formulas=[["=ABS(H4/F4)"]]; detail.getRange("I4:I75").fillDown();
detail.getRange("J4").formulas=[[maxFormula]]; detail.getRange("J4:J75").fillDown(); detail.getRange("K4").formulas=[["=J4-F4"]]; detail.getRange("K4:K75").fillDown(); detail.getRange("L4").formulas=[["=ABS(K4/F4)"]]; detail.getRange("L4:L75").fillDown(); detail.getRange("M4").formulas=[["=IF(I4<=5%,\"平均方案绝对误差≤5%\",\"平均方案绝对误差>5%\")"]]; detail.getRange("M4:M75").fillDown();
detail.getRange("A3:M75").format.borders=border; detail.getRange("D4:H75").format.numberFormat="0.000000"; detail.getRange("I4:I75").format.numberFormat="0.00%"; detail.getRange("J4:K75").format.numberFormat="0.000000"; detail.getRange("L4:L75").format.numberFormat="0.00%";
detail.getRange("I4:I75").conditionalFormats.add("cellIs",{operator:"greaterThan",formula:0.05,format:{fill:"#FCE4D6",font:{color:"#C00000"}}}); detail.getRange("L4:L75").conditionalFormats.add("cellIs",{operator:"greaterThan",formula:0.05,format:{fill:"#FCE4D6",font:{color:"#C00000"}}}); detail.tables.add("A3:M75",true,"KbOptimizationDetail"); detail.freezePanes.freezeRows(3);

for(const sh of [overview,params,detail]){sh.getUsedRange().format.autofitColumns();sh.getUsedRange().format.autofitRows();}
overview.getRange("A1:L10").format.columnWidth=14; overview.getRange("K4:L7").format.columnWidth=34; overview.getRange("K5:L7").format.rowHeight=42; params.getRange("A1:F14").format.columnWidth=18; detail.getRange("A1:M75").format.columnWidth=14; detail.getRange("B4:B75").format.columnWidth=12; detail.getRange("M4:M75").format.columnWidth=18;

await fs.mkdir(outDir,{recursive:true}); await fs.mkdir(previewDir,{recursive:true});
for(const [sheetName,range] of [["方案汇总","A1:L10"],["K_B参数与建议","A1:F14"],["72段预测明细","A1:M20"]]){const image=await wb.render({sheetName,range,scale:1.4,format:"png"});await fs.writeFile(path.join(previewDir,`${sheetName}.png`),new Uint8Array(await image.arrayBuffer()));}
const output=await SpreadsheetFile.exportXlsx(wb); await output.save(outFile); console.log(JSON.stringify({outFile,previewDir,models}));
