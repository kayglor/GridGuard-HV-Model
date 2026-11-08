const ExcelJS = require('exceljs');
const fs = require('fs');
const E = require('./enterprise.js');
const path = require('path');
const DIR = path.join(__dirname, '..', 'model') + path.sep;
const R = E.calc();
const NAME = 'GridGuard HV Services';

// ================= EXCEL MODEL =================
const wb = new ExcelJS.Workbook(); wb.calcProperties = { fullCalcOnLoad: true };
const F = 'Arial', NAVY = 'FF0B2545', BLUE = { argb: 'FF0000FF' };
const font = (o = {}) => Object.assign({ name: F, size: 10 }, o);
const fill = (a) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: a } });
const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } }, box = { top: thin, left: thin, bottom: thin, right: thin };
const hdr = (row, n) => { for (let c = 1; c <= n; c++) { const x = row.getCell(c); x.font = font({ bold: true, color: { argb: 'FFFFFFFF' } }); x.fill = fill(NAVY); x.border = box; x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; } };
const cell = (c, o = {}) => { c.font = font(o.font || {}); c.border = box; c.alignment = Object.assign({ vertical: 'middle', wrapText: true }, o.align || {}); };
const title = (ws, t, s) => { ws.getCell('A1').value = t; ws.getCell('A1').font = font({ size: 14, bold: true, color: { argb: NAVY } }); if (s) { ws.getCell('A2').value = s; ws.getCell('A2').font = font({ italic: true, color: { argb: 'FF595959' } }); } };
const land = (ws) => { ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; };
const CU = '#,##0;(#,##0);-';

// ---- Read Me
const wr = wb.addWorksheet('Read Me', { properties: { tabColor: { argb: NAVY } } });
title(wr, `${NAME}: Enterprise Financial Model (3 years)`, 'A concept-stage model with assumed inputs. Blue cells on yellow are inputs. Currency is generic currency units (CU): relabel to your own currency.');
wr.columns = [{ width: 4 }, { width: 34 }, { width: 90 }];
[['What this is', 'A service business that sells HV maintenance (transformers, cables, switchgear, surveys, call-outs, monitoring) built on the GridGuard HV workflows.'],
 ['Status', 'Concept only. No company has been registered and no licences, insurance or contracts exist. Confirm legal, tax, licensing and insurance requirements locally before launching.'],
 ['Reference estate', 'Each customer site is modelled like the project estate: 10 transformers, 10 cable circuits and 1 11 kV board.'],
 ['How to use', 'Change blue cells in Inputs (sites, prices, hours, salaries, overheads). Projections and Summary recalculate.'],
 ['Sheets', 'Inputs, Projections (revenue, costs, profit, capacity, cash), Summary.'],
 ['Key simplifications', 'No loss carry-forward for tax; straight-line depreciation; no working capital or financing costs; one price escalation rate; no seasonality.'],
 ['Assumptions', 'All figures are assumptions. See HV_Assumptions_Register.xlsx for the wider project assumptions.']].forEach((r, i) => {
  const rr = 4 + i; wr.getCell(`B${rr}`).value = r[0]; cell(wr.getCell(`B${rr}`), { font: { bold: true }, align: { vertical: 'top' } });
  wr.getCell(`C${rr}`).value = r[1]; cell(wr.getCell(`C${rr}`), { align: { vertical: 'top' } }); wr.getRow(rr).height = 34;
});

// ---- Inputs
const wi = wb.addWorksheet('Inputs', { properties: { tabColor: { argb: 'FFF2A900' } } });
title(wi, 'Inputs (assumptions)', 'Edit the blue cells. CU = currency units.');
const inp = (c, nf) => { c.font = font({ color: BLUE }); c.fill = fill('FFFFFF00'); if (nf) c.numFmt = nf; c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = box; };
wi.getRow(4).values = ['Item', 'Unit', 'Year 1', 'Year 2', 'Year 3', 'Note']; hdr(wi.getRow(4), 6);
wi.getRow(5).values = ['Sites under contract', 'sites', ...E.SITES, 'Reference estate each'];
wi.getRow(6).values = ['New sites in year', 'sites'];
wi.getCell('C6').value = { formula: 'C5', result: R.years[0].newSites }; wi.getCell('D6').value = { formula: 'D5-C5', result: R.years[1].newSites }; wi.getCell('E6').value = { formula: 'E5-D5', result: R.years[2].newSites }; wi.getCell('F6').value = 'Drives one-off setup revenue';
for (let r = 5; r <= 6; r++) for (let c = 1; c <= 6; c++) cell(wi.getCell(r, c), { align: { horizontal: c <= 2 || c === 6 ? 'left' : 'center' } });
['C5', 'D5', 'E5'].forEach(a => inp(wi.getCell(a), '0'));
wi.getCell('A8').value = 'Services per site per year'; wi.getCell('A8').font = font({ bold: true, size: 11 });
wi.getRow(9).values = ['Service', 'Unit', 'Units per site per year', 'Price per unit (CU)', 'Person-hours per unit', 'Note']; hdr(wi.getRow(9), 6); wi.getRow(9).height = 30;
E.SERVICES.forEach((s, i) => {
  const rr = 10 + i; wi.getRow(rr).values = [s[0], s[1], s[2], s[3], s[4], s[5]];
  for (let c = 1; c <= 6; c++) cell(wi.getCell(rr, c), { align: { horizontal: c >= 3 && c <= 5 ? 'center' : 'left' } });
  inp(wi.getCell(rr, 3), '0'); inp(wi.getCell(rr, 4), '#,##0'); inp(wi.getCell(rr, 5), '0'); wi.getCell(rr, 4).alignment = { horizontal: 'center', vertical: 'middle' };
  wi.getCell(rr, 1).font = font({ bold: true });
});
const gl = [['Annual price increase', '%', E.ESC, '0.0%', 'Applied from year 2'], ['Direct cost (materials, subcontract, travel, lab)', '% of revenue', E.DIRECT, '0%', ''], ['Payroll on-cost (taxes, benefits)', '% of salary', E.ONCOST, '0%', ''], ['Productive hours per technical FTE per year', 'hours', E.HOURS, '#,##0', ''], ['Target utilisation of technical staff', '%', E.UTIL, '0%', 'Billable share of productive hours'], ['Tax rate', '%', E.TAX, '0%', 'No loss carry-forward'], ['Equipment life', 'years', E.LIFE, '0', 'Straight-line depreciation']];
gl.forEach((g, i) => { const rr = 18 + i; wi.getRow(rr).values = [g[0], g[1], g[2], null, null, g[4]]; for (let c = 1; c <= 6; c++) cell(wi.getCell(rr, c), { align: { horizontal: c === 3 ? 'center' : 'left' } }); inp(wi.getCell(rr, 3), g[3]); });
wi.getCell('A25').value = 'Team'; wi.getCell('A25').font = font({ bold: true, size: 11 });
wi.getRow(26).values = ['Role', 'Technical (Y/N)', 'Year 1 headcount', 'Year 2 headcount', 'Year 3 headcount', 'Monthly salary (CU)']; hdr(wi.getRow(26), 6); wi.getRow(26).height = 30;
E.TEAM.forEach((t, i) => { const rr = 27 + i; wi.getRow(rr).values = [t[0], t[1], ...t[3], t[2]]; for (let c = 1; c <= 6; c++) cell(wi.getCell(rr, c), { align: { horizontal: c === 1 ? 'left' : 'center' } }); [2, 3, 4, 5].forEach(c => inp(wi.getCell(rr, c), '0')); inp(wi.getCell(rr, 6), '#,##0'); wi.getCell(rr, 1).font = font({ bold: true }); wi.getCell(rr, 2).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Y,N"'] }; });
wi.getCell('A37').value = 'Overheads and equipment'; wi.getCell('A37').font = font({ bold: true, size: 11 });
wi.getRow(38).values = ['Item', 'Unit', 'Year 1', 'Year 2', 'Year 3', 'Note']; hdr(wi.getRow(38), 6);
E.OVERHEAD.forEach((o, i) => { const rr = 39 + i; wi.getRow(rr).values = [o[0], 'CU per year', ...o[1], i === 0 ? 'Test sets, thermal camera, VLF, fault locator, vehicles tooling' : '']; for (let c = 1; c <= 6; c++) cell(wi.getCell(rr, c), { align: { horizontal: c >= 3 && c <= 5 ? 'center' : 'left' } }); [3, 4, 5].forEach(c => inp(wi.getCell(rr, c), '#,##0')); wi.getCell(rr, 1).font = font({ bold: true }); });
// capex is row 39; overheads rows 40-45
[46, 16, 14, 14, 14, 52].forEach((w, i) => wi.getColumn(i + 1).width = w); wi.views = [{ state: 'frozen', ySplit: 4 }]; land(wi);

// ---- Projections
const wp = wb.addWorksheet('Projections', { properties: { tabColor: { argb: 'FF0E8C85' } } });
title(wp, 'Projections: revenue, costs, profit, capacity and cash', 'All formulas read the Inputs sheet. CU = currency units.');
wp.getRow(4).values = ['Line', 'Unit', 'Year 1', 'Year 2', 'Year 3']; hdr(wp.getRow(4), 5);
const yc = ['C', 'D', 'E'];
const put = (r, label, unit, fn, vals, nf, o = {}) => {
  wp.getCell(`A${r}`).value = label; wp.getCell(`B${r}`).value = unit;
  yc.forEach((L, y) => { wp.getCell(`${L}${r}`).value = fn ? { formula: fn(L, y), result: vals[y] } : vals[y]; wp.getCell(`${L}${r}`).numFmt = nf; });
  for (let c = 1; c <= 5; c++) { cell(wp.getCell(r, c), { align: { horizontal: c <= 2 ? 'left' : 'right' }, font: o.bold ? { bold: true } : {} }); if (o.fill) wp.getCell(r, c).fill = fill(o.fill); }
};
const sec = (r, t) => { wp.getCell(`A${r}`).value = t; wp.getCell(`A${r}`).font = font({ bold: true, size: 11, color: { argb: NAVY } }); };
const Y = R.years;
put(5, 'Year index', '', null, [1, 2, 3], '0');
put(6, 'Sites under contract', 'sites', (L) => `Inputs!${L}5`, Y.map(y => y.sites), '0');
put(7, 'New sites in year', 'sites', (L) => `Inputs!${L}6`, Y.map(y => y.newSites), '0');
sec(8, 'Revenue');
E.SERVICES.forEach((s, i) => {
  const r = 9 + i, ir = 10 + i;
  put(r, s[0], 'CU', (L) => (i === 6 ? `Inputs!$C${ir}*Inputs!${L}$6*Inputs!$D${ir}*(1+Inputs!$C$18)^(${L}$5-1)` : `Inputs!$C${ir}*Inputs!${L}$5*Inputs!$D${ir}*(1+Inputs!$C$18)^(${L}$5-1)`), Y.map(y => y.rev[i]), CU);
});
put(16, 'Total revenue', 'CU', (L) => `SUM(${L}9:${L}15)`, Y.map(y => y.revenue), CU, { bold: true, fill: 'FFF2F2F2' });
put(17, 'Direct costs', 'CU', (L) => `${L}16*Inputs!$C$19`, Y.map(y => y.direct), CU);
put(18, 'Gross margin', 'CU', (L) => `${L}16-${L}17`, Y.map(y => y.gross), CU, { bold: true });
put(19, 'Gross margin %', '%', (L) => `${L}18/${L}16`, Y.map(y => y.gross / y.revenue), '0.0%');
sec(20, 'Operating costs');
put(21, 'Payroll (salary plus on-cost)', 'CU', (L) => `SUMPRODUCT(Inputs!$F$27:$F$35,Inputs!${L}$27:${L}$35)*12*(1+Inputs!$C$20)`, Y.map(y => y.payroll), CU);
E.OVERHEAD.slice(1).forEach((o, i) => put(22 + i, o[0], 'CU', (L) => `Inputs!${L}${40 + i}`, Y.map(y => y.ov[i]), CU));
put(28, 'Total overheads', 'CU', (L) => `SUM(${L}22:${L}27)`, Y.map(y => y.overheads), CU, { bold: true });
put(29, 'EBITDA', 'CU', (L) => `${L}18-${L}21-${L}28`, Y.map(y => y.ebitda), CU, { bold: true, fill: 'FFF2F2F2' });
put(30, 'Depreciation', 'CU', (L) => `SUM(Inputs!$C$39:${L}$39)/Inputs!$C$24`, Y.map(y => y.dep), CU);
put(31, 'EBIT (profit before tax)', 'CU', (L) => `${L}29-${L}30`, Y.map(y => y.ebit), CU, { bold: true });
put(32, 'Tax', 'CU', (L) => `MAX(0,${L}31)*Inputs!$C$23`, Y.map(y => y.tax), CU);
put(33, 'Net profit', 'CU', (L) => `${L}31-${L}32`, Y.map(y => y.net), CU, { bold: true, fill: 'FFF2F2F2' });
put(34, 'Net margin %', '%', (L) => `${L}33/${L}16`, Y.map(y => y.net / y.revenue), '0.0%');
sec(35, 'Capacity');
put(36, 'Required person-hours', 'hours', (L) => `Inputs!${L}$5*SUMPRODUCT(Inputs!$C$10:$C$15,Inputs!$E$10:$E$15)+Inputs!${L}$6*Inputs!$C$16*Inputs!$E$16`, Y.map(y => y.hours), '#,##0');
put(37, 'Technical FTE', 'FTE', (L) => `SUMPRODUCT((Inputs!$B$27:$B$35="Y")*Inputs!${L}$27:${L}$35)`, Y.map(y => y.fte), '0');
put(38, 'Available billable hours', 'hours', (L) => `${L}37*Inputs!$C$21*Inputs!$C$22`, Y.map(y => y.avail), '#,##0');
put(39, 'Utilisation', '%', (L) => `${L}36/${L}38`, Y.map(y => y.util), '0%');
put(40, 'Capacity check', '', (L) => `IF(${L}39<=1,"OK","Short: add staff or subcontract")`, Y.map(y => (y.util <= 1 ? 'OK' : 'Short: add staff or subcontract')), '@');
sec(41, 'Cash');
put(42, 'Equipment purchase', 'CU', (L) => `Inputs!${L}39`, Y.map(y => y.capex), CU);
put(43, 'Operating cash flow (net profit plus depreciation)', 'CU', (L) => `${L}33+${L}30`, Y.map(y => y.opcash), CU);
put(44, 'Net cash flow', 'CU', (L) => `${L}43-${L}42`, Y.map(y => y.ncf), CU, { bold: true });
put(45, 'Cumulative cash', 'CU', (L) => `SUM($C44:${L}44)`, Y.map(y => y.cum), CU, { bold: true, fill: 'FFF2F2F2' });
wp.getCell('A46').value = 'Peak funding need (start-up cash required)'; wp.getCell('B46').value = 'CU'; wp.getCell('C46').value = { formula: '-MIN(0,MIN(C45:E45))', result: R.peak }; wp.getCell('C46').numFmt = CU;
for (let c = 1; c <= 5; c++) { cell(wp.getCell(46, c), { font: { bold: true }, align: { horizontal: c <= 2 ? 'left' : 'right' } }); wp.getCell(46, c).fill = fill('FFFCE8B2'); }
put(48, 'Recurring revenue per site (excluding setup)', 'CU', (L) => `(${L}16-${L}15)/${L}6`, Y.map(y => (y.revenue - y.rev[6]) / y.sites), CU);
wp.addConditionalFormatting({ ref: 'C40:E40', rules: [{ type: 'expression', priority: 1, formulae: ['C40<>"OK"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
wp.addConditionalFormatting({ ref: 'C33:E33', rules: [{ type: 'expression', priority: 1, formulae: ['C33<0'], style: { font: { color: { argb: 'FFC0392B' } } } }] });
[48, 10, 16, 16, 16].forEach((w, i) => wp.getColumn(i + 1).width = w); wp.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }]; land(wp);

// ---- Summary
const ws = wb.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFC0392B' } } });
title(ws, `${NAME}: three-year summary`, 'Linked to Projections. Assumption-based; not a forecast.');
ws.getRow(4).values = ['Measure', 'Year 1', 'Year 2', 'Year 3']; hdr(ws.getRow(4), 4);
const sm = [['Sites under contract', 6, '0', (y) => y.sites], ['Revenue (CU)', 16, CU, (y) => y.revenue], ['Gross margin (CU)', 18, CU, (y) => y.gross], ['EBITDA (CU)', 29, CU, (y) => y.ebitda], ['Net profit (CU)', 33, CU, (y) => y.net], ['Net margin', 34, '0.0%', (y) => y.net / y.revenue], ['Technical FTE', 37, '0', (y) => y.fte], ['Utilisation', 39, '0%', (y) => y.util], ['Cumulative cash (CU)', 45, CU, (y) => y.cum]];
sm.forEach((m, i) => { const rr = 5 + i; ws.getCell(`A${rr}`).value = m[0]; ['B', 'C', 'D'].forEach((L, y) => { ws.getCell(`${L}${rr}`).value = { formula: `Projections!${yc[y]}${m[1]}`, result: m[3](Y[y]) }; ws.getCell(`${L}${rr}`).numFmt = m[2]; }); for (let c = 1; c <= 4; c++) cell(ws.getCell(rr, c), { align: { horizontal: c === 1 ? 'left' : 'right' }, font: c === 1 ? { bold: true } : {} }); });
ws.getCell('A15').value = 'Peak funding need (CU)'; ws.getCell('B15').value = { formula: 'Projections!C46', result: R.peak }; ws.getCell('B15').numFmt = CU;
for (let c = 1; c <= 4; c++) { cell(ws.getCell(15, c), { font: { bold: true }, align: { horizontal: c === 1 ? 'left' : 'right' } }); ws.getCell(15, c).fill = fill('FFFCE8B2'); }
[34, 16, 16, 16].forEach((w, i) => ws.getColumn(i + 1).width = w);
wb.xlsx.writeFile(DIR + 'GridGuard_HV_Enterprise_Model.xlsx').then(() => console.log('xlsx written'));

// ================= RTF PLAN =================
const colours = ['0B2545', 'FFFFFF', 'F4F6F8', '595959', 'FCE8B2'];
const rgb = (h) => `\\red${parseInt(h.slice(0, 2), 16)}\\green${parseInt(h.slice(2, 4), 16)}\\blue${parseInt(h.slice(4, 6), 16)};`;
const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}').replace(/[^\x00-\x7F]/g, (c) => `\\u${c.charCodeAt(0)}?`);
const p = (t, o = {}) => `{\\pard\\sa${o.sa ?? 100}\\sb${o.sb ?? 0}${o.li ? `\\li${o.li}\\fi-${o.hang || 0}` : ''} ${o.cf ? `\\cf${o.cf}` : ''}\\fs${o.fs || 20}${o.b ? '\\b' : ''}${o.i ? '\\i' : ''} ${esc(t)}\\par}\n`;
const h1 = (t) => p(t, { fs: 28, b: 1, cf: 1, sb: 260, sa: 110 });
const h2 = (t) => p(t, { fs: 22, b: 1, cf: 1, sb: 160, sa: 70 });
const bullet = (t) => p('\u2022\t' + t, { li: 360, hang: 360, sa: 60 });
function table(headers, rows, widths, center = []) {
  let x = 0; const cells = widths.map(w => (x += w));
  const bord = '\\clbrdrt\\brdrs\\brdrw10\\brdrcf4\\clbrdrl\\brdrs\\brdrw10\\brdrcf4\\clbrdrb\\brdrs\\brdrw10\\brdrcf4\\clbrdrr\\brdrs\\brdrw10\\brdrcf4';
  const row = (vals, head, ri) => {
    let s = `{\\trowd\\trgaph70\\trleft0${head ? '\\trhdr' : ''}\\trkeep`;
    vals.forEach((v, i) => { const sp = (v && typeof v === 'object') ? v : { t: v }; const fill = head ? 1 : sp.fill || (ri % 2 ? 3 : 0); s += `${bord}${fill ? `\\clcbpat${fill}` : ''}\\clvertalc\\cellx${cells[i]}`; });
    s += '\n';
    vals.forEach((v, i) => { const sp = (v && typeof v === 'object') ? v : { t: v }; const al = head || center.includes(i) ? '\\qc' : '\\ql'; s += `\\pard\\intbl${al}\\fs17${head || sp.b ? '\\b' : ''}${head ? '\\cf2' : ''} ${esc(sp.t)}\\cell\n`; });
    return s + '\\row}\n';
  };
  return row(headers, true, 0) + rows.map((r, i) => row(r, false, i)).join('') + p('', { sa: 80 });
}
const n = (v) => Math.round(v).toLocaleString('en-US').replace(/^-/, '-');
const W = 9638;
let b = '';
b += p('ENTERPRISE PLAN  |  CONCEPT STAGE', { b: 1, cf: 4, sa: 40, sb: 400 });
b += p(NAME, { fs: 44, b: 1, cf: 1, sa: 80 });
b += p('HV maintenance services for estates, campuses and industrial sites', { fs: 24, cf: 4, sa: 160 });
b += table(['Item', 'Detail'], [
  [{ t: 'Enterprise', b: 1 }, `${NAME} (working name)`],
  [{ t: 'Type', b: 1 }, 'Service business: preventive and corrective maintenance, testing, condition monitoring and reporting for 11 kV networks and 11/0.415 kV transformers'],
  [{ t: 'Status', b: 1 }, 'Concept. Not registered; no licences, insurance or contracts. Confirm legal, tax and licensing requirements locally.'],
  [{ t: 'Founding asset', b: 1 }, 'The GridGuard HV program: annual plan, Q1 delivery pack, 82 atomic tasks, checklists, estate layout and AI data model'],
  [{ t: 'Companion model', b: 1 }, 'GridGuard_HV_Enterprise_Model.xlsx (all numbers below come from it)'],
  [{ t: 'Currency', b: 1 }, 'Generic currency units (CU). All figures are assumptions, not forecasts.'],
], [2200, W - 2200]);

b += h1('1. Opportunity and mission');
b += p('Estates, campuses and industrial parks own 11 kV boards, cables and transformers but rarely have in-house HV engineers. Maintenance is often reactive, test records sit on paper, and outages are badly planned. GridGuard HV Services sells a complete, documented maintenance service built on one proven method: plan, baseline, test, certify, monitor, improve.');
b += p('Mission: keep HV estates safe and available, with every test traceable and every outage planned.', { b: 1 });

b += h1('2. Services and pricing (assumed)');
b += table(['Service', 'Unit', 'Per site per year', 'Price (CU)', 'Person-hours / unit'], E.SERVICES.map(s => [{ t: s[0], b: 1 }, s[1], String(s[2]), n(s[3]), String(s[4])]), [3300, 1600, 1500, 1400, 1838], [2, 3, 4]);
b += p('Reference site: 10 transformers, 10 cable circuits and 1 11 kV board, the same as the project estate. Prices rise 3% a year from year 2. Setup is billed once on each new site.', { fs: 18, cf: 4 });

b += h1('3. Customers');
['Residential and mixed-use estates with private HV networks', 'Commercial parks, malls and hospitals that need high availability', 'Campuses and industrial parks with several substations', 'Developers handing over new estates who need an operating contractor'].forEach(t => b += bullet(t));
b += p('Entry offer: a fixed-price baseline survey (thermography and PD) that produces a health index and a costed maintenance plan.');

b += h1('4. Operating model');
b += p('The enterprise runs the same workflows as the project, with the seven improvement targets built in:');
['Test records and certificates entered once on mobile and compiled automatically', 'Thermography and PD analysis with AI hot-spot detection and a rule-based A to D health grade', 'Load-aware shutdown planning and bulk work-order creation', 'Fault location and RCA supported by relay waveform data and GIS', 'Permit and switching programme templates checked against the single-line diagram', 'Asset register and data standard so AI can learn from clean records', 'Automatic KPI and audit reporting'].forEach(t => b += bullet(t));
b += p('Human gates never change: isolation, switching, return to service and safety sign-off stay with a competent Authorised Person. AI is advisory only.', { b: 1 });

b += h1('5. Team and organisation');
b += table(['Role', 'Technical', 'Year 1', 'Year 2', 'Year 3', 'Monthly salary (CU)'], E.TEAM.map(t => [{ t: t[0], b: 1 }, t[1], String(t[3][0]), String(t[3][1]), String(t[3][2]), n(t[2])]), [3100, 1100, 1000, 1000, 1000, 2438], [1, 2, 3, 4, 5]);
b += p(`Total headcount: ${E.TEAM.reduce((a, t) => a + t[3][0], 0)} in year 1, ${E.TEAM.reduce((a, t) => a + t[3][1], 0)} in year 2, ${E.TEAM.reduce((a, t) => a + t[3][2], 0)} in year 3. Technical staff are the capacity limit: utilisation is ${(R.years[0].util * 100).toFixed(0)}%, ${(R.years[1].util * 100).toFixed(0)}% and ${(R.years[2].util * 100).toFixed(0)}% against a 70% billable target.`);
b += h2('Governance');
['Operations manager is accountable for safety, quality and delivery', 'HSE officer (from year 2) audits permits, RAMS and drills; in year 1 the operations manager covers this', 'Competency and instrument calibration registers kept current; staff stood down when authorisations expire', 'Quarterly management review using the KPI dashboard'].forEach(t => b += bullet(t));

b += h1('6. Compliance checklist (confirm locally)');
b += p('These items must be checked with a local adviser before trading. This plan does not state what the law requires.', { i: 1, cf: 4 });
table; ['Company registration and tax registration', 'Electrical contractor licence or competence scheme membership, where required', 'Public liability and professional indemnity insurance', 'Health and safety policy, risk assessments and permit-to-work system', 'Client site induction and permit requirements', 'Calibration certificates for all test instruments', 'Data protection and contract terms for client data stored in the CMMS'].forEach(t => b += bullet(t));

b += h1('7. Financial summary (assumed, CU)');
const Yr = R.years;
b += table(['Measure', 'Year 1', 'Year 2', 'Year 3'], [
  ['Sites under contract', ...Yr.map(y => String(y.sites))], ['Revenue', ...Yr.map(y => n(y.revenue))], ['Direct costs', ...Yr.map(y => n(y.direct))], ['Gross margin', ...Yr.map(y => n(y.gross))],
  ['Payroll', ...Yr.map(y => n(y.payroll))], ['Overheads', ...Yr.map(y => n(y.overheads))], ['EBITDA', ...Yr.map(y => n(y.ebitda))], ['Depreciation', ...Yr.map(y => n(y.dep))],
  ['Tax', ...Yr.map(y => n(y.tax))], [{ t: 'Net profit', b: 1 }, ...Yr.map(y => ({ t: n(y.net), b: 1 }))], ['Equipment purchase', ...Yr.map(y => n(y.capex))], [{ t: 'Cumulative cash', b: 1 }, ...Yr.map(y => ({ t: n(y.cum), b: 1 }))],
], [3600, 2013, 2013, 2012], [1, 2, 3]);
b += p(`Peak funding need: about ${n(R.peak)} CU (equipment, software and the year-1 operating loss). Year 1 is an investment year. Cumulative cash is close to break-even by the end of year 3 (${n(Yr[2].cum)} CU).`, { b: 1 });
b += p('Main drivers: number of sites, price per service, technical staff utilisation and equipment cost. Change them in the Inputs sheet.');

b += h1('8. Go-to-market');
['Pilot on the project estate to produce a reference case, test records and a health-index report', 'Fixed-price baseline survey as the entry offer, then a 3-year maintenance contract', 'Bundles: Essential (PM and tests), Plus (adds surveys and subscription), Full (adds corrective call-outs)', 'Partnerships with utilities, OEM service agents and calibration labs', 'Bids to estate managers and developers; case studies built from anonymised results'].forEach(t => b += bullet(t));

b += h1('9. Risks and responses');
b += table(['Risk', 'Response'], [
  ['Too few sites in year 1, so overheads are not covered', 'Keep the team lean (7 people), subcontract protection work, sell the baseline survey first'],
  ['Safety incident or uninsured loss', 'Strict permit-to-work, competence checks, insurance in place before first job'],
  ['Technical staff overloaded', 'Watch utilisation monthly; subcontract peaks; hire ahead of site wins'],
  ['Prices too low for the work involved', 'Track hours per job from day one and reprice after the first two sites'],
  ['AI results not trusted', 'Rule-based first; engineers confirm all grades; report accuracy quarterly'],
  ['Equipment cost higher than assumed', 'Hire or lease specialist sets first; buy once utilisation is proven'],
], [4000, W - 4000]);

b += h1('10. 90-day launch plan');
b += table(['Days', 'Step', 'Output'], [
  ['1-15', 'Confirm legal form, name, tax and insurance requirements with an adviser', 'Registration and insurance checklist signed off'],
  ['1-30', 'Finalise service catalogue, price list and contract terms', 'Rate card and standard contract'],
  ['15-45', 'Recruit or contract the core team (manager, test engineer, cable engineer, 2 technicians, planner)', 'Team in place; competency register complete'],
  ['30-60', 'Procure or hire test equipment; set up CMMS, permit system and data standard', 'Tools ready; calibration certificates filed'],
  ['45-75', 'Run the pilot baseline survey and Q1 plan on the project estate', 'Reference report and health index'],
  ['60-90', 'First three proposals to prospective sites using the reference report', 'Proposals issued; pipeline tracked'],
], [900, 5200, 3538], [0]);

b += h1('11. Assumptions and next steps');
['All prices, hours, salaries, overheads and site counts are assumptions: replace them with quotes and timed jobs.', 'Currency is generic: set real currency and local rates.', 'Next: validate prices with two prospective customers; get insurance and licensing quotes; refine the model.'].forEach(t => b += bullet(t));

const rtf = `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1\\deflang2057
{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}}
{\\colortbl;${colours.map(rgb).join('')}}
{\\info{\\title ${NAME} Enterprise Plan}{\\author Facility Manager}}
\\paperw11906\\paperh16838\\margl1134\\margr1134\\margt1134\\margb1134
{\\footer\\pard\\qc\\fs16 ${NAME} enterprise plan (concept, assumptions) - Page {\\field{\\*\\fldinst PAGE}{\\fldrslt 1}}\\par}
\\f0\\fs20
${b}}`;
fs.writeFileSync(DIR + 'GridGuard_HV_Enterprise_Plan.rtf', rtf, 'ascii');
console.log('rtf written', rtf.length);
