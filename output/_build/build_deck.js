const pptxgen = require('./node/node_modules/pptxgenjs');
const fs = require('fs');
const { fmt, pct, mx } = require('./deck_common.js');
const [,, packJson, marketJson, outPath] = process.argv;
const D = JSON.parse(fs.readFileSync(packJson)); const MK = JSON.parse(fs.readFileSync(marketJson));
const I = D.info, T = D.T, A = D.A, B = D.B, S1 = B.S.S1.res, F = D.fund;
const NAME = I.name; const NAVY = '17324D', INK = '1B1F2A', BRASS = 'B8893B', TEAL = '2A7F83', GREY = '6B7280', LIGHT = 'F2F4F7', RED = 'A23B3B', WHITE = 'FFFFFF', MID = 'D5DAE1';
const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE';
pres.author = 'Jalour Developments'; pres.company = 'Jalour Developments'; pres.title = `${NAME} - Investment Presentation (EGP ${T} million)`; pres.subject = 'Strictly private and confidential';
pres.theme = { headFontFace: 'Cambria', bodyFontFace: 'Calibri' };
const W = 13.33;
pres.defineSlideMaster({ title: 'CONTENT', background: { color: WHITE },
  objects: [ { text: { text: `Jalour Developments  |  ${NAME}  |  Strictly private and confidential. Projections, not guarantees.`, options: { x: 0.6, y: 7.05, w: 10, h: 0.3, fontSize: 9, color: GREY, margin: 0, isTextBox: true } } } ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: GREY } });
pres.defineSlideMaster({ title: 'DARK', background: { color: NAVY }, objects: [] });
let sn = 0;
function content(title, notes) {
  const s = pres.addSlide({ masterName: 'CONTENT' }); sn++;
  s.addText(title, { x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontFace: 'Cambria', fontSize: 28, bold: true, color: NAVY, valign: 'middle', margin: 0, isTextBox: true, fit: 'shrink' });
  s.addNotes(notes || `Figures are taken from the Financial Annex (Summary sheet). All returns are projections.`);
  return s;
}
function bullets(s, items, x, y, w, h, fs = 16, color = INK) {
  const arr = items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } }));
  s.addText(arr, { x, y, w, h, fontSize: fs, color, valign: 'top', margin: 0, isTextBox: true });
}
function card(s, x, y, w, h, fill = LIGHT) { s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }); }
function stat(s, x, y, w, big, label, dark = false) {
  card(s, x, y, w, 1.55, dark ? NAVY : LIGHT);
  s.addText(big, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: 0.8, fontFace: 'Cambria', fontSize: 30, bold: true, color: dark ? WHITE : NAVY, valign: 'middle', margin: 0, isTextBox: true, fit: 'shrink' });
  s.addText(label, { x: x + 0.2, y: y + 0.92, w: w - 0.4, h: 0.55, fontSize: 12, color: dark ? 'D5DAE1' : GREY, valign: 'top', margin: 0, isTextBox: true });
}
function tbl(s, rows, x, y, w, colW, fs = 11, opts = {}) {
  const body = rows.map((r, ri) => r.map((c, ci) => ({ text: String(c), options: { bold: ri === 0 || (opts.boldFirst && ci === 0), color: ri === 0 ? WHITE : INK, fill: { color: ri === 0 ? NAVY : (ri % 2 ? WHITE : LIGHT) }, fontSize: fs, align: (ci === 0 || opts.left) ? 'left' : 'right', valign: 'middle' } })));
  s.addTable(body, { x, y, w, colW, border: { type: 'solid', pt: 0.5, color: MID }, margin: [0.04, 0.08, 0.04, 0.08], rowH: opts.rowH || 0.3, autoPage: false });
}
function source(s, text, y = 6.72) { s.addText('Source: ' + text, { x: 0.6, y, w: 12.1, h: 0.3, fontSize: 9, color: GREY, italic: true, margin: 0, isTextBox: true }); }
const yr = (a) => a.map(r => 'Y' + r.year);
const bal = 'The balance of project funding is provided by Jalour sponsor equity, project collections and other capital sources.';

// 1 Cover
{ const s = pres.addSlide({ masterName: 'DARK' }); sn++;
  s.addText('STRICTLY PRIVATE AND CONFIDENTIAL', { x: 0.8, y: 0.6, w: 8, h: 0.3, fontSize: 11, color: BRASS, bold: true, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText(NAME, { x: 0.8, y: 2.0, w: 11.5, h: 1.2, fontFace: 'Cambria', fontSize: 54, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText('Commercial and office development, El Mostakbal City, Cairo', { x: 0.8, y: 3.2, w: 11.5, h: 0.5, fontSize: 22, color: 'D5DAE1', margin: 0, isTextBox: true });
  s.addText(`Private placement: investor ticket of EGP ${T} million`, { x: 0.8, y: 4.3, w: 11.5, h: 0.5, fontSize: 20, bold: true, color: BRASS, margin: 0, isTextBox: true });
  s.addText('Investment presentation  |  Jalour Developments, a subsidiary of Al Jalal Holding  |  October 2026', { x: 0.8, y: 6.4, w: 11.5, h: 0.4, fontSize: 13, color: 'D5DAE1', margin: 0, isTextBox: true });
  s.addNotes('Cover. All figures are projections based on the Jalour financial model; see the Financial Annex.'); }
// 2 Highlights
{ const s = content('Investment highlights');
  const w = 2.9, g = 0.17, x0 = 0.6;
  stat(s, x0, 1.5, w, fmt(I.bua, 0) + ' m2', `Built-up area, ${I.height}: ${fmt(I.retail, 0)} m2 retail and ${fmt(I.office, 0)} m2 offices`);
  stat(s, x0 + (w + g), 1.5, w, `Month ${I.delivery}`, `Handover; sales from month ${I.sales_start}, construction from month ${I.con_start} for ${I.con_months} months`);
  stat(s, x0 + 2 * (w + g), 1.5, w, 'EGP ' + fmt(I.total80 / 1000, 2) + ' bn', 'Sales plan: 80% of units sold over the sales window, at modelled prices');
  stat(s, x0 + 3 * (w + g), 1.5, w, '35 / 65', 'Revenue split between the landowner (Al Ahly Sabbour) and Jalour', true);
  bullets(s, [
    `Land owned by Al Ahly Sabbour; Jalour designs, permits, builds (core and shell), markets, sells, collects and operates.`,
    `Landlord receives 35% of collections with a minimum guarantee of EGP ${fmt(I.guarantee, 0)} million over six years; Jalour keeps 65%.`,
    `Option A: units with a face value of EGP ${fmt(A.face, 0)} million (2.0x the ticket) at launch list prices. Projected investor IRR ${pct(A.inv.Base.irr)} (base).`,
    `Option B: cash return of EGP ${fmt(B.mult * T, 0)} million (${mx(B.mult, 1)} the ticket) in 12 quarterly instalments after a 24-month grace period. Projected investor IRR ${pct(S1.Base.irr)}.`,
    `Project-level exposure only: the investment has no exposure to other Jalour projects.`], 0.6, 3.5, 12.1, 3.0, 17);
  source(s, 'Jalour financial model; Financial Annex, Summary sheet. IRRs are projections, not guarantees.'); }
// 3 Transaction overview
{ const s = content('Transaction overview');
  tbl(s, [['Term', 'Summary'],
    ['Investor ticket', `EGP ${fmt(T, 0)} million (USD ${fmt(T / 48, 1)} million at EGP 48 per USD)`],
    ['Issuer', 'Project company for this development (name, jurisdiction and instrument to be confirmed by counsel)'],
    ['Use of proceeds', `Down payment to the landowner, guarantee instalments and working capital to the funding peak, and pre-construction costs (total modelled need EGP ${fmt(F.total, 1)} million)`],
    ['Option A: units', `Units with a face value of EGP ${fmt(A.face, 0)} million at launch list prices, allocated from the first two sales tranches; resale rights subject to consents`],
    ['Option B: cash', `EGP ${fmt(B.mult * T, 0)} million in 12 equal quarterly instalments of EGP ${fmt(B.mult * T / 12, 2)} million; first payment month ${B.pay_months[0]}, last month ${B.pay_months[11]}`],
    ['Choice', 'The investor chooses Option A or Option B at signing'],
    ['Balance of funding', bal],
    ['Further capital', 'Jalour may raise further capital at project or holding level, subject to the investor\'s stated rights (ranking, collateral, information rights)']],
    0.6, 1.45, 12.1, [2.4, 9.7], 13, { left: true, boldFirst: true, rowH: 0.5 });
  source(s, 'Term summary; final terms are set by the definitive agreements.'); }
// 4 Market supply and demand
{ const s = content('Market: Cairo office and commercial supply and demand');
  const f = Object.fromEntries(MK.facts.map(x => [x.id, x]));
  stat(s, 0.6, 1.5, 2.9, '73%', 'of Cairo office stock, current and future, is in New Cairo');
  stat(s, 3.67, 1.5, 2.9, '~9%', 'office vacancy, stable (Q1 2026); Grade A occupancy about 90%');
  stat(s, 6.74, 1.5, 2.9, '+82%', 'office pipeline to 2029 relative to current stock: supply is the main risk');
  stat(s, 9.81, 1.5, 2.9, 'USD 348', 'Grade A rent per m2 per year in New Cairo business parks (Q2 2026)');
  bullets(s, [f.mc_scale.text + '.', f.mc_infra.text + '.', f.jll_retail.text + '.', f.hap.text + '.'], 0.6, 3.35, 12.1, 3.2, 16);
  source(s, 'Knight Frank (via Arabian Business, Jul 2026; Enterprise, Jul 2026); JLL Q1 2026 (via Invest-Gate); MIDAR; Nawy; Bayut; Invest-Gate. Knight Frank and JLL use different office-stock definitions.', 6.6); }
// 5 Market pricing and macro
{ const s = content('Market: pricing evidence and demand drivers');
  tbl(s, [['Indicator', 'Evidence', 'Model assumption'],
    ['Office prices', 'Mostakbal City listings: EGP 106,000 to 120,000 per m2 (asking)', `Launch EGP ${fmt(D.prices.launch[1] * 1000, 0)} per m2, rising to ${fmt(D.prices.final[1] * 1000, 0)} at the delivery list`],
    ['Retail prices', 'New Cairo listings: EGP 139,000 to 325,000 per m2 (asking, individual listings)', `Launch EGP ${fmt(D.prices.launch[0] * 1000, 0)} per m2, rising to ${fmt(D.prices.final[0] * 1000, 0)}`],
    ['Price growth', 'New Cairo secondary prices +3.1% y/y (Q1 2026); broker estimates 10 to 16% a year for 2026', `Launch to delivery list: +${fmt((D.prices.final[0] / D.prices.launch[0] - 1) * 100, 0)}% retail, +${fmt((D.prices.final[1] / D.prices.launch[1] - 1) * 100, 0)}% offices; offices +${fmt(D.pricecmp.off_vs_ask * 100, 0)}% on today's asking, about ${fmt(D.pricecmp.off_cagr * 100, 0)}% a year`],
    ['Interest rates', 'CBE overnight deposit rate 19.00% (Sep 2026); urban inflation 14.5% (Aug 2026)', 'Discount rate 14% in the model'],
    ['Developer sales', 'Top 10 developers EGP 670 bn in H1 2026, +2.9% y/y; units sold -5%', 'Sales over about 3 years from launch'],
    ['Payment plans', 'Typically 8 to 12 years with 1.5% to 10% down payment', 'Down payments 5% to 20%; instalments 2 to 8 years']],
    0.6, 1.5, 12.1, [1.9, 5.6, 4.6], 12, { left: true, boldFirst: true, rowH: 0.6 });
  s.addText(`Reading across: the sales plan needs price growth that sits at the upper end of recent evidence. The downside case on slide ${D.combined ? 15 : 14} tests lower prices and later sales.`, { x: 0.6, y: 5.65, w: 12.1, h: 0.7, fontSize: 13, color: INK, margin: 0, isTextBox: true });
  source(s, 'Property Finder listings (2026); JLL via Enterprise (2026); Sands of Wealth (2026); CBE via Bnok24; CAPMAS via Ahram Online; Daily News Egypt; Bayut. Listings are asking prices, not transactions.', 6.55); }
// 6 The project
{ const s = content('The project');
  tbl(s, [['Item', 'Detail'], ['Land area', fmt(I.land, 0) + ' m2'], ['Footprint (30%)', fmt(I.footprint, 0) + ' m2'], ['Height', I.height], ['Retail built-up area', fmt(I.retail, 0) + ' m2'], ['Office built-up area', fmt(I.office, 0) + ' m2'], ['Total built-up area', fmt(I.bua, 0) + ' m2'],
    ['Finishing standard', 'Core and shell'], ['Sales start', `Month ${I.sales_start}`], ['Construction', `Month ${I.con_start} to ${I.delivery} (${I.con_months} months)`], ['Handover', `Month ${I.delivery}`]], 0.6, 1.5, 6.6, [2.9, 3.7], 14, { left: true, boldFirst: true, rowH: 0.45 });
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Built-up area', labels: ['Retail', 'Offices'], values: [I.retail, I.office] }], { x: 7.6, y: 1.5, w: 5.1, h: 4.2, holeSize: 55, chartColors: [BRASS, NAVY], showLegend: true, legendPos: 'b', legendFontSize: 12, showPercent: true, showValue: false, dataLabelColor: WHITE, dataLabelFontSize: 12, showTitle: true, title: 'Built-up area mix (m2)', titleFontSize: 13, titleColor: INK });
  source(s, 'Jalour financial model (project summary). Jalour retains 20% of units, outside the sales plan and outside this offering.', 6.6); }
// 7 Partnership
{ const s = content('Partnership with Al Ahly Sabbour and the revenue share');
  card(s, 0.6, 1.5, 3.0, 1.3); s.addText([{ text: 'Al Ahly Sabbour', options: { bold: true, breakLine: true } }, { text: 'Landowner. Receives 35% of collections, subject to a minimum guarantee.' }], { x: 0.75, y: 1.55, w: 2.7, h: 1.2, fontSize: 12, color: INK, margin: 0, isTextBox: true, valign: 'top' });
  card(s, 0.6, 3.0, 3.0, 1.7); s.addText([{ text: 'Jalour', options: { bold: true, breakLine: true } }, { text: 'Developer and operator. Keeps 65% of collections. Carries design, permits, core-and-shell construction, marketing, sales, collection and operation costs.' }], { x: 0.75, y: 3.05, w: 2.7, h: 1.6, fontSize: 12, color: INK, margin: 0, isTextBox: true, valign: 'top' });
  bullets(s, [`Minimum guarantee EGP ${fmt(I.guarantee, 0)} million over six years; down payment EGP ${fmt(I.dp, 0)} million (10%) paid up front.`, 'Landlord share is 35% of collections or the guaranteed instalment, whichever is higher.', `Any excess of 35% over the guarantee is settled from year ${I.excess_years}.`, 'Allocation of units, assignment of collections or any encumbrance affecting the landlord share requires the landlord\'s written consent.'], 0.6, 4.85, 12.1, 1.8, 13);
  const sched = [I.dp].concat(I.sched);
  s.addChart(pres.charts.BAR, [{ name: 'Guaranteed payment to landlord (EGP m)', labels: ['Down payment', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Year 6'], values: sched }], { x: 4.0, y: 1.4, w: 8.7, h: 3.4, barDir: 'col', chartColors: [NAVY], showValue: true, dataLabelFontSize: 11, dataLabelColor: INK, dataLabelPosition: 'outEnd', valAxisLabelFontSize: 11, catAxisLabelFontSize: 11, valGridLine: { color: MID, size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, showTitle: true, title: 'Minimum guarantee schedule (EGP million)', titleFontSize: 13, titleColor: INK });
  source(s, 'Jalour financial model (landlord terms); agreement terms to be confirmed in the data room.', 6.65); }
// 8 Jalour
{ const s = content('Jalour Developments: role and track record');
  card(s, 0.6, 1.5, 5.9, 4.9); s.addText('Role in this project', { x: 0.8, y: 1.6, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: NAVY, margin: 0, isTextBox: true });
  bullets(s, ['Design and permits', 'Construction to core-and-shell standard', 'Marketing, branding and sales', 'Collection of instalments', 'Operation, management and maintenance, directly or through a facility management company', 'Covers project funding gaps from its cash reserves and cash from other projects it is constructing (to be documented)'], 0.8, 2.1, 5.5, 4.2, 14);
  card(s, 6.8, 1.5, 5.9, 4.9, 'FBF3E4'); s.addText('Track record and team', { x: 7.0, y: 1.6, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: BRASS, margin: 0, isTextBox: true });
  bullets(s, ['Completed projects, area delivered and units handed over', 'Management team biographies', 'Audited financial statements of Jalour and Al Jalal Holding', 'References from landlords, contractors and buyers', 'All provided after the investor signs a non-disclosure agreement'], 7.0, 2.1, 5.5, 4.2, 14);
  source(s, 'Jalour Developments is the real estate arm of Al Jalal Holding, Cairo. Sponsor materials are provided after the investor signs a non-disclosure agreement.', 6.6); }
// 9 Sales plan and pricing
{ const s = content('Sales plan and pricing');
  const tr = D.tranches; const lab = tr.map(t => 'M' + t.month);
  s.addChart(pres.charts.BAR, [{ name: 'Retail', labels: lab, values: tr.map(t => t.comm) }, { name: 'Offices', labels: lab, values: tr.map(t => t.admin) }], { x: 0.6, y: 1.4, w: 7.4, h: 4.4, barDir: 'col', barGrouping: 'stacked', chartColors: [BRASS, NAVY], showLegend: true, legendPos: 'b', legendFontSize: 11, valAxisLabelFontSize: 11, catAxisLabelFontSize: 11, valGridLine: { color: MID, size: 0.5 }, catGridLine: { style: 'none' }, showTitle: true, title: 'Sales value by quarter of sale (EGP million, 80% of units)', titleFontSize: 13, titleColor: INK });
  tbl(s, [['EGP thousand per m2', 'Launch', 'Delivery list', 'Average'], ['Retail', fmt(D.prices.launch[0], 0), fmt(D.prices.final[0], 0), fmt(D.prices.avg[0], 0)], ['Offices', fmt(D.prices.launch[1], 0), fmt(D.prices.final[1], 0), fmt(D.prices.avg[1], 0)]], 8.3, 1.6, 4.4, [1.8, 0.85, 0.85, 0.9], 12, { rowH: 0.45 });
  bullets(s, [`Total sales at 100% of units: EGP ${fmt(I.total100, 0)} million; plan sells 80%: EGP ${fmt(I.total80, 0)} million.`, 'Payment plans: down payment 5% to 20%, a second equal payment three months later, instalments of 2 to 8 years and a delivery payment of 20% to 25%.', 'Delivery payments are collected on a contractual milestone ahead of handover (see risks).'], 8.3, 3.2, 4.4, 3.4, 12);
  source(s, 'Jalour financial model (sales sheet). Prices are model assumptions; see slide 5 for market evidence.', 6.6); }
// 10 Collections and cash flow
{ const s = content('Collections, costs and cash flow');
  const a = D.annual.slice(0, 10); const lab = a.map(r => 'Y' + r.year);
  s.addChart([{ type: pres.charts.BAR, data: [{ name: 'Collections', labels: lab, values: a.map(r => r.coll) }, { name: 'Landlord payments', labels: lab, values: a.map(r => -r.land) }, { name: 'Construction, commission, SG&A', labels: lab, values: a.map(r => -r.cost) }], options: { barDir: 'col', barGrouping: 'clustered', chartColors: [NAVY, BRASS, GREY] } },
    { type: pres.charts.LINE, data: [{ name: 'Cumulative net cash flow', labels: lab, values: a.map(r => r.cum) }], options: { chartColors: [RED], lineSize: 3, lineDataSymbolSize: 7 } }],
    { x: 0.6, y: 1.4, w: 8.2, h: 5.0, showLegend: true, legendPos: 'b', legendFontSize: 11, valAxisLabelFontSize: 11, catAxisLabelFontSize: 11, valGridLine: { color: MID, size: 0.5 }, catGridLine: { style: 'none' }, showTitle: true, title: 'Project cash flow by project year (EGP million)', titleFontSize: 13, titleColor: INK });
  stat(s, 9.1, 1.5, 3.6, fmt(D.base.net, 0), 'Jalour net cash flow, nominal, EGP million');
  stat(s, 9.1, 3.2, 3.6, fmt(D.base.npv, 0), 'NPV at 14%, EGP million');
  stat(s, 9.1, 4.9, 3.6, fmt(F.model_peak, 1), `Peak cumulative shortage, EGP million (month ${D.base.peak_m})`);
  source(s, 'Jalour financial model; Financial Annex, Project_CF sheet. Project years: Y1 = months 1 to 12. Excludes the 20% of units retained by Jalour.', 6.6); }
// 10b Combined cash position (125M packs only)
if (D.combined) { const C = D.combined, O = C.other; const s = content(`Combined cash position: ${NAME} and ${O.name}`);
  const n = 28; const lab = C.months.slice(0, n).map(m => 'M' + m);
  s.addChart(pres.charts.LINE, [{ name: NAME, labels: lab, values: C.self_cum.slice(0, n) }, { name: O.name, labels: lab, values: C.other_cum.slice(0, n) }, { name: 'Combined', labels: lab, values: C.comb_cum.slice(0, n) }], { x: 0.6, y: 1.35, w: 6.4, h: 3.9, chartColors: [NAVY, BRASS, RED], lineSize: 2.5, lineDataSymbol: 'none', showLegend: true, legendPos: 'b', legendFontSize: 11, valAxisLabelFontSize: 10, catAxisLabelFontSize: 9, catAxisLabelFrequency: 3, valGridLine: { color: MID, size: 0.5 }, catGridLine: { style: 'none' }, showTitle: true, title: 'Cumulative net cash flow before investor instruments (EGP million)', titleFontSize: 12, titleColor: INK });
  const ob = C.other_base;
  tbl(s, [['EGP million', NAME, O.name], ['Handover (month)', String(I.delivery), String(O.delivery)], ['Sales plan, 80% of units', fmt(I.total80, 0), fmt(O.total80, 0)], ['Net cash flow', fmt(D.base.net, 0), fmt(ob.net, 0)], ['NPV at 14%', fmt(D.base.npv, 0), fmt(ob.npv, 0)], ['Peak shortage (month)', fmt(-D.base.peak, 1) + ' (' + D.base.peak_m + ')', fmt(-ob.peak, 1) + ' (' + ob.peak_m + ')']], 7.3, 1.4, 5.4, [2.4, 1.5, 1.5], 11, { rowH: 0.32 });
  const sc = C.scen; const nm = ['Base case', 'Delivery payments at handover', 'Collections slip 2 quarters', 'Downside: sales +12 months, prices -10%'];
  const labs = ['Base', 'Delivery payments at handover', 'Collections slip 2 quarters', 'Sales +12 months, prices -10%'];
  tbl(s, [['Peak shortage (EGP m)', NAME.split(' ')[0], O.name.split(' ')[0], 'Combined']].concat(nm.map((k, i) => [labs[i], fmt(-sc[k].self_peak, 0), fmt(-sc[k].other_peak, 0), fmt(-sc[k].comb_peak, 0)])), 7.3, 3.55, 5.4, [2.4, 1.0, 1.0, 1.0], 11, { rowH: 0.32 });
  bullets(s, [`Base case: the combined position is lowest at EGP ${fmt(-C.comb_min, 1)} million in month ${C.comb_min_m} and turns positive in month ${C.comb_pos_month}.`, 'In the stress cases both projects are short at the same time, so one project\'s cash cannot be assumed to fund the other.', 'Shown: these two projects only, before any investor instrument. Jalour\'s cash reserves and other projects are not included. No cross-collateral or cross-default between the projects.'], 0.6, 5.5, 12.1, 1.1, 12);
  source(s, 'Jalour financial model; Financial Annex, Combined_Position sheet. Projections, not guarantees.', 6.65); }
// 11 Sources and uses
{ const s = content('Sources and uses of funds');
  const u = F.uses;
  tbl(s, [['Uses', 'EGP million'], ['Down payment to landowner', fmt(u.dp, 1)], ['Guarantee instalments and working capital to the funding peak', fmt(u.guar_wc, 1)], ['Design, permits and pre-construction (assumption)', fmt(u.precon, 1)], ['Total uses at peak', fmt(F.total, 1)]], 0.6, 1.5, 6.0, [4.3, 1.7], 13, { rowH: 0.55 });
  const srcRows = [['Sources', 'EGP million', '% of uses'], ['Investor ticket' + (F.excess_over_uses > 0 ? ' (applied to uses)' : ''), fmt(F.applied, 1), pct(F.share_uses, 0)]];
  if (F.excess_over_uses > 0) srcRows.push(['Ticket above modelled uses (general liquidity)', fmt(F.excess_over_uses, 1), '']);
  srcRows.push(['Balance of project funding', fmt(F.balance, 1), pct(F.balance / F.total, 0)]); srcRows.push(['Total sources', fmt(F.applied + F.balance + F.excess_over_uses, 1), F.excess_over_uses > 0 ? '' : '100%']);
  tbl(s, srcRows, 6.9, 1.5, 5.8, [3.4, 1.2, 1.2], 13, { rowH: 0.55 });
  const lines = [bal, `Modelled peak cumulative shortage EGP ${fmt(F.model_peak, 1)} million; the ticket is ${pct(F.share_peak, 0)} of it.`];
  if (F.excess_over_peak > 0) lines.push(`The ticket exceeds the modelled peak by EGP ${fmt(F.excess_over_peak, 1)} million. The excess is general project liquidity: it absorbs the timing risk on delivery payments (peak shortage EGP ${fmt(F.stress_peak, 0)} million if they arrive at handover). It is not held in a restricted account.`);
  else lines.push(`If delivery payments arrive at handover the peak shortage is EGP ${fmt(F.stress_peak, 0)} million; if sales slip 12 months with prices 10% lower it is EGP ${fmt(F.downside_peak, 0)} million. Jalour will cover gaps above project cash from its cash reserves and cash from other projects it is constructing.`);
  lines.push('Nothing in this structure is ring-fenced: the ticket is project funding.');
  bullets(s, lines, 0.6, 4.6, 12.1, 2.0, 13);
  source(s, 'Financial Annex, Sources_Uses sheet. Pre-construction cost is an assumption pending Jalour\'s budget.', 6.65); }
// 12 Options side by side
{ const s = content('Option A (units) and Option B (cash) side by side');
  tbl(s, [['', 'Option A: units', 'Option B: cash'],
    ['Entitlement', `Units with a face value of EGP ${fmt(A.face, 0)} million (2.0x) at launch list`, `EGP ${fmt(B.mult * T, 0)} million (${mx(B.mult, 1)}) in cash`],
    ['Timing', `Allocation at signing; handover month ${I.delivery}; resale from month ${I.delivery + 3}`, `Grace 24 months; 12 quarterly instalments, months ${B.pay_months[0]} to ${B.pay_months[11]}`],
    ['Price exposure', `Yes: value moves with Mostakbal City prices (delivery list ${fmt(A.appr, 2)}x launch)`, 'No unit price exposure; fixed instalments'],
    ['Projected IRR, base', pct(A.inv.Base.irr), pct(S1.Base.irr)],
    ['Projected multiple, base', mx(A.inv.Base.moic), mx(S1.Base.moic)],
    ['Upside IRR', pct(A.inv.Upside.irr) + ` (${mx(A.inv.Upside.moic)})`, 'Contractual: ' + pct(S1.Upside.irr) + ' if paid'],
    ['Downside IRR', pct(A.inv.Downside.irr) + ` (${mx(A.inv.Downside.moic)})`, 'Contractual: ' + pct(S1.Downside.irr) + ' if paid; depends on Jalour funding'],
    ['Payback', `Month ${D.A_payback || 'n/a'}`, `Month ${D.B_payback || 'n/a'}`],
    ['Main risk', 'Resale price and timing; consents to assign', 'Payment capacity during the construction peak']], 0.6, 1.5, 12.1, [2.2, 5.0, 4.9], 13, { left: true, boldFirst: true, rowH: 0.5 });
  source(s, 'Financial Annex, Option_A and Option_B sheets. Projections; downside = sales delayed 12 months and prices 10% lower.', 6.6); }
// 13 Security
{ const s = content('Security package and investor protections');
  card(s, 0.6, 1.5, 5.9, 4.9); s.addText('Proposed base package', { x: 0.8, y: 1.6, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: NAVY, margin: 0, isTextBox: true });
  bullets(s, ['Option A: a registered allocation of specific units from the first two sales tranches; the units are the investor\'s asset', 'Option B: earmark of specific unsold units with a list value of at least the ticket, with a right to substitute equal value', 'Payment priority covenant: no shareholder distributions or shareholder-loan repayments while an instalment is overdue', 'Negative pledge on the earmarked units', 'Information rights: monthly reporting and audit rights', 'Sponsor funding undertaking for project gaps from Jalour cash reserves and other projects\' cash; no guarantee of return'], 0.8, 2.1, 5.5, 4.2, 13);
  card(s, 6.8, 1.5, 5.9, 4.9, 'FBF3E4'); s.addText('Optional and for counsel', { x: 7.0, y: 1.6, w: 5.5, h: 0.4, fontSize: 16, bold: true, color: BRASS, margin: 0, isTextBox: true });
  bullets(s, ['Promissory note or post-dated cheques for each Option B instalment (common practice in Egypt)', 'Independent engineer reporting on construction progress', 'Step-in right to market the earmarked units after a payment default', 'Landlord acknowledgement of the unit allocation', 'Parent guarantee: not offered in this structure'], 7.0, 2.1, 5.5, 4.2, 13);
  source(s, 'Proposal for review by Jalour\'s lawyers. Landlord consent is required for any allocation or encumbrance affecting its 35% share.', 6.6); }
// 14 Returns and sensitivity
{ const s = content('Returns and sensitivity');
  s.addChart(pres.charts.BAR, [{ name: 'Investor IRR', labels: ['Option A base', 'Option A downside', 'Option A upside', 'Option B (contractual)'], values: [A.inv.Base.irr * 100, A.inv.Downside.irr * 100, A.inv.Upside.irr * 100, S1.Base.irr * 100] }], { x: 0.6, y: 1.4, w: 5.9, h: 3.6, barDir: 'col', chartColors: [NAVY], showValue: true, dataLabelFormatCode: '0.0"%"', dataLabelFontSize: 11, dataLabelColor: INK, dataLabelPosition: 'outEnd', valAxisLabelFontSize: 11, catAxisLabelFontSize: 10, valGridLine: { color: MID, size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, showTitle: true, title: 'Projected investor IRR (%)', titleFontSize: 13, titleColor: INK });
  const sp = D.sens; const row = (lab, o) => [lab, fmt(o.npv, 0), fmt(-o.peak, 0)];
  tbl(s, [['Project case', 'NPV (EGP m)', 'Peak shortage (EGP m)'], row('Base', sp.price['1.0']), row('Prices -10%', sp.price['0.9']), row('Prices +10%', sp.price['1.1']), row('Sales delay 12 months', sp.delay['4']), row('Cost overrun +10%', sp.cost['1.1']), row('Collections slip 2 quarters', sp.slip['2']), ['Delivery payments at handover', fmt(D.proj['Delivery payments at handover'].npv, 0), fmt(-D.proj['Delivery payments at handover'].peak, 0)]], 6.8, 1.5, 5.9, [3.0, 1.35, 1.55], 12, { rowH: 0.42 });
  bullets(s, [`Option A break-even: the units could be resold at ${pct(D.breakeven.A_price_factor_moic1, 0)} of the delivery list price before the ticket is lost.`, `Option B: cash available covers each instalment at least ${mx(S1.Base.liq_min, 1)} on a pooled-cash basis; net cash flow alone covers it in ${S1.Base.n - S1.Base.strict_below} of ${S1.Base.n} quarters because of the construction peak.`, 'Returns are projections, not guarantees.'], 0.6, 5.1, 12.1, 1.5, 12);
  source(s, 'Financial Annex, Scenarios, Sensitivity and Option sheets. Project cases show Jalour net cash flow before any investor instrument.', 6.65); }
// 15 Risks
{ const s = content('Risks and mitigants');
  tbl(s, [['Risk', 'Mitigant'],
    ['Market and price: sales plan assumes list-price rises of ' + fmt(Math.min(D.prices.final[1] / D.prices.launch[1], D.prices.final[0] / D.prices.launch[0]) * 100 - 100, 0) + '% to ' + fmt(Math.max(D.prices.final[1] / D.prices.launch[1], D.prices.final[0] / D.prices.launch[0]) * 100 - 100, 0) + '% over the sales window', 'Downside case shown; Option B has no unit price exposure; price evidence in the data room'],
    ['Delivery payments collected ahead of handover', 'Contractual milestone; Jalour covers any gap from cash reserves and other projects\' cash; stress case shown'],
    ['Construction cost and inflation (costs flat in the model)', 'Cost +10% case; fixed-price contracting to be sought; contingency 5% in the model'],
    ['Landlord: consent and minimum guarantee payments', 'Written consent as a condition precedent; guarantee schedule fully funded in the plan'],
    ['Delivery and permits', 'Permit status in the data room; independent engineer (optional)'],
    ['Legal and regulatory: private placement, off-plan rules, assignment of contracts', 'Counsel review before signing; developer and landlord consent mechanics defined'],
    ['Currency and FX: EGP returns; USD value may fall', 'USD equivalents shown at EGP 48; foreign investors take FX advice'],
    ['Concentration: single project', 'No cross-collateral or cross-default with other Jalour projects; investor exposed to this project only']], 0.6, 1.45, 12.1, [5.4, 6.7], 12, { left: true, rowH: 0.52 });
  source(s, 'Jalour; full risk factors in the Investment Memorandum and the Due Diligence Memorandum.', 6.65); }
// 16 Timeline
{ const s = content('Timeline and milestones');
  const ms = [[I.sales_start, 'Sales start'], [I.con_start, 'Construction start'], [B.pay_months[0] - 3 - 24 + 24, 'Option B grace ends'], [I.delivery, 'Handover'], [I.delivery + 3, 'First unit resales'], [B.pay_months[11], 'Last Option B payment']];
  const x0 = 0.9, x1 = 12.4, mmax = Math.max(B.pay_months[11], I.delivery + 12) + 3; const X = m => x0 + (x1 - x0) * m / mmax;
  s.addShape(pres.shapes.LINE, { x: x0, y: 3.4, w: x1 - x0, h: 0, line: { color: NAVY, width: 2 } });
  const all = [[0, 'Model start'], [D.info.sales_start, 'Sales start'], [I.con_start, 'Construction start'], [D.B.pay_months[0] - 3, 'Grace ends'], [B.pay_months[0], 'First Option B payment'], [I.delivery, 'Handover'], [I.delivery + 3, 'First resales'], [B.pay_months[11], 'Last Option B payment']];
  all.sort((a, b) => a[0] - b[0]); const seen = new Set();
  all.forEach((m, i) => { const up = i % 2 === 0; const x = X(m[0]);
    s.addShape(pres.shapes.OVAL, { x: x - 0.09, y: 3.31, w: 0.18, h: 0.18, fill: { color: BRASS }, line: { color: WHITE, width: 1 } });
    s.addText([{ text: m[1], options: { bold: true, breakLine: true } }, { text: 'Month ' + m[0] }], { x: x - 0.85, y: up ? 2.3 : 3.65, w: 1.7, h: 0.9, fontSize: 11, color: INK, align: 'center', valign: up ? 'bottom' : 'top', margin: 0, isTextBox: true }); });
  bullets(s, [`Funding and signing: month ${D.info.sales_start > 0 ? (D.k === 'GS' ? 3 : 12) : 0} (ticket due when the down payment of EGP ${fmt(I.dp, 0)} million falls due).`, `Option A: units are allocated at signing and delivered at handover (month ${I.delivery}); the investor may resell after consents.`, `Option B: 24-month grace, then 12 quarterly instalments from month ${B.pay_months[0]}.`], 0.6, 4.9, 12.1, 1.6, 13);
  source(s, 'Jalour financial model; Financial Annex. Months count from the start of the model timeline.', 6.65); }
// 17 Ask
{ const s = pres.addSlide({ masterName: 'DARK' }); sn++;
  s.addText('The ask and next steps', { x: 0.8, y: 0.7, w: 11.5, h: 0.9, fontFace: 'Cambria', fontSize: 34, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText(`Jalour invites an investment of EGP ${T} million in ${NAME}, with the choice of Option A (units, 2.0x face value) or Option B (cash, ${mx(B.mult, 1)} over 3 years after a 24-month grace period).`, { x: 0.8, y: 1.8, w: 11.5, h: 1.2, fontSize: 18, color: 'D5DAE1', margin: 0, isTextBox: true });
  const steps = ['Review this presentation, the Investment Memorandum and the Financial Annex', 'Select Option A or Option B and confirm the ticket', 'Due diligence: data room, legal and technical review', 'Landlord consent and definitive documents', 'Signing and funding'];
  steps.forEach((t, i) => { const y = 3.3 + i * 0.62; s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.42, h: 0.42, fill: { color: BRASS }, line: { color: BRASS, width: 0 } }); s.addText(String(i + 1), { x: 0.8, y, w: 0.42, h: 0.42, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: WHITE, margin: 0, isTextBox: true }); s.addText(t, { x: 1.45, y, w: 10.8, h: 0.42, fontSize: 16, color: WHITE, valign: 'middle', margin: 0, isTextBox: true }); });
  s.addText('Jalour may raise further capital at project or holding level, subject to the investor\'s stated rights. Projections are not guarantees. Strictly private and confidential.', { x: 0.8, y: 6.55, w: 11.5, h: 0.5, fontSize: 11, color: 'AAB2BD', margin: 0, isTextBox: true });
  s.addNotes('Next steps and contact details to be added by Jalour.'); }
pres.writeFile({ fileName: outPath }).then(() => console.log('deck written', sn));
