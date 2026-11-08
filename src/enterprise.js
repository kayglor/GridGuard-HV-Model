// GridGuard HV Services: assumptions and calculated results (all currency in generic currency units, CU).
const SITES = [5, 14, 22];
const SERVICES = [
  // name, unit, units per site per year, price per unit, person-hours per unit, note
  ['Transformer annual PM', 'per transformer', 10, 1500, 10, 'IR/PI, winding R, TTR, oil tests, earth (reference estate: 10 units)'],
  ['Cable test (sheath, IR, VLF)', 'per circuit', 10, 1200, 8, 'One test visit per circuit per year (reference estate: 10 circuits)'],
  ['Switchgear annual shutdown', 'per board', 1, 10000, 160, 'Clean, torque, contact R, CB timing, relay tests'],
  ['Thermography and PD survey', 'per survey', 4, 3000, 40, 'Quarterly surveys of board, cables and transformers'],
  ['Corrective call-out', 'per call-out', 3, 4000, 40, 'Fault response, location, repair support, RCA'],
  ['Monitoring and reporting subscription', 'per month', 12, 800, 4, 'Dashboard, health index, monthly KPI report'],
  ['CMMS setup and data digitising', 'per new site', 1, 15000, 120, 'One-off at contract start (billed on new sites only)'],
];
const ESC = 0.03, DIRECT = 0.15, ONCOST = 0.20, HOURS = 1600, UTIL = 0.70, TAX = 0.25, LIFE = 5;
const TEAM = [
  ['Operations manager', 'N', 4500, [1, 1, 1]], ['HV test engineer', 'Y', 3800, [1, 2, 3]], ['Cable test engineer', 'Y', 3600, [1, 2, 2]],
  ['Protection engineer', 'Y', 4000, [0, 1, 2]], ['HV technician', 'Y', 2600, [2, 4, 6]], ['Thermographer / analyst', 'Y', 3200, [0, 1, 2]],
  ['Planner (CMMS)', 'N', 2800, [1, 1, 2]], ['HSE officer', 'N', 3000, [0, 1, 1]], ['Admin and finance', 'N', 2200, [1, 1, 2]],
];
const OVERHEAD = [['Equipment purchase (capex)', [180000, 60000, 80000]], ['Software licences', [24000, 36000, 54000]], ['Insurance', [18000, 24000, 36000]], ['Vehicles and travel', [30000, 48000, 78000]], ['Office and admin', [24000, 30000, 40000]], ['Training and calibration', [20000, 28000, 40000]], ['Marketing and bids', [15000, 20000, 25000]]];

const calc = () => {
  const out = { years: [] };
  let cum = 0, minCum = 0;
  for (let y = 0; y < 3; y++) {
    const sites = SITES[y], newSites = y === 0 ? SITES[0] : SITES[y] - SITES[y - 1];
    const esc = Math.pow(1 + ESC, y);
    const rev = SERVICES.map((s, i) => (i === 6 ? newSites : sites) * s[2] * s[3] * esc);
    const revenue = rev.reduce((a, b) => a + b, 0);
    const direct = revenue * DIRECT, gross = revenue - direct;
    const payroll = TEAM.reduce((a, t) => a + t[2] * t[3][y], 0) * 12 * (1 + ONCOST);
    const ov = OVERHEAD.slice(1).map(o => o[1][y]); const overheads = ov.reduce((a, b) => a + b, 0);
    const capex = OVERHEAD[0][1][y];
    const dep = OVERHEAD[0][1].slice(0, y + 1).reduce((a, b) => a + b, 0) / LIFE;
    const ebitda = gross - payroll - overheads, ebit = ebitda - dep, tax = Math.max(0, ebit) * TAX, net = ebit - tax;
    const hours = SERVICES.reduce((a, s, i) => a + (i === 6 ? newSites : sites) * s[2] * s[4], 0);
    const fte = TEAM.reduce((a, t) => a + (t[1] === 'Y' ? t[3][y] : 0), 0);
    const avail = fte * HOURS * UTIL;
    const opcash = net + dep, ncf = opcash - capex; cum += ncf; minCum = Math.min(minCum, cum);
    out.years.push({ sites, newSites, rev, revenue, direct, gross, payroll, ov, overheads, capex, dep, ebitda, ebit, tax, net, hours, fte, avail, util: hours / avail, opcash, ncf, cum });
  }
  out.peak = -minCum;
  return out;
};
module.exports = { SITES, SERVICES, ESC, DIRECT, ONCOST, HOURS, UTIL, TAX, LIFE, TEAM, OVERHEAD, calc };
