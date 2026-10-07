// helpers shared by deck builder
const fmt = (x, d = 1) => (x === null || x === undefined || Number.isNaN(x)) ? 'n/a' : Number(x).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (x, d = 1) => (x === null || x === undefined || Number.isNaN(x)) ? 'n/a' : (x * 100).toFixed(d) + '%';
const mx = (x, d = 2) => Number(x).toFixed(d) + 'x';
module.exports = { fmt, pct, mx };
