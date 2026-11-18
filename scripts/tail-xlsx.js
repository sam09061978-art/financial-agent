// Print last N rows of the first sheet as TSV (skipping header).
// Usage: node tail-xlsx.js <file> [lastN]
const XLSX = require('xlsx');
const file = process.argv[2];
const lastN = parseInt(process.argv[3] || '30', 10);
const wb = XLSX.readFile(file);
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
const data = aoa.filter(r => r.some(c => String(c).trim() !== ''));
console.log('HEADER: ' + aoa[0].map(c => String(c)).join(' | '));
data.slice(-lastN).forEach(r =>
  console.log(r.map(c => String(c).replace(/\t/g, ' ')).join(' | ')));