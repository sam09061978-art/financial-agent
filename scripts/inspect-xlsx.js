// Inspect workbook structure: sheet names, dimensions, first N rows as TSV to stdout.
// Usage: node inspect-xlsx.js <file> [maxRows] [sheetName]
const path = require('path');
const XLSX = require('xlsx');

const file = process.argv[2];
const maxRows = parseInt(process.argv[3] || '40', 10);
const onlySheet = process.argv[4];

const wb = XLSX.readFile(file, { cellDates: false });
for (const name of wb.SheetNames) {
  if (onlySheet && name !== onlySheet) continue;
  const ws = wb.Sheets[name];
  const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  const dim = range ? `rows=${range.e.r - range.s.r + 1} cols=${range.e.c - range.s.c + 1}` : 'EMPTY';
  console.log(`=== SHEET: "${name}" | ${dim} ===`);
  if (!range) continue;
  const data = XLSX.utils.sheet_to_csv(ws, { FS: '\t' });
  const lines = data.split('\n').filter(l => l !== '');
  lines.slice(0, maxRows).forEach(l => console.log(l.replace(/\r$/, '')));
  if (lines.length > maxRows) console.log(`... (${lines.length - maxRows} more rows)`);
}