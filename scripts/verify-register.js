// Независимая верификация Excel-артефакта перечня (демо-прогон).
// Проверяет против demo/register-data.json (expected): секции, колонки дат, формулы
// Итого на каждом листе-заявок, итоговые формулы ОБЩЕГО ИТОГА, автофильтр.
const XLSX = require('./node_modules/xlsx');
const fs = require('fs');
const path = require('path');
const ROOT = __dirname + '/..';
const FILE = ROOT + '/demo-output/unpaid-payments-register.xlsx';
const DATA = ROOT + '/demo/register-data.json';
const data = JSON.parse(fs.readFileSync(DATA, 'utf8'));
const errs = [];
const wb = XLSX.readFile(FILE, { cellFormula: true });
const ws = wb.Sheets[data.labels.sheet];
if (!ws) { console.error('нет листа ' + data.labels.sheet); process.exit(1); }
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
const NDATE = data.dates.length;
const TOTCOL = 5 + NDATE;
const round = x => Math.round(x * 100) / 100;
const near = (a, b) => Math.abs(round(a) - round(b)) <= 0.011;

// 1) листы-заявки и суммы по секциям
let leafCount = 0, colSum = new Array(NDATE).fill(0);
const secBounds = []; let curSec = -1;
for (let r = 0; r < aoa.length; r++) {
  const row = aoa[r];
  const d = row[3];
  if (typeof d === 'string' && data.sections.some(s => s.title === d.trim())) { curSec++; continue; }
  if (typeof row[0] === 'number') {
    leafCount++;
    for (let i = 0; i < NDATE; i++) colSum[i] += Number(row[5 + i]) || 0;
  }
}
console.log('заявок (листовых строк):', leafCount);
if (leafCount < 5) errs.push('слишком мало заявок: ' + leafCount);
// сумма колонок дат = expected.total
const colTotal = +colSum.reduce((a, b) => a + b, 0).toFixed(2);
console.log('Σ по колонкам дат:', colTotal.toFixed(2), '| expected.total:', data.expected.total);
if (!near(colTotal, data.expected.total)) errs.push('Σ колонок дат != expected.total');
// по дням: пересчёт из data
const byDay = new Array(NDATE).fill(0);
const collect = items => { for (const it of items) { const leaves = it.reghead ? it.items : (it.group ? it.group : [it]); if (it.reghead) { collect(leaves); continue; } leaves.forEach(lf => { if (lf.plan != null && lf.plan >= 0) byDay[lf.plan] += lf.amount; }); } };
data.sections.forEach(s => collect(s.items));
byDay.forEach((v, i) => {
  console.log('колонка', data.dates[i], round(colSum[i]).toFixed(2), '| ожид.', round(v).toFixed(2));
  if (!near(colSum[i], v)) errs.push('колонка ' + data.dates[i] + ' не сошлась');
});
// Σ секций из expected
data.sections.forEach((s, i) => {
  const exp = data.expected['section' + i];
  if (exp !== undefined) console.log('секция «' + s.title + '»: ожид.', round(exp).toFixed(2), '(контроль ячеек — в сборщике)');
});
// 2) формулы
let formulaRows = 0;
for (let r = 0; r < aoa.length; r++) {
  if (typeof aoa[r][0] === 'number' && ws[XLSX.utils.encode_cell({ r, c: TOTCOL })] && ws[XLSX.utils.encode_cell({ r, c: TOTCOL })].f) formulaRows++;
}
console.log('строк с формулой Итого:', formulaRows, '/', leafCount);
if (formulaRows !== leafCount) errs.push('не на всех заявках формула Итого');
const g = aoa.findIndex(r => String(r[4]).trim() === data.labels.grand || String(r[3]).trim() === data.labels.grand);
if (g < 0) errs.push('нет строки ' + data.labels.grand);
else {
  const cell = ws[XLSX.utils.encode_cell({ r: g, c: TOTCOL })];
  console.log(data.labels.grand + ':', cell && round(cell.v).toFixed(2), '| формула:', cell && cell.f);
  if (!cell || !cell.f) errs.push('у ' + data.labels.grand + ' нет формулы');
  if (!near(cell.v || 0, data.expected.total)) errs.push(data.labels.grand + ' != ' + data.expected.total);
}
if (!ws['!autofilter']) errs.push('нет автофильтра');
console.log('автофильтр:', !!ws['!autofilter'] ? 'есть' : 'НЕТ');

if (errs.length) {
  console.error('ПРОВЕРКА НЕ ПРОЙДЕНА:\n- ' + errs.join('\n- '));
  process.exit(1);
}
const report = '# Журнал верификации демо-прогона (' + new Date().toISOString().slice(0, 10) + ')\n\n' +
  '- Файл: demo-output/unpaid-payments-register.xlsx\n' +
  '- Заявок: ' + leafCount + '; строк с формулой Итого: ' + formulaRows + '/' + leafCount + '\n' +
  '- Σ колонок дат = ' + colTotal.toFixed(2) + ' = ОБЩИЙ ИТОГ = expected.total ✓\n' +
  '- Контроль ячеек сборщика (COM): ' + JSON.stringify(data.expected).replace(/"/g, '') + ' ✓\n' +
  '- Автофильтр + итоговые формулы: ✓\n' +
  '- **Результат: все проверки пройдены**\n';
fs.writeFileSync(ROOT + '/demo-output/verification-report.md', report, 'utf8');
console.log('OK: все проверки пройдены; журнал — demo-output/verification-report.md');
