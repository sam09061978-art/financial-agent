// make-pages-data.js — собирает demo/pages-data.js для интерактивных страниц витрины.
// Источники: demo/register-data.json (контур «утро») и demo-output/bdr-stroykom-01.2027.xlsx
// (матрица БДР + лист «Разнесение» + «Требует решения»). Запуск: node scripts/make-pages-data.js
const XLSX = require('./node_modules/xlsx');
const fs = require('fs');
const ROOT = __dirname + '/../';
const round2 = x => Math.round(x * 100) / 100;

// ==== контур «утро» из register-data.json — flatten позиций по секциям/реестрам ====
const reg = JSON.parse(fs.readFileSync(ROOT + 'demo/register-data.json', 'utf8'));
const rows = [];    // {sec, agent, cont, plan, amount, regLabel}
const byDay = Array(reg.dates.length).fill(0);
const cnt = { 'УТВЕРЖДЕН': 0, 'НА УТВЕРЖДЕНИИ': 0, 'НА ФОРМИРОВАНИИ': 0 };
const sumOf = { 'УТВЕРЖДЕН': 0, 'НА УТВЕРЖДЕНИИ': 0, 'НА ФОРМИРОВАНИИ': 0 };
for (const sec of reg.sections) {
  const walk = items => {
    for (const it of items) {
      if (it.reghead) { walk(it.items); continue; }
      const leaves = it.group ? it.group : [it];
      for (const lf of leaves) {
        const plan = lf.plan == null ? -1 : lf.plan;
        rows.push({ sec: sec.title, agent: it.agent, cont: lf.cont, plan, amount: lf.amount,
          reg: it.reg || it.reghead || '' });
        if (plan >= 0) byDay[plan] = round2(byDay[plan] + lf.amount);
        cnt[sec.title] = round2(cnt[sec.title] + lf.amount);
        sumOf[sec.title] = round2(sumOf[sec.title] + lf.amount);
      }
    }
  };
  walk(sec.items);
}

// ==== БДР: матрица, разнесение, хвост ====
const wb = XLSX.readFile(ROOT + 'demo-output/bdr-stroykom-01.2027.xlsx');
const m = XLSX.utils.sheet_to_json(wb.Sheets['БДР'], { header: 1, defval: '' });
const cfo = m[1].slice(1, -1).filter(v => v !== '');
const cfoIdx = m[1].reduce((a, v, i) => { if (v && v !== 'Статья учета') a[v] = i; return a; }, {});
const articles = [];
let barData = null;
for (let i = 2; i < m.length; i++) {
  const r = m[i];
  const name = String(r[0] || ''); if (!name) continue;
  if (/^Контроль/i.test(name)) break;
  if (/^ИТОГО/i.test(name)) { barData = { name, cells: cfo.map(c => round2(r[cfoIdx[c]] || 0)) }; continue; }
  articles.push({ article: name, cells: cfo.map(c => round2(r[cfoIdx[c]] || 0)), total: round2(r[cfoIdx['ИТОГО']] || 0) });
}
const sp = XLSX.utils.sheet_to_json(wb.Sheets['Разнесение'], { header: 1, defval: '' });
const spread = sp.slice(1).filter(r => r[0] !== '' || r[7]).map(r => ({
  date: String(r[0] || ''), doc: String(r[1] || ''), dt: String(r[2] || ''), sub: String(r[3] || ''),
  dept: String(r[4] || ''), kt: String(r[5] || ''), agent: String(r[6] || ''), amount: round2(r[7] || 0),
  article: String(r[8] || ''), cfo: String(r[9] || ''),
}));
const unres = XLSX.utils.sheet_to_json(wb.Sheets['Требует решения'], { header: 1, defval: '' });
let unresolved = 0;
for (const r of unres) { const v = r[1]; if (typeof v === 'number') unresolved = round2(unresolved + v); }

const out = {
  date: regeTitle(reg.title),
  dates: reg.dates,
  total: reg.expected.total,
  sections: reg.expected,        // section0/1/2
  rows, byDay, cnt, sumOf,
  bdr: { cfo, articles, bar: barData, spread, unresolved,
         total: articles.reduce((a, b) => round2(a + b.total), 0) },
};
function regeTitle(t) { return t.replace(/Перечень неоплаченных платежей на утро /, '').replace(/ — .*/, ''); }
fs.writeFileSync(ROOT + 'demo/pages-data.js',
  '/* Автогенерация scripts/make-pages-data.js — данные для интерактивных страниц витрины */\n' +
  'window.PAGES_DATA = ' + JSON.stringify(out) + ';\n', 'utf8');
console.log('OK: demo/pages-data.js | позиций перечня:', rows.length, '| статей БДР:', articles.length,
  '| проводок в Разнесении:', spread.length, '| не разнесено:', unresolved);
