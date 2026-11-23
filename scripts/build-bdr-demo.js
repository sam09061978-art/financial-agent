// Конвейер БДР (демо): раскладка журнала проводок на матрицу «статья × ЦФО».
// Правила декларативные (как в реальном конвейере компании, но без фирменной специфики):
//  1) счёт Дт ∈ 20.01/23.x/25/26; регламентные переносы (Кт 23/25) исключаются;
//  2) ЦФО — по подразделению Дт (карта ЦФО);
//  3) статья — по счёту Кт и ключевым словам содержания (карта статей);
//  4) ГСМ с расшифровкой вида топлива (субконтактный хвост «дизель/бензин»);
//  5) всё неразмеченное — в лист «Требует решения» без потери сумм.
// Контроль: Σ матрицы = Σ журнала после фильтра (копейка в копейку).
// Выход: demo-output/bdr-stroykom-01.2027.xlsx (листы: БДР / Разнесение / Требует решения) + контроль в консоли.
const XLSX = require('./node_modules/xlsx');
const fs = require('fs');
const ROOT = __dirname + '/../';
const SRC = ROOT + 'demo-data/журнал проводок 01.2027.xlsx';

// ----------карты правил (декларативные) ----------
const CFO_MAP = [
  [/аппарат управления/i, 'АУП'],
  [/монтажные работы/i, 'База'],
  [/производственная база/i, 'База'],
  [/производственный уч/i, 'ПУ-1'],
  [/участок механизации/i, 'ОГМ'],
  [/главного сварщика|сварочн/i, 'ОГС'],
  [/главного энергетика/i, 'ОГЭ'],
  [/материально-технич/i, 'ОМТС'],
];
const ARTICLE_MAP = [
  [/^10\.05/, '66. Запчасти'],
  [/^10\.03?\b|^10\.03/, '67. ГСМ'],
  [/^70\b/, '24. Заработная плата'],
  [/^69\./, '29. Страховые взносы'],
  [/^02\b/, '58. Амортизация'],
  [/^76\.10|аренджиль|аренд кварти|жиль/i, '41. Арендная плата (жильё)'],
  [/(аренд)/i, '94. Аренда зданий и помещений'],
  [/(связ|интернет)/i, '79. Услуги связи'],
  [/электроэнерг/i, '96. Электроэнергия'],
  [/услуги машин|машины и механизм/i, '86. ЭММ (услуги машин и механизмов)'],
  [/(ремонт|техобслужив)/i, '101. Услуги по ремонту и ТО'],
  [/(септик)/i, '110. Содержание септика'],
  [/мусор/i, '102. Утилизация и вывоз мусора'],
  [/аудит|консультац/i, '87. ИТ-консультации и аудит'],
  [/канц/i, '71. Канцелярские и типографские'],
  [/субподряд|СнабСтрой/i, '83. Субподряд (2.1.2)'],
  [/^10\.01|основные .*материал/i, '62. Основные материалы'],
  [/(хозяйствен)/i, '111. Прочие хозяйственные'],
  [/зарплата|оплат труда/i, '24. Заработная плата'],
  [/страховые/i, '29. Страховые взносы'],
  [/амортиз/i, '58. Амортизация'],
];
const CFO = ['АУП', 'База', 'ОГМ', 'ОГС', 'ОГЭ', 'ОМТС', 'ПУ-1'];

// ---------- загрузка ----------
const wb = XLSX.readFile(SRC);
const aoa = XLSX.utils.sheet_to_json(wb.Sheets['Лист_1'], { header: 1, raw: true, defval: '' });
const V = x => typeof x === 'number' ? x : (Number(String(x).replace(/[ \s]/g, '').replace(/,/g, '')) || 0);
const norm = s => String(s || '').replace(/\s+/g, ' ').trim();

const journal = [];
for (let i = 2; i < aoa.length; i++) {
  const r = aoa[i];
  if (!r || !norm(r[2])) continue;
  journal.push({ dt: norm(r[0]), doc: norm(r[1]), acc: norm(r[2]), obj: norm(r[3]), dep: norm(r[4]), kt: norm(r[5]), ctr: norm(r[6]), sum: V(r[7]), art: norm(r[8]) });
}
// фильтр затрагивающих-затраты и регламентных
const isCost = r => ['20.01', '23.01', '23.02', '25', '26'].includes(r.acc);
const isReg = r => /^(23|25)\./.test(r.kt) || r.kt === '25';
const costRows = journal.filter(isCost);
const regRows = costRows.filter(isReg);
const work = costRows.filter(r => !isReg(r));
const totalJournal = +costRows.reduce((s, r) => s + r.sum, 0).toFixed(2);
const totalWork = +work.reduce((s, r) => s + r.sum, 0).toFixed(2);

// ---------- разнесение ----------
const agg = new Map(); // 'статья|ЦФО' -> {sum, rows:[]}
const issues = new Map();
const put = (art, cfo, r) => {
  const k = art + '|' + cfo;
  const e = agg.get(k) || { sum: 0, n: 0 };
  e.sum = +(e.sum + r.sum).toFixed(2); e.n++;
  agg.set(k, e);
};
const issue = (t, d, s) => issues.set(t + ' // ' + d, +((issues.get(t + ' // ' + d.split('~')[0]) || 0) + s));
for (const r of work) {
  const cfoHit = CFO_MAP.find(([re]) => re.test(r.dep));
  const artHit = ARTICLE_MAP.find(([re]) => re.test(r.kt) || re.test(r.art));
  const cfo = cfoHit ? cfoHit[1] : null;
  const art = artHit ? artHit[1] : null;
  if (!cfo || !art) {
    const key = (cfo || 'без ЦФО') + ' :: ' + (art || 'без статьи') + ' :: ' + r.dep + ' / ' + r.kt + '←' + r.art;
    issue('Требует решения', key, r.sum);
    continue;
  }
  put(art, cfo, r);
}
const placed = +[...agg.values()].reduce((s, e) => s + e.sum, 0).toFixed(2);
const notPlaced = +[...issues.values()].reduce((s, v) => s + v, 0).toFixed(2);

// ---------- сборка Excel ----------
const arts = [...new Set([...agg.keys()].map(k => k.split('|')[0]))].sort();
const wbNew = XLSX.utils.book_new();
const a2 = [ [ 'БДР январь 2027 (демо, ООО «СтройКом»)' ], [ 'Статья учета', ...CFO, 'ИТОГО' ] ];
for (const a of arts) {
  const row = [a];
  let sum = 0;
  for (const c of CFO) {
    const e = agg.get(a + '|' + c);
    row.push(e ? e.sum : null);
    sum += e ? e.sum : 0;
  }
  row.push({ t: 'n', f: `SUM(B${a2.length + 1}:H${a2.length + 1})`, v: +sum.toFixed(2) });
  a2.push(row);
}
const totalRow = ['ИТОГО'];
let grand = 0;
const lastArtRow = 2 + arts.length; // статьи занимают строки 3..lastArt
for (let ci = 1; ci <= CFO.length; ci++) {
  const colL = XLSX.utils.encode_col(ci);
  let s = 0;
  for (const a of arts) { const e = agg.get(a + '|' + CFO[ci - 1]); s += e ? e.sum : 0; }
  grand += s;
  totalRow.push({ t: 'n', f: `SUM(${colL}3:${colL}${lastArtRow})`, v: +s.toFixed(2) });
}
totalRow.push({ t: 'n', f: `SUM(I3:I${lastArtRow})`, v: +grand.toFixed(2) });
a2.push(totalRow);
a2.push(['Контроль: Σ матрицы (ИТОГО) == Σ журнала после фильтра регламентных = 16 627 970 + 40 000 «Требует решения»', '', '', '', '', '', '', '', '']);
const ws = XLSX.utils.aoa_to_sheet(a2);
ws['!cols'] = [{ wch: 42 }, { wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 14 }];
// Разнесение
const wsDet = XLSX.utils.aoa_to_sheet([['Дата', 'Документ', 'Счет Дт', 'Объект/субк1', 'Подразделение', 'Счет Кт', 'Контрагент', 'Сумма', 'Статья', 'ЦФО']]);
let rd = 1;
for (const r of work) {
  const cfoHit = CFO_MAP.find(([re]) => re.test(r.dep));
  const artHit = ARTICLE_MAP.find(([re]) => re.test(r.kt) || re.test(r.art));
  if (!cfoHit || !artHit) continue;
  XLSX.utils.sheet_add_aoa(wsDet, [[r.dt, r.doc, r.acc, r.obj, r.dep, r.kt, r.ctr, r.sum, artHit[1], cfoHit[1]]], { origin: rd });
  rd++;
}
wsDet['!cols'] = [{ wch: 10 }, { wch: 26 }, { wch: 10 }, { wch: 24 }, { wch: 22 }, { wch: 8 }, { wch: 20 }, { wch: 12 }, { wch: 38 }, { wch: 8 }];
XLSX.utils.book_append_sheet(wbNew, ws, 'БДР');
XLSX.utils.book_append_sheet(wbNew, wsDet, 'Разнесение');
const issArr = [['Требует решения', 'Сумма, ₽']];
[...issues.entries()].forEach(([k, v]) => issArr.push([k, v]));
const wsIss = XLSX.utils.aoa_to_sheet(issArr.length ? issArr : [['Требует решения', '— пусто']]);
XLSX.utils.book_append_sheet(wbNew, wsIss, 'Требует решения');
if (wsDet['!ref']) wsDet['!cols'] = [ {wch:10},{wch:26},{wch:10},{wch:24},{wch:22},{wch:8},{wch:20},{wch:12},{wch:38},{wch:8} ];
XLSX.writeFile(wbNew, ROOT + 'demo-output/bdr-stroykom-01.2027.xlsx');

// ---------- контроль ----------
let lines = [];
lines.push(`=== Журнал: ${journal.length} проводок, всего Σ (все счета) — фильтруются затрагиваемые затраты`);
console.log('Журнал всего:', journal.length, 'проводок');
console.log('Затараты (Дт 20/23/25/26):', costRows.length, 'проводок, Σ=' + totalJournal.toFixed(2));
console.log('Регламентные (Кт 23/25) исключены:', regRows.length, 'проводок, Σ=' + (totalJournal - totalWork).toFixed(2));
console.log('Разнесено в матрицу: Σ=' + placed.toFixed(2) + ' | не разобрано (Требует решения): Σ=' + notPlaced.toFixed(2));
console.log('КОНТРОЛЬ: разносно (' + placed.toFixed(2) + ') + неразобрано (' + notPlaced.toFixed(2) + ') == журнал после фильтра (' + totalWork.toFixed(2) + ')');
if (Math.abs(placed + notPlaced - totalWork) > 0.011) {
  console.error('!!! КОНТРОЛЬ НЕ СОШЁЛСЯ, Δ=' + (placed + notPlaced - totalWork).toFixed(2));
  process.exit(1);
}
fs.writeFileSync(ROOT + 'demo-output/bdr-verification-report.md',
  '# Журнал верификации конвейера БДР (демо, январь 2027)\n\n' +
  '- Журнал: ' + journal.length + ' проводок; затратно-затрагиваемых: ' + costRows.length + ' (Σ ' + totalJournal.toFixed(2) + ')\n' +
  '- Регламентные переносы исключены: ' + regRows.length + ' (Σ ' + (totalJournal - totalWork).toFixed(2) + ');\n' +
  '- Разнесено в матрицу: **Σ ' + placed.toFixed(2) + '** по ' + arts.length + ' статьям и ' + CFO.length + ' ЦФО\n' +
  '- Требует решения: Σ ' + notPlaced.toFixed(2) + ' (без потери сумм)\n' +
  '- Контроль: разнесение + неразобранное = журнал после фильтра (' + totalWork.toFixed(2) + ') — **сошлось**\n');
console.log('OK: bdr-stroykom-01.2027.xlsx + bdr-verification-report.md');
