// Receives quiz attempts from the website, checks the answers on the server,
// issues one prize code per phone number and logs every attempt in the sheet.
const SHEET_ID = '18Tuay5d2a2jqMxgYmXAsyjiJRi0TJMbSEhT-6F_H8sA';
const SHEET_NAME = 'Відповіді';
const HEADERS = ['Дата', "Ім'я", 'Телефон', 'Заклад', 'Мова', 'Бали', 'Виграв', 'Код купона', 'Спроба', 'Згода', 'Приз видано'];
const CORRECT = [1, 0, 2, 2, 0, 1, 1, 1, 2, 1];
const VENUES = {
  'sb-kat': 'SushiBang Resto.Bar Katowice',
  'lab-poz': 'Labirynt Lounge Bar Poznań',
  'lab-wro': 'Labirynt Resto.Bar Wrocław'
};
const LANGS = ['pl', 'uk', 'en'];
const MAX_ATTEMPTS_PER_HOUR = 30;

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return reply({ ok: false, error: 'bad_request' }); }

  const name = String(d.name || '').trim().slice(0, 60);
  const phone = normPhone(d.phone);
  const venue = VENUES[d.venueId];
  const lang = LANGS.indexOf(d.lang) >= 0 ? d.lang : 'pl';
  const answers = Array.isArray(d.answers) ? d.answers.slice(0, CORRECT.length).map(Number) : [];
  if (!name || !phone || !venue || d.consent !== true || answers.length !== CORRECT.length) {
    return reply({ ok: false, error: 'invalid' });
  }

  const cache = CacheService.getScriptCache();
  const key = 'n_' + phone;
  const tries = Number(cache.get(key) || 0) + 1;
  if (tries > MAX_ATTEMPTS_PER_HOUR) return reply({ ok: false, error: 'too_many' });
  cache.put(key, String(tries), 3600);

  const score = answers.reduce((s, a, i) => s + (a === CORRECT[i] ? 1 : 0), 0);
  const won = score === CORRECT.length;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet();
    let code = '';
    let repeat = false;
    if (won) {
      const existing = findPrize(sh, phone);
      if (existing) { code = existing; repeat = true; }
      else code = newCode(sh);
    }
    sh.appendRow([
      new Date(), safe(name), safe(phone), venue, lang, score,
      won ? (repeat ? 'так (повторно)' : 'так') : 'ні',
      code, Number(d.attempt) || 1, 'так', ''
    ]);
    return reply({ ok: true, score: score, won: won, code: code, repeat: repeat, venue: venue });
  } finally {
    lock.releaseLock();
  }
}

function sheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Returns the prize code already issued to this phone number, if any.
function findPrize(sh, phone) {
  const last = sh.getLastRow();
  if (last < 2) return '';
  const rows = sh.getRange(2, 3, last - 1, 6).getValues(); // columns C..H
  for (let i = 0; i < rows.length; i++) {
    if (normPhone(rows[i][0]) === phone && rows[i][5]) return String(rows[i][5]);
  }
  return '';
}

function newCode(sh) {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const last = sh.getLastRow();
  const used = last < 2 ? [] : sh.getRange(2, 8, last - 1, 1).getValues().map(r => String(r[0]));
  for (;;) {
    let s = 'BOO-';
    for (let i = 0; i < 5; i++) s += abc[Math.floor(Math.random() * abc.length)];
    if (used.indexOf(s) < 0) return s;
  }
}

function normPhone(p) {
  const s = String(p || '').replace(/[\s\-()']/g, '');
  if (!/^\+?\d{9,15}$/.test(s)) return '';
  const digits = s.replace('+', '');
  return '+' + (digits.length === 9 ? '48' + digits : digits);
}

// Prefix values that a spreadsheet would treat as a formula.
function safe(v) {
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
