/**
 * База рапортов прораба — ГК «КРАШМАШ»
 * Google Apps Script, привязанный к Google Таблице.
 * Принимает рапорты из формы (GitHub Pages) и отдаёт историю по коду объекта.
 *
 * Установка: Расширения → Apps Script → вставить этот код → Сохранить →
 * Развернуть → Новое развертывание → Тип: Веб-приложение →
 * Запуск от имени: Я; У кого есть доступ: Все → Развернуть → Разрешить доступ →
 * скопировать URL веб-приложения и вставить его в Конструктор рапорта.
 */

// Ключи прорабов: у каждого объекта свой, вычисляется из секрета и кода объекта (выдаёт Claude вместе со ссылкой).
// Ключ объекта разрешает только сдать рапорт по своему объекту и прочитать рапорты своего объекта.
const FOREMAN_SECRET = 'ВСТАВЬТЕ_СЕКРЕТ_ПРОРАБОВ';
// Отозвать ссылки одного объекта: увеличить его номер, например { shum: 2 }, и выпустить новую ссылку.
const FOREMAN_GEN = {};

function foremanKey_(pid) {
  const gen = FOREMAN_GEN[pid] || 1;
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(pid) + '|' + gen, FOREMAN_SECRET)).replace(/=+$/, '').slice(0, 22);
}
function isForeman_(k, pid) { return !!k && !!pid && k === foremanKey_(pid); }
// Ключ кабинета руководителя: сводка и рапорты по всем объектам. Никому, кроме руководителей, не давать.
const DIRECTOR_KEY = 'ВСТАВЬТЕ_КЛЮЧ_КАБИНЕТА';
const TZ = 'Europe/Moscow';

const SHEETS = {
  R: ['Рапорты',  ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Сохранено на телефоне', 'Получено базой', 'Людей всего', 'Рейсов', 'Прочие работы', 'JSON']],
  W: ['Работы',   ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Код работы', 'Работа', 'Ед.', 'Объём за смену', 'По проекту']],
  H: ['Вывоз',    ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Перевозчик', 'Материал', 'Рейсов', 'м³', 'т']],
  P: ['Люди',     ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Организация', 'Людей']],
  M: ['Техника',  ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Техника', 'Кол-во', 'Моточасы']],
  D: ['Простои',  ['ID', 'Код объекта', 'Объект', 'Дата', 'Смена', 'Прораб', 'Что стояло', 'Часы', 'Причина', 'По чьей вине', 'Вид']],
  O: ['Объекты',  ['Код объекта', 'Объект', 'Обновлено', 'Настройки JSON']]
};

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEETS[key][0], hdr = SHEETS[key][1];
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#2E671F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < hdr.length) {
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#2E671F').setFontColor('#ffffff');
  }
  return sh;
}

/** Первичная настройка: создаёт все листы. Можно запустить вручную из редактора. */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  const s1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Лист1') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && SpreadsheetApp.getActiveSpreadsheet().getSheets().length > 1) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(s1);
}

function toDate_(iso) {
  try { return Utilities.parseDate(String(iso), TZ, 'yyyy-MM-dd'); } catch (e) { return iso; }
}

function deleteById_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return null;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  let first = null;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === id) { sh.deleteRow(i + 2); first = i + 2; }
  }
  return first;
}

/** Удалить рапорт по ID на всех листах (только если он того же объекта). */
function deleteReport_(id, pid) {
  if (!id || !pid) return false;
  const shR = sheet_('R'), n = shR.getLastRow();
  if (n < 2) return false;
  const v = shR.getRange(2, 1, n - 1, 2).getValues();
  let ok = false;
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === String(id) && String(v[i][1]) === String(pid)) { shR.deleteRow(i + 2); ok = true; }
  if (ok) ['W', 'H', 'P', 'M', 'D'].forEach(function (k) { deleteById_(sheet_(k), String(id)); });
  return ok;
}

function findRow_(sh, col, value) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === value) return i + 2;
  return 0;
}

function saveCfg_(cfg) {
  if (!cfg || !cfg.pid) return;
  const sh = sheet_('O');
  const row = [cfg.pid, cfg.name || '', new Date(), JSON.stringify(cfg)];
  const r = findRow_(sh, 1, cfg.pid);
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}

function saveReport_(r, cfg) {
  if (!r || !r.id || !r.pid || !r.date) throw new Error('bad report');
  const objName = (cfg && cfg.name) || '';
  const works = (cfg && cfg.works) || [];
  const base = [r.id, r.pid, objName, toDate_(r.date), r.shift === 'N' ? 'Ночь' : 'День', r.foreman || ''];
  const people = ((r.people && +r.people.own) || 0) + (r.subs || []).reduce(function (a, x) { return a + (+x.p || 0); }, 0);

  const shR = sheet_('R');
  const rowR = base.concat([r.savedAt || '', new Date(), people, r.trips || 0, r.other || '', JSON.stringify(r)]);
  const at = findRow_(shR, 1, r.id);
  if (at) shR.getRange(at, 1, 1, rowR.length).setValues([rowR]); else shR.appendRow(rowR);

  const put = function (key, rows) {
    const sh = sheet_(key);
    deleteById_(sh, r.id);
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  };
  put('W', Object.keys(r.works || {}).map(function (code) {
    const w = works.filter(function (x) { return x.c === code; })[0] || {};
    return base.concat([code, w.n || '', w.u || '', r.works[code], w.t || '']);
  }).concat((r.extra || []).map(function (x) {
    return base.concat(['новая', x.n || '', x.u || '', x.v || 0, '']);
  })));
  put('H', (r.haul || []).map(function (h) { return base.concat([h.c || '', h.m || '', h.t || 0, h.v || 0, h.w == null ? '' : h.w]); }));
  const ppl = [];
  if (r.people && r.people.own != null) ppl.push(base.concat(['Свои', r.people.own]));
  (r.subs || []).forEach(function (x) { ppl.push(base.concat([x.n || '', x.p || 0])); });
  put('P', ppl);
  put('M', (r.mach || []).map(function (m) { return base.concat([m.n || '', m.q || 0, m.h == null ? '' : m.h]); }));
  put('D', (r.downtime || []).map(function (d) { return base.concat([d.what || '', d.h || 0, d.why || '', d.fault || '', d.k || 'Простой']); }));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const dir = p.k === DIRECTOR_KEY;
  if (!dir && !(p.a === 'list' && isForeman_(p.k, p.pid))) return out_({ ok: false, error: 'key' });
  if (p.a === 'list' && !p.pid) return out_({ ok: false, error: 'pid' });
  if (p.a === 'ping') return out_({ ok: true, time: new Date().toISOString() });
  if (p.a === 'list') {
    const sh = sheet_('R');
    const n = sh.getLastRow();
    const reports = [];
    if (n >= 2) {
      const vals = sh.getRange(2, 1, n - 1, 12).getValues();
      vals.forEach(function (v) {
        if (String(v[1]) === String(p.pid)) { try { reports.push(JSON.parse(v[11])); } catch (err) {} }
      });
    }
    return out_({ ok: true, reports: reports });
  }
  if (p.a === 'objects') {
    const sh = sheet_('O');
    const n = sh.getLastRow();
    const objs = [];
    if (n >= 2) sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (v) { try { objs.push(JSON.parse(v[3])); } catch (err) {} });
    return out_({ ok: true, objects: objs });
  }
  if (p.a === 'dash') {
    const all = readAll_();
    const date = p.date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    return out_(Object.assign({ ok: true }, dashFrom_(all.cfgs, all.reports, date, p.days, DIRECTOR_PAGE)));
  }
  if (p.a === 'digest') {
    const date = p.date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    return out_(Object.assign({ ok: true }, buildDigest_(date)));
  }
  return out_({ ok: false, error: 'action' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (body.k !== DIRECTOR_KEY) {
    const pid = (body.cfg && body.cfg.pid) || (body.r && body.r.pid) || body.pid;
    if (!isForeman_(body.k, pid)) return out_({ ok: false, error: 'key' });
    if (body.r && body.r.pid !== pid) return out_({ ok: false, error: 'key' });
    if (body.a === 'save' && !body.r) return out_({ ok: false, error: 'key' });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    if (body.cfg) saveCfg_(body.cfg);
    if (body.a === 'save') {
      saveReport_(body.r, body.cfg);
      if (body.r.replaces && body.r.replaces !== body.r.id) deleteReport_(body.r.replaces, body.r.pid);
    }
    if (body.a === 'del' && body.id) deleteReport_(body.id, (body.cfg && body.cfg.pid) || body.pid);
    return out_({ ok: true, id: body.r ? body.r.id : null });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// СВОДКА ДЛЯ РУКОВОДИТЕЛЯ — каждый день в 20:00 (письмо и/или Telegram)
// =====================================================================
// Кому отправлять письмо (через запятую). Пусто — письмо не отправляется.
const DIGEST_TO = '';
// Telegram (необязательно): токен бота от @BotFather и id чатов через запятую.
const TG_BOT_TOKEN = '';
const TG_CHAT_IDS = '';
// Ссылка на страницу руководителя (подставляется в письмо).
const DIRECTOR_PAGE = '';

/** Запустить один раз вручную: создаёт ежедневный запуск в 20:00 по Москве. */
function setupDigest() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(20).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  if (DIGEST_TO) {
    MailApp.sendEmail({ to: DIGEST_TO, name: 'Рапорты прорабов', subject: 'Рассылка сводки включена',
      body: 'Сводка по объектам будет приходить каждый день около 20:00 по Москве.' + (DIRECTOR_PAGE ? '\nКабинет руководителя: ' + DIRECTOR_PAGE : '') });
  }
}

/** Отправить сводку за вчера — для проверки вида письма. */
function testDigestYesterday() {
  const y = new Date(); y.setDate(y.getDate() - 1);
  const d = buildDigest_(Utilities.formatDate(y, TZ, 'yyyy-MM-dd'));
  if (DIGEST_TO) MailApp.sendEmail({ to: DIGEST_TO, subject: '[проверка] ' + d.subject, body: d.text, htmlBody: d.html, name: 'Рапорты прорабов' });
}

/** Собрать сводку за сегодня и отправить. Можно запустить вручную для проверки. */
function dailyDigest() {
  const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const d = buildDigest_(date);
  if (DIGEST_TO) {
    MailApp.sendEmail({ to: DIGEST_TO, subject: d.subject, body: d.text, htmlBody: d.html, name: 'Рапорты прорабов' });
  }
  if (TG_BOT_TOKEN && TG_CHAT_IDS) {
    const parts = splitText_(d.text, 3900);
    TG_CHAT_IDS.split(',').map(function (s) { return s.trim(); }).filter(String).forEach(function (chat) {
      parts.forEach(function (p) {
        UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_BOT_TOKEN + '/sendMessage', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ chat_id: chat, text: p, disable_web_page_preview: true })
        });
      });
    });
  }
  return d;
}

function splitText_(t, n) {
  const out = []; let cur = '';
  t.split('\n').forEach(function (line) { if ((cur + line).length > n) { out.push(cur); cur = ''; } cur += line + '\n'; });
  if (cur.trim()) out.push(cur);
  return out;
}

function readAll_() {
  const cfgs = [], reports = [];
  const so = sheet_('O'), no = so.getLastRow();
  if (no >= 2) so.getRange(2, 1, no - 1, 4).getValues().forEach(function (v) {
    try { const c = JSON.parse(v[3]); if (v[2] instanceof Date) c._upd = Utilities.formatDate(v[2], TZ, 'yyyy-MM-dd'); cfgs.push(c); } catch (e) {}
  });
  const sr = sheet_('R'), nr = sr.getLastRow();
  if (nr >= 2) sr.getRange(2, 12, nr - 1, 1).getValues().forEach(function (v) { try { reports.push(JSON.parse(v[0])); } catch (e) {} });
  return { cfgs: cfgs, reports: reports };
}

function buildDigest_(date) {
  const all = readAll_();
  return digestFrom_(all.cfgs, all.reports, date, DIRECTOR_PAGE);
}

/** Чистая функция: настройки объектов + рапорты → текст и HTML сводки. */
function digestFrom_(cfgs, reports, date, page) {
  const f = function (v, d) {
    if (v === null || v === undefined || v === '' || isNaN(v)) return '—';
    const p = Math.pow(10, d === undefined ? 2 : d); let s = String(Math.round(Number(v) * p) / p);
    const neg = s[0] === '-'; if (neg) s = s.slice(1);
    const sp = s.split('.'); sp[0] = sp[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (neg ? '-' : '') + sp.join(',');
  };
  const ru = function (s) { return s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(0, 4); };
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const nrm = function (s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); };
  const addDays = function (iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const from7 = addDays(date, -7);
  const byPid = {}; reports.forEach(function (r) { if (r && r.pid && r.date) (byPid[r.pid] = byPid[r.pid] || []).push(r); });
  const cfgBy = {}; cfgs.forEach(function (c) { if (c && c.pid) cfgBy[c.pid] = c; });
  Object.keys(byPid).forEach(function (p) { if (!cfgBy[p]) cfgBy[p] = { pid: p, name: p, works: [] }; });
  const objs = Object.keys(cfgBy).map(function (p) { return cfgBy[p]; }).filter(function (c) {
    return (c._upd && c._upd >= addDays(date, -14) && c._upd <= date) || (byPid[c.pid] || []).some(function (r) { return r.date >= from7 && r.date <= date; });
  }).sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'ru'); });

  const T = [], H = [];
  let okN = 0, ppl = 0, vol = 0, ton = 0, dtH = 0; const missing = [];
  const blocks = objs.map(function (c) {
    const rs = (byPid[c.pid] || []);
    const day = rs.filter(function (r) { return r.date === date; }).sort(function (a, b) { return a.shift < b.shift ? -1 : 1; });
    const cum = rs.filter(function (r) { return r.date <= date; });
    const hl = c.hl || {}; const haulT = hl.t || 'Вывоз';
    const b = { name: c.name || c.pid, ok: day.length > 0, lines: [] };
    if (!day.length) { missing.push(b.name); return b; }
    okN++;
    b.who = day.map(function (r) { return (r.shift === 'N' ? 'ночь' : 'день') + ' — ' + (r.foreman || '?'); }).join('; ');
    // работы
    (c.works || []).forEach(function (w) {
      const s = day.reduce(function (a, r) { const v = r.works && r.works[w.c]; return a + (typeof v === 'number' ? v : 0); }, 0);
      if (!s) return;
      if (String(w.src || '').indexOf('mach:') === 0 && !w.t) return; // моточасы без плана — видны в строке «Техника»
      const t = cum.reduce(function (a, r) { const v = r.works && r.works[w.c]; return a + (typeof v === 'number' ? v : 0); }, 0) + (w.d0 || 0);
      b.lines.push(['w', w.n + ': ' + f(s) + ' ' + w.u + ' · с начала ' + f(t) + (w.t ? ' из ' + f(w.t) + ' (' + f(t / w.t * 100, 1) + '%)' : ''), w.t && t > w.t]);
    });
    day.forEach(function (r) { (r.extra || []).forEach(function (e) { b.lines.push(['w', e.n + ': ' + f(e.v) + ' ' + e.u]); }); });
    // вывоз
    const hm = {}; day.forEach(function (r) { (r.haul || []).forEach(function (h) { const k = h.m || '—'; hm[k] = hm[k] || { t: 0, v: 0, w: 0 }; hm[k].t += +h.t || 0; hm[k].v += +h.v || 0; hm[k].w += +h.w || 0; }); });
    const hk = Object.keys(hm);
    if (hk.length) {
      const tv = hk.reduce(function (a, k) { return a + hm[k].v; }, 0), tt = hk.reduce(function (a, k) { return a + hm[k].t; }, 0), tw = hk.reduce(function (a, k) { return a + hm[k].w; }, 0);
      vol += tv; ton += tw;
      b.lines.push(['h', haulT + ': ' + (tt ? f(tt, 0) + ' рейс., ' : '') + f(tv) + ' м³' + (hl.tn ? ' / ' + f(tw) + ' т' : '') + ' — ' + hk.map(function (k) { return k + ' ' + f(hm[k].v) + (hl.tn ? '/' + f(hm[k].w) + ' т' : ''); }).join('; ')]);
    }
    // люди
    const pp = []; let pn = 0, hasOwn = false, ownS = 0;
    day.forEach(function (r) {
      const own = r.people && r.people.own; if (own != null && own !== '') { hasOwn = true; ownS += +own || 0; pn += +own || 0; }
      (r.subs || []).forEach(function (x) { pn += +x.p || 0; pp.push(x.n + ' ' + (x.p != null ? x.p : '?')); });
    });
    ppl += pn;
    if (hasOwn) pp.unshift('свои ' + ownS);
    b.lines.push(['p', pp.length ? 'Люди: ' + pn + ' (' + pp.join(', ') + ')' : 'Люди: не указаны']);
    // техника
    const mm = {}; day.forEach(function (r) { (r.mach || []).forEach(function (m) { const k = nrm(m.n); mm[k] = mm[k] || { n: m.n, q: 0, h: 0, hh: false }; mm[k].q = Math.max(mm[k].q, +m.q || 0); if (m.h != null) { mm[k].h += +m.h || 0; mm[k].hh = true; } }); });
    const mk = Object.keys(mm);
    if (mk.length) b.lines.push(['m', 'Техника: ' + mk.map(function (k) { const m = mm[k]; return m.n + (m.q > 1 ? ' ×' + m.q : '') + (m.hh ? ' — ' + f(m.h) + ' м/ч' : ''); }).join('; ')]);
    // простои
    const dts = []; day.forEach(function (r) { (r.downtime || []).forEach(function (d) { dtH += +d.h || 0; dts.push((d.k || 'Простой') + ': ' + (d.what || '') + ' — ' + f(d.h) + ' ч' + (d.why ? ', ' + d.why : '') + (d.fault ? ' (вина: ' + d.fault + ')' : '')); }); });
    dts.forEach(function (s) { b.lines.push(['d', s, true]); });
    day.forEach(function (r) { if (r.other) b.lines.push(['n', 'Примечание: ' + r.other]); });
    return b;
  });

  const subject = 'Сводка по объектам за ' + ru(date) + ' — рапортов ' + okN + ' из ' + objs.length;
  T.push(subject.toUpperCase());
  T.push('Людей на объектах: ' + ppl + ' · вывезено/продано: ' + f(vol) + ' м³' + (ton ? ' (' + f(ton) + ' т)' : '') + ' · простои и ремонт: ' + f(dtH) + ' ч');
  if (missing.length) T.push('Нет рапорта: ' + missing.join('; '));
  blocks.forEach(function (b) {
    T.push(''); T.push((b.ok ? '■ ' : '□ ') + b.name);
    if (!b.ok) { T.push('  нет рапорта за ' + ru(date)); return; }
    T.push('  ' + b.who);
    b.lines.forEach(function (l) { T.push('  — ' + l[1]); });
  });
  if (!objs.length) T.push('Активных объектов нет (за последние 7 дней рапортов не было).');
  if (page) { T.push(''); T.push('Подробно: ' + page); }

  const G = '#2E671F';
  H.push('<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;color:#1d1d1b">');
  H.push('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:' + G + ';border-radius:8px 8px 0 0"><tr><td style="padding:16px 18px;color:#ffffff;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="font-size:26px;line-height:30px;font-weight:bold;letter-spacing:.5px;color:#ffffff">ГК «КРАШМАШ»</div>' +
    '<div style="font-size:19px;line-height:24px;font-weight:bold;color:#ffffff;margin-top:6px">Сводка по объектам за ' + ru(date) + '</div>' +
    '<div style="font-size:13px;line-height:18px;color:#dfe9da;margin-top:4px">Рапорты прорабов · ' + okN + ' из ' + objs.length + ' объектов сдали рапорт</div>' +
    '</td></tr></table>');
  H.push('<table style="width:100%;border-collapse:collapse;background:#f3f5f1;font-size:14px"><tr>' +
    [['Рапортов', okN + ' из ' + objs.length], ['Людей', ppl], [vol || !ton ? 'Вывоз, м³' : '', f(vol)], ['Простои/ремонт, ч', f(dtH)]].map(function (k) {
      return '<td style="padding:10px 12px"><div style="color:#666;font-size:12px">' + esc(k[0]) + '</div><div style="font-size:18px;font-weight:bold">' + esc(k[1]) + '</div></td>';
    }).join('') + '</tr></table>');
  if (missing.length) H.push('<div style="background:#fdecea;color:#8a1f17;padding:10px 12px;font-size:14px"><b>Нет рапорта:</b> ' + esc(missing.join('; ')) + '</div>');
  blocks.forEach(function (b) {
    H.push('<div style="border:1px solid #dde3d8;border-top:0;padding:12px 14px">');
    H.push('<div style="font-weight:bold;font-size:15px;color:' + (b.ok ? G : '#8a1f17') + '">' + esc(b.name) + '</div>');
    if (!b.ok) { H.push('<div style="color:#8a1f17;font-size:14px">Нет рапорта за ' + ru(date) + '</div></div>'); return; }
    H.push('<div style="color:#666;font-size:12px;margin-bottom:6px">' + esc(b.who) + '</div><ul style="margin:0;padding-left:18px;font-size:14px;line-height:1.5">');
    b.lines.forEach(function (l) { H.push('<li' + (l[2] ? ' style="color:#8a1f17"' : '') + '>' + esc(l[1]) + '</li>'); });
    H.push('</ul></div>');
  });
  if (!objs.length) H.push('<div style="padding:12px">Активных объектов нет (за последние 7 дней рапортов не было).</div>');
  if (page) H.push('<div style="padding:12px 0;font-size:13px"><a href="' + esc(page) + '" style="color:' + G + '">Открыть страницу руководителя</a></div>');
  H.push('</div>');
  return { date: date, subject: subject, text: T.join('\n'), html: H.join(''), total: objs.length, ok: okN, missing: missing };
}

// ---------- данные для приложения руководителя ----------
function dashFrom_(cfgs, reports, date, days, page) {
  days = Math.max(1, Math.min(14, +days || 3));
  const f = function (v, d) {
    if (v === null || v === undefined || v === '' || isNaN(v)) return '—';
    const p = Math.pow(10, d === undefined ? 2 : d); let s = String(Math.round(Number(v) * p) / p);
    const neg = s[0] === '-'; if (neg) s = s.slice(1);
    const sp = s.split('.'); sp[0] = sp[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (neg ? '-' : '') + sp.join(',');
  };
  const addDays = function (iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const nrm = function (s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); };
  const from = addDays(date, -(days - 1)), from7 = addDays(date, -7);
  const byPid = {}; reports.forEach(function (r) { if (r && r.pid && r.date) (byPid[r.pid] = byPid[r.pid] || []).push(r); });
  const cfgBy = {}; cfgs.forEach(function (c) { if (c && c.pid) cfgBy[c.pid] = c; });
  Object.keys(byPid).forEach(function (p) { if (!cfgBy[p]) cfgBy[p] = { pid: p, name: p, works: [] }; });
  const objects = [], feed = [];
  Object.keys(cfgBy).forEach(function (pid) {
    const c = cfgBy[pid], rs = (byPid[pid] || []).filter(function (r) { return r.date <= date; });
    const cum = {}; rs.forEach(function (r) { Object.keys(r.works || {}).forEach(function (k) { if (typeof r.works[k] === 'number') cum[k] = (cum[k] || 0) + r.works[k]; }); });
    const ex = {}; rs.forEach(function (r) { (r.extra || []).forEach(function (e) { const k = nrm(e.n) + '|' + nrm(e.u); ex[k] = ex[k] || { n: e.n, u: e.u, v: 0 }; ex[k].v += +e.v || 0; }); });
    const last = rs.slice().sort(function (a, b) { return (b.date + b.shift).localeCompare(a.date + a.shift); })[0];
    const machSrc = function (w) { return String(w.src || '').indexOf('mach:') === 0; };
    objects.push({
      pid: pid, name: c.name || pid, title: c.title || '', info: c.info || [], d0date: c.d0date || '',
      works: (c.works || []).map(function (w) { return { c: w.c, n: w.n, u: w.u, t: w.t || 0, done: Math.round(((cum[w.c] || 0) + (w.d0 || 0)) * 1000) / 1000, auto: !!w.src, ro: !!w.ro }; }),
      extra: Object.keys(ex).map(function (k) { return ex[k]; }),
      last: last ? { date: last.date, shift: last.shift, foreman: last.foreman || '' } : null,
      today: rs.some(function (r) { return r.date === date; }),
      active: rs.some(function (r) { return r.date >= from7; }),
      reports: rs.length
    });
    rs.filter(function (r) { return r.date >= from; }).forEach(function (r) { feed.push({ pid: pid, name: c.name || pid, date: r.date, shift: r.shift, foreman: r.foreman || '', savedAt: r.savedAt || '', text: reportText_(c, r, cum, f) }); });
  });
  objects.sort(function (a, b) { return (b.active - a.active) || String(a.name).localeCompare(String(b.name), 'ru'); });
  feed.sort(function (a, b) { return (b.date + b.shift + b.savedAt).localeCompare(a.date + a.shift + a.savedAt); });
  const dg = digestFrom_(cfgs, reports, date, page);
  return { date: date, days: days, objects: objects, feed: feed, digest: { subject: dg.subject, html: dg.html, text: dg.text, ok: dg.ok, total: dg.total, missing: dg.missing } };
}

function reportText_(c, r, cumAll, f) {
  const L = [], hl = c.hl || {};
  (c.works || []).forEach(function (w) { const v = r.works && r.works[w.c]; if (typeof v === 'number' && v) L.push('— ' + w.n + ': ' + f(v) + ' ' + w.u); });
  (r.extra || []).forEach(function (e) { L.push('— ' + e.n + ': ' + f(e.v) + ' ' + e.u); });
  if (!L.length) L.push('— объёмов нет');
  if ((r.haul || []).length) { L.push((hl.t || 'Вывоз') + ':'); r.haul.forEach(function (h) { L.push('— ' + (h.c || '') + ': ' + (h.m || '') + ' — ' + (h.t != null ? f(h.t, 0) + ' рейс., ' : '') + f(h.v) + ' м³' + (h.w != null ? ' (' + f(h.w) + ' т)' : '')); }); }
  const pp = []; let tot = 0; if (r.people && r.people.own != null) { pp.push('свои ' + r.people.own); tot += +r.people.own || 0; }
  (r.subs || []).forEach(function (x) { pp.push((x.n || '') + ' ' + (x.p != null ? x.p : '?')); tot += +x.p || 0; });
  L.push('Люди: ' + (pp.length ? pp.join(', ') + ' — всего ' + tot : 'нет данных'));
  if ((r.mach || []).length) L.push('Техника: ' + r.mach.map(function (m) { return m.n + ' ' + (m.q || 1) + (m.h != null ? ' (' + f(m.h) + ' м/ч)' : ''); }).join(', '));
  (r.downtime || []).forEach(function (d) { L.push((d.k || 'Простой') + ': ' + (d.what || '') + ' — ' + f(d.h) + ' ч' + (d.why ? ' — ' + d.why : '') + (d.fault ? ' (вина: ' + d.fault + ')' : '')); });
  if (r.other) L.push('Примечание: ' + r.other);
  return L.join('\n');
}
