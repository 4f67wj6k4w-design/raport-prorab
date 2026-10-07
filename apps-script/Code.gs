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

/** Выполнено с начала по работе на дату d. «Выполнено ранее» (w.d0) — по состоянию на c.d0date:
 *  рапорты до этой даты включительно — только история, к итогу не прибавляются (иначе двойной счёт). */
function cumAt_(c, rs, code, d) {
  const d0 = (((c.works || []).filter(function (w) { return w.c === code; })[0]) || {}).d0 || 0, dd = c.d0date || '';
  let t = +d0 || 0;
  rs.forEach(function (r) {
    const v = r.works && r.works[code]; if (typeof v !== 'number') return;
    if (!dd) { if (r.date <= d) t += v; }
    else if (r.date > dd && r.date <= d) t += v;
    else if (r.date <= dd && r.date > d) t -= v;
  });
  return t;
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
  if (p.a === 'exec' && (dir || (EXEC_KEY.length >= 24 && p.k === EXEC_KEY))) {
    const date = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '') ? p.date : Utilities.formatDate(new Date(Date.now() - 864e5), TZ, 'yyyy-MM-dd');
    return out_(Object.assign({ ok: true }, execModel_(date)));
  }
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
      const t = cumAt_(c, rs, w.c, date);
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
      works: (c.works || []).map(function (w) { return { c: w.c, n: w.n, u: w.u, t: w.t || 0, done: Math.round(cumAt_(c, byPid[pid] || [], w.c, date) * 1000) / 1000, auto: !!w.src, ro: !!w.ro }; }),
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

/** Меню в таблице: «Рапорты → Удалить выделенный рапорт». Удаляет рапорт целиком со всех листов. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Рапорты')
    .addItem('Удалить выделенный рапорт', 'menuDeleteReport')
    .addToUi();
}
function menuDeleteReport() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const keys = ['R', 'W', 'H', 'P', 'M', 'D'].map(function (k) { return SHEETS[k][0]; });
  const r = sh.getActiveRange().getRow();
  if (keys.indexOf(sh.getName()) < 0 || r < 2) { ui.alert('Встаньте на строку рапорта (лист «Рапорты», «Работы», «Техника» и т.п.) и выберите пункт меню ещё раз.'); return; }
  const v = sh.getRange(r, 1, 1, 6).getValues()[0];
  const id = String(v[0]), pid = String(v[1]);
  if (!id || !pid) { ui.alert('В этой строке нет рапорта.'); return; }
  const ans = ui.alert('Удалить рапорт?', 'Объект: ' + v[2] + '\nДата: ' + Utilities.formatDate(new Date(v[3]), TZ, 'dd.MM.yyyy') + ', смена: ' + v[4] + '\nПрораб: ' + v[5] + '\n\nРапорт удалится со всех листов (работы, вывоз, люди, техника, простои).', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  ui.alert(deleteReport_(id, pid) ? 'Рапорт удалён.' : 'Рапорт не найден на листе «Рапорты».');
}

// =====================================================================
// ПАНЕЛЬ РУКОВОДСТВА ГК — светофор по объектам, графики, неделя к неделе, письмо в 8:00
// =====================================================================
// Ключ панели руководства: видит только светофор и графики (не рапорты и не настройки).
const EXEC_KEY = 'ВСТАВЬТЕ_КЛЮЧ_РУКОВОДСТВА';
// Кому утреннее письмо со светофором (через запятую). Пусто — письмо не отправляется.
const EXEC_TO = '';
// Ссылка на панель руководства (подставляется в письмо).
const EXEC_PAGE = '';
// Пороги светофора
const EXEC_RULES = {
  dtYellow: 8,      // простои+ремонт за 7 дней, ч — жёлтый
  dtRed: 24,        // … — красный
  paceYellow: 0.67, // объём за неделю меньше 2/3 прошлой недели — жёлтый
  paceRed: 0.34,    // … меньше 1/3 — красный
  dueWarnDays: 7,   // прогноз окончания ближе 7 дней к сроку — жёлтый
  lagYellow: 5,     // отставание от равномерного графика, п.п. — жёлтый
  lagRed: 15        // … — красный
};

/** Работа участвует в «% выполнения объекта»: есть план, не техника, не «в т.ч.», не справочно. */
function execProgressWork_(w) {
  if (!w || !(+w.t > 0) || w.ro) return false;
  const u = String(w.u || '').trim().toLowerCase(), n = String(w.n || '').trim().toLowerCase();
  if (u.indexOf('маш') === 0) return false;
  if (/^мч/i.test(String(w.c || ''))) return false;
  if (/^(в т\.ч\.|работа|мобилизац|демобилизац|доставка быто|бытовк|пожарн)/.test(n)) return false;
  return true;
}

/** Срок из паспорта: cfg.due / cfg.start или строка «Сроки…: 01.09.2026 – 12.11.2026». */
function execDue_(c) {
  const iso = function (s) { const m = /(\d{2})\.(\d{2})\.(\d{4})/.exec(s || ''); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; };
  let start = c.start || null, due = c.due || null;
  (c.info || []).forEach(function (s) {
    if (due || !/срок/i.test(s)) return;
    const m = /(\d{2}\.\d{2}\.\d{4})\s*[–—-]\s*(\d{2}\.\d{2}\.\d{4})/.exec(s);
    if (m) { start = start || iso(m[1]); due = iso(m[2]); }
    else { const m2 = /до\s*(\d{2}\.\d{2}\.\d{4})/.exec(s); if (m2) due = iso(m2[1]); }
  });
  return { start: start, due: due };
}

/** Чистая функция: настройки объектов + рапорты → модель панели руководства на дату date (обычно вчера). */
function execFrom_(cfgs, reports, date, today) {
  const R = EXEC_RULES;
  const add = function (iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const diff = function (a, b) { return Math.round((new Date(a + 'T12:00:00Z') - new Date(b + 'T12:00:00Z')) / 864e5); };
  const ru = function (s) { return s ? s.slice(8, 10) + '.' + s.slice(5, 7) : ''; };
  const ruY = function (s) { return s ? s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(0, 4) : ''; };
  const f = function (v, d) { const k = Math.pow(10, d == null ? 1 : d); return String(Math.round(v * k) / k).replace('.', ','); };
  const nrm = function (s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); };
  const sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  const N = 30, from = add(date, -(N - 1));
  const dates = []; for (let i = 0; i < N; i++) dates.push(add(from, i));
  const byPid = {}; reports.forEach(function (r) { if (r && r.pid && r.date) (byPid[r.pid] = byPid[r.pid] || []).push(r); });
  const cfgBy = {}; cfgs.forEach(function (c) { if (c && c.pid) cfgBy[c.pid] = c; });
  Object.keys(byPid).forEach(function (p) { if (!cfgBy[p]) cfgBy[p] = { pid: p, name: p, works: [] }; });
  const rank = { r: 3, y: 2, g: 1, n: 0 };

  const objects = Object.keys(cfgBy).map(function (pid) {
    const c = cfgBy[pid], all = (byPid[pid] || []).filter(function (r) { return r.date <= date; });
    if (!all.length) {
      if (!(c._upd && diff(date, c._upd) <= 14 && c._upd <= date)) return null;
      const nm = String(c.name || pid);
      return { pid: pid, name: nm, short: nm.replace(/^г\.\s*Москва,\s*/i, '').split(' — ')[0], kind: nm.split(' — ').slice(1).join(' — '), light: 'n',
        reasons: [{ k: 'rep', l: 'n', t: 'Рапортов по объекту ещё не было' }], pct: null, due: execDue_(c).due, start: execDue_(c).start, forecast: null, first: null,
        last: null, today: !!today && (byPid[pid] || []).some(function (r) { return r.date === today; }), day: null, series: null, works: [], week: null, empty: true };
    }
    const first = all.reduce(function (a, r) { return r.date < a ? r.date : a; }, all[0].date);
    const last = all.reduce(function (a, r) { return !a || r.date > a.date || (r.date === a.date && r.shift > a.shift) ? r : a; }, null);
    if (diff(date, last.date) > 14) return null; // объект не ведётся
    const on = function (d) { return all.filter(function (r) { return r.date === d; }); };
    const has = function (d) { return on(d).length > 0; };
    const works = c.works || [];
    const prog = works.filter(execProgressWork_);
    const reasons = [];

    // --- ряды по дням
    const S = { people: [], mach: [], mh: [], m3: [], trips: [], dt: [], rep: [], idle: [] };
    dates.forEach(function (d) {
      const rs = on(d);
      S.rep.push(rs.length ? 1 : 0);
      S.idle.push(rs.length && rs.every(function (r) { return r.idle; }) ? 1 : 0);
      S.people.push(rs.length ? sum(rs.map(function (r) { return ((r.people && +r.people.own) || 0) + sum((r.subs || []).map(function (x) { return +x.p || 0; })); })) : null);
      const mm = {}; let mh = 0;
      rs.forEach(function (r) { (r.mach || []).forEach(function (m) { const k = nrm(m.n); mm[k] = Math.max(mm[k] || 0, +m.q || 1); mh += +m.h || 0; }); });
      S.mach.push(rs.length ? sum(Object.keys(mm).map(function (k) { return mm[k]; })) : null);
      S.mh.push(rs.length ? mh : null);
      S.m3.push(rs.length ? sum(rs.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.v || 0; })); })) : null);
      S.trips.push(rs.length ? sum(rs.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.t || 0; })); })) : null);
      S.dt.push(rs.length ? sum(rs.map(function (r) { return sum((r.downtime || []).filter(function (x) { return x.k !== 'ТО'; }).map(function (x) { return +x.h || 0; })); })) : null);
    });

    // --- % выполнения объекта на дату
    const doneBy = function (w, d) {
      return cumAt_(c, byPid[pid] || [], w.c, d);
    };
    const pctOn = function (d) { return prog.length ? sum(prog.map(function (w) { return Math.max(0, Math.min(1, doneBy(w, d) / w.t)); })) / prog.length * 100 : null; };
    const pct = pctOn(date);
    S.pct = prog.length ? dates.map(function (d) { return d < first ? null : Math.round(pctOn(d) * 10) / 10; }) : null;

    // --- 1. Рапорты
    if (!has(date) && !has(add(date, -1))) reasons.push({ k: 'rep', l: 'r', t: 'Нет рапортов 2 дня и больше — последний ' + ru(last.date) });
    else if (!has(date)) reasons.push({ k: 'rep', l: 'y', t: 'Нет рапорта за ' + ru(date) });
    else reasons.push({ k: 'rep', l: 'g', t: on(date).every(function (r) { return r.idle; }) ? 'Рапорт сдан: работ не было' : 'Рапорт сдан' });

    // --- 2. Простои за 7 дней
    const w1 = all.filter(function (r) { return r.date > add(date, -7); });
    const dts = []; w1.forEach(function (r) { (r.downtime || []).forEach(function (x) { if (x.k !== 'ТО') dts.push(x); }); });
    const dtH = sum(dts.map(function (x) { return +x.h || 0; }));
    const cust = sum(dts.filter(function (x) { return (x.k || 'Простой') === 'Простой' && (x.fault === 'Заказчик' || x.fault === 'Генподрядчик'); }).map(function (x) { return +x.h || 0; }));
    let dl = dtH >= R.dtRed ? 'r' : dtH >= R.dtYellow ? 'y' : 'g';
    if (cust > 0 && dl === 'g') dl = 'y';
    reasons.push({ k: 'dt', l: dl, t: dtH ? 'Простои и ремонт за 7 дней: ' + f(dtH) + ' ч' + (cust ? ' (по вине заказчика/генподрядчика ' + f(cust) + ' ч — основание для продления срока)' : '') : 'Простоев за 7 дней нет' });

    // --- 3. Темп: неделя к неделе
    const wk = function (a, b) { return all.filter(function (r) { return r.date > add(date, a) && r.date <= add(date, b); }); };
    const W1 = wk(-7, 0), W0 = wk(-14, -7);
    const volW = function (rs, code) { return sum(rs.map(function (r) { const v = r.works && r.works[code]; return typeof v === 'number' ? v : 0; })); };
    const items = works.filter(function (w) { return !w.ro && !/^мч/i.test(String(w.c || '')) && String(w.u || '').toLowerCase().indexOf('маш') !== 0 && !/^в т\.ч\./i.test(String(w.n || '')); })
      .map(function (w) { return { n: w.n, u: w.u, a: volW(W1, w.c), b: volW(W0, w.c) }; });
    const extraW = {}; W1.concat(W0).forEach(function (r) { (r.extra || []).forEach(function (e) { const k = e.n + '|' + e.u; extraW[k] = extraW[k] || { n: e.n, u: e.u, a: 0, b: 0 }; extraW[k][r.date > add(date, -7) ? 'a' : 'b'] += +e.v || 0; }); });
    Object.keys(extraW).forEach(function (k) { items.push(extraW[k]); });
    const haulA = sum(W1.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.v || 0; })); })), haulB = sum(W0.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.v || 0; })); }));
    const act = items.filter(function (x) { return x.a + x.b > 0; });
    if (haulA + haulB > 0 && !act.some(function (x) { return /вывоз/i.test(x.n); })) act.push({ n: 'Вывоз', u: 'м³', a: haulA, b: haulB });
    if (pct !== null && pct >= 99.5) reasons.push({ k: 'pace', l: 'n', t: 'План выполнен — темп не оценивается' });
    else if (diff(date, first) < 13) reasons.push({ k: 'pace', l: 'n', t: 'Темп неделя к неделе — после ' + ru(add(first, 13)) + ' (нужно 2 недели рапортов)' });
    else if (!act.length) reasons.push({ k: 'pace', l: W1.length ? 'y' : 'n', t: 'Объёмов работ за 2 недели нет' });
    else {
      const sh = sum(act.map(function (x) { return x.a / (x.a + x.b); })) / act.length;
      const ratio = sh >= 0.999 ? Infinity : sh / (1 - sh);
      const pl = ratio < R.paceRed ? 'r' : ratio < R.paceYellow ? 'y' : 'g';
      const ch = isFinite(ratio) ? Math.round((ratio - 1) * 100) : null;
      reasons.push({ k: 'pace', l: pl, t: ch === null ? 'Работы начались на этой неделе' : ch <= -5 ? 'Темп упал на ' + (-ch) + '% к прошлой неделе' : ch >= 5 ? 'Темп вырос на ' + ch + '% к прошлой неделе' : 'Темп как на прошлой неделе' });
    }

    // --- 4. Срок и прогноз окончания
    const dd = execDue_(c);
    let forecast = null, still = false;
    const span = Math.min(7, diff(date, first) + 1);
    if (pct !== null && pct < 99.5 && span >= 3) {
      const p0 = pctOn(add(date, -span)), v = (pct - p0) / span;
      if (v > 0.01) forecast = add(date, Math.ceil((100 - pct) / v)); else still = true;
    }
    let due = { k: 'due', l: 'n', t: '' };
    if (pct !== null && pct >= 99.5) due = { k: 'due', l: 'g', t: 'План выполнен' };
    else if (dd.due && pct !== null) {
      if (date > dd.due) due = { k: 'due', l: 'r', t: 'Срок по договору ' + ruY(dd.due) + ' истёк, выполнено ' + f(pct, 0) + '%' };
      else if (forecast) {
        const late = diff(forecast, dd.due);
        due = { k: 'due', l: late > 0 ? 'r' : late > -R.dueWarnDays ? 'y' : 'g', t: 'Прогноз окончания ' + ruY(forecast) + (late > 0 ? ' — позже срока ' + ruY(dd.due) + ' на ' + late + ' дн.' : ', срок ' + ruY(dd.due)) };
      } else if (dd.start) {
        const plan = Math.max(0, Math.min(100, (diff(date, dd.start) + 1) / (diff(dd.due, dd.start) + 1) * 100)), lag = plan - pct;
        due = { k: 'due', l: lag >= R.lagRed ? 'r' : lag >= R.lagYellow ? 'y' : 'g', t: (lag >= R.lagYellow ? 'Отставание от графика ' + f(lag, 0) + ' п.п.' : 'Идём по графику') + ': выполнено ' + f(pct, 0) + '%, по сроку к ' + ru(date) + ' нужно ' + f(plan, 0) + '% (срок ' + ruY(dd.due) + ')' };
      } else due.t = 'Срок ' + ruY(dd.due) + (still ? ', по работам с объёмом по проекту движения нет — прогноз невозможен' : ', прогноз появится после 3 дней рапортов');
    } else if (pct === null) due.t = 'Прогноз невозможен: в паспорте нет объёмов по проекту';
    else due.t = forecast ? 'Прогноз окончания ' + ruY(forecast) + ' (срока в паспорте нет)' : still ? 'Срока в паспорте нет; по работам с объёмом по проекту за неделю движения нет' : 'Срока в паспорте нет; прогноз — после 3 дней рапортов';
    reasons.push(due);

    const light = reasons.reduce(function (a, x) { return rank[x.l] > rank[a] ? x.l : a; }, 'n');

    // --- работы для графиков: с планом или с объёмами за 30 дней
    const wlist = works.filter(function (w) { return !w.ro && !/^мч/i.test(String(w.c || '')) && String(w.u || '').toLowerCase().indexOf('маш') !== 0; }).map(function (w) {
      const d = dates.map(function (x) { const rs = on(x); return rs.length ? sum(rs.map(function (r) { const v = r.works && r.works[w.c]; return typeof v === 'number' ? v : 0; })) : null; });
      const done = doneBy(w, date);
      return { c: w.c, n: w.n, u: w.u, t: +w.t || 0, done: Math.round(done * 100) / 100, pct: +w.t > 0 ? Math.round(done / w.t * 1000) / 10 : null, d: d, act: sum(d.map(function (x) { return x || 0; })) };
    }).filter(function (w) { return w.act > 0 || (w.t > 0 && execProgressWork_(works.filter(function (x) { return x.c === w.c; })[0])); });

    const week = { a: [add(date, -6), date], b: [add(date, -13), add(date, -7)], rows: [] };
    const avg = function (rs, fn) { const ds = {}; rs.forEach(function (r) { ds[r.date] = (ds[r.date] || 0) + fn(r); }); const k = Object.keys(ds); return k.length ? sum(k.map(function (x) { return ds[x]; })) / k.length : 0; };
    const ppl = function (r) { return ((r.people && +r.people.own) || 0) + sum((r.subs || []).map(function (x) { return +x.p || 0; })); };
    items.filter(function (x) { return x.a + x.b > 0; }).forEach(function (x) { week.rows.push({ n: x.n, u: x.u, a: x.a, b: x.b }); });
    week.rows.push({ n: 'Людей в среднем за день', u: 'чел.', a: Math.round(avg(W1, ppl)), b: Math.round(avg(W0, ppl)), s: 1 });
    week.rows.push({ n: 'Моточасы техники', u: 'маш.-ч', a: sum(W1.map(function (r) { return sum((r.mach || []).map(function (m) { return +m.h || 0; })); })), b: sum(W0.map(function (r) { return sum((r.mach || []).map(function (m) { return +m.h || 0; })); })), s: 1 });
    week.rows.push({ n: 'Рейсов вывоза', u: 'рейс.', a: sum(W1.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.t || 0; })); })), b: sum(W0.map(function (r) { return sum((r.haul || []).map(function (h) { return +h.t || 0; })); })), s: 1 });
    week.rows.push({ n: 'Вывоз', u: 'м³', a: haulA, b: haulB, s: 1 });
    week.rows.push({ n: 'Простои и ремонт', u: 'ч', a: dtH, b: sum(W0.map(function (r) { return sum((r.downtime || []).filter(function (x) { return x.k !== 'ТО'; }).map(function (x) { return +x.h || 0; })); })), s: 1, bad: 1 });
    week.rows.push({ n: 'Дней с рапортом', u: 'дн.', a: Object.keys(W1.reduce(function (o, r) { o[r.date] = 1; return o; }, {})).length, b: Object.keys(W0.reduce(function (o, r) { o[r.date] = 1; return o; }, {})).length, s: 1 });
    week.full = diff(date, first) >= 13;

    const short = String(c.name || pid).replace(/^г\.\s*Москва,\s*/i, '').split(' — ')[0];
    const kind = String(c.name || '').split(' — ').slice(1).join(' — ');
    const di = dates.length - 1;
    return {
      pid: pid, name: c.name || pid, short: short, kind: kind, light: light, reasons: reasons,
      pct: pct === null ? null : Math.round(pct * 10) / 10, due: dd.due, start: dd.start, forecast: forecast, first: first,
      last: { date: last.date, shift: last.shift, foreman: last.foreman || '' },
      today: !!today && (byPid[pid] || []).some(function (r) { return r.date === today; }),
      day: { people: S.people[di], mach: S.mach[di], mh: S.mh[di], m3: S.m3[di], trips: S.trips[di], dt: S.dt[di], idle: S.idle[di] },
      series: S, works: wlist.sort(function (a, b) { return b.act - a.act; }), week: week
    };
  }).filter(Boolean).sort(function (a, b) { return rank[b.light] - rank[a.light] || a.short.localeCompare(b.short, 'ru'); });

  const counts = { r: 0, y: 0, g: 0, n: 0 }; objects.forEach(function (o) { counts[o.light]++; });
  return { date: date, today: today || null, dates: dates, objects: objects, counts: counts };
}

function execModel_(date) {
  const all = readAll_();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return execFrom_(all.cfgs, all.reports, date, today);
}

/** HTML утреннего письма руководству. */
function execMailFrom_(m, page) {
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const ru = function (s) { return s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(0, 4); };
  const col = { r: '#d03b3b', y: '#e09a00', g: '#0ca30c', n: '#9aa898' };
  const lab = { r: 'КРАСНЫЙ', y: 'ЖЁЛТЫЙ', g: 'ЗЕЛЁНЫЙ', n: 'нет данных' };
  const dot = function (l) { return '<span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:' + col[l] + ';vertical-align:-2px"></span>'; };
  const rows = m.objects.map(function (o) {
    const bad = o.reasons.filter(function (x) { return x.l === 'r' || x.l === 'y'; });
    const show = bad.length ? bad : o.reasons.filter(function (x) { return x.k === 'due' && x.t; });
    return '<tr><td style="padding:10px 8px;border-top:1px solid #dde3da;vertical-align:top;width:22px">' + dot(o.light) + '</td>' +
      '<td style="padding:10px 8px;border-top:1px solid #dde3da;vertical-align:top"><b style="font-size:15px">' + esc(o.short) + '</b> <span style="font-size:11px;letter-spacing:.05em;color:' + (o.light === 'y' ? '#8a5a00' : col[o.light]) + '">' + lab[o.light] + '</span>' +
      (o.pct !== null ? ' <span style="color:#5a6758">· выполнено ' + String(Math.round(o.pct)) + '%</span>' : '') +
      '<div style="font-size:13px;color:#33402f;margin-top:3px">' + show.map(function (x) { return (x.l === 'r' || x.l === 'y' ? dot(x.l) + ' ' : '') + esc(x.t); }).join('<br>') + '</div></td></tr>';
  }).join('');
  const c = m.counts;
  const html = '<div style="font-family:Arial,sans-serif;color:#18221a;max-width:640px">' +
    '<table style="width:100%;border-collapse:collapse;background:#2E671F;color:#fff"><tr><td style="padding:16px 18px">' +
    '<div style="font-size:26px;font-weight:700;letter-spacing:.02em">ГК «КРАШМАШ»</div>' +
    '<div style="font-size:19px;margin-top:4px">Светофор по объектам за ' + ru(m.date) + '</div>' +
    '<div style="font-size:14px;margin-top:6px;opacity:.9">' + dot('r') + ' ' + c.r + ' &nbsp; ' + dot('y') + ' ' + c.y + ' &nbsp; ' + dot('g') + ' ' + c.g + (c.n ? ' &nbsp; ' + dot('n') + ' ' + c.n : '') + '</div>' +
    '</td></tr></table><table style="width:100%;border-collapse:collapse">' + rows + '</table>' +
    (page ? '<p style="margin:16px 0"><a href="' + esc(page) + '" style="background:#2E671F;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font-weight:600">Открыть панель с графиками</a></p>' : '') +
    '<p style="font-size:12px;color:#5a6758">Красный — нет рапортов 2 дня, простои от ' + EXEC_RULES.dtRed + ' ч за неделю, темп упал втрое, срок истёк или прогноз позже срока. Жёлтый — нет рапорта за день, простои от ' + EXEC_RULES.dtYellow + ' ч или по вине заказчика, темп упал на треть, прогноз впритык к сроку.</p></div>';
  const text = 'ГК «КРАШМАШ» — светофор по объектам за ' + ru(m.date) + '\nКрасных: ' + c.r + ', жёлтых: ' + c.y + ', зелёных: ' + c.g + '\n\n' +
    m.objects.map(function (o) { return '[' + lab[o.light] + '] ' + o.short + (o.pct !== null ? ' — ' + Math.round(o.pct) + '%' : '') + '\n' + o.reasons.filter(function (x) { return x.t && x.l !== 'g' && x.l !== 'n'; }).map(function (x) { return '  — ' + x.t; }).join('\n'); }).join('\n') +
    (page ? '\n\nПанель: ' + page : '');
  const subject = (c.r ? '🔴 ' + c.r + ' ' : '') + (c.y ? '🟡 ' + c.y + ' ' : '') + '🟢 ' + c.g + ' — объекты ГК КРАШМАШ за ' + ru(m.date);
  return { subject: subject, html: html, text: text };
}

/** Запустить один раз вручную: письмо руководству каждый день около 8:00 по Москве. */
function setupExec() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'execMorning') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('execMorning').timeBased().atHour(8).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  testExecMail();
}

/** Утреннее письмо: светофор за вчера. */
function execMorning() {
  if (!EXEC_TO) return;
  const y = new Date(); y.setDate(y.getDate() - 1);
  const d = execMailFrom_(execModel_(Utilities.formatDate(y, TZ, 'yyyy-MM-dd')), EXEC_PAGE);
  MailApp.sendEmail({ to: EXEC_TO, subject: d.subject, body: d.text, htmlBody: d.html, name: 'ГК КРАШМАШ — объекты' });
}

/** Проверка: отправить письмо светофора за вчера только первому адресу (себе). */
function testExecMail() {
  const y = new Date(); y.setDate(y.getDate() - 1);
  const d = execMailFrom_(execModel_(Utilities.formatDate(y, TZ, 'yyyy-MM-dd')), EXEC_PAGE);
  MailApp.sendEmail({ to: String(EXEC_TO || DIGEST_TO).split(',')[0].trim(), subject: '[проверка] ' + d.subject, body: d.text, htmlBody: d.html, name: 'ГК КРАШМАШ — объекты' });
}

/** Один раз после загрузки истории: дописать листы «Работы», «Вывоз», «Люди», «Техника», «Простои»
 *  для рапортов, которые есть только на листе «Рапорты» (внесены пакетом). Повторный запуск ничего не дублирует. */
function rebuildDetails() {
  const all = readAll_(), cfgBy = {};
  all.cfgs.forEach(function (c) { cfgBy[c.pid] = c; });
  const ids = {};
  ['W', 'H', 'P', 'M', 'D'].forEach(function (k) { const sh = sheet_(k), n = sh.getLastRow(); if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (v) { ids[String(v[0])] = 1; }); });
  let done = 0;
  all.reports.forEach(function (r) {
    if (!r || !r.id || ids[r.id]) return;
    const any = Object.keys(r.works || {}).length || (r.extra || []).length || (r.haul || []).length || (r.subs || []).length || (r.mach || []).length || (r.downtime || []).length || (r.people && r.people.own != null);
    if (!any) return;
    saveReport_(r, cfgBy[r.pid]); done++;
  });
  Logger.log('Дописано рапортов: ' + done);
}
