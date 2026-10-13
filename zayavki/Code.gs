/** Заявки на вывоз — ГК КРАШМАШ. Сервер для страницы заявок. Данные лежат в таблице SS_ID. */
var SS_ID = '1-mXmQEWOwusEQxlmTbUf6rKzsF9_VPeEdlrvXqus9jg';

var SPEC = {
  foremen:   {sheet:'Прорабы',      cols:[['id','s'],['name','s'],['objects','csv'],['token','s'],['active','b']]},
  staff:     {sheet:'Сотрудники',   cols:[['id','s'],['name','s'],['role','s'],['title','s'],['token','s']]},
  objects:   {sheet:'Объекты',      cols:[['id','s'],['name','s'],['short','s'],['dname','s']]},
  carriers:  {sheet:'Перевозчики',  cols:[['id','s'],['name','s']]},
  fleet:     {sheet:'Машины',       cols:[['id','s'],['name','s'],['model','s'],['kind','s'],['status','s'],['order','n'],['bort','s'],['cap','n']]},
  drivers:   {sheet:'Водители',    cols:[['id','s'],['name','s'],['active','b']]},
  routes:    {sheet:'Маршруты',    cols:[['id','s'],['from','s'],['to','s'],['km','n'],['price','n'],['cargo','s']]},
  places:    {sheet:'Места',       cols:[['id','s'],['name','s'],['type','s']]},
  requests:  {sheet:'Заявки',       cols:[['id','s'],['date','s'],['object','s'],['foremanId','s'],['foremanName','s'],['items','j'],['comment','s'],['sentAt','s'],['byStaff','s'],['editedAt','n'],['log','j'],['createdAt','n'],['summary','d']]},
  approvals: {sheet:'Утверждения',  cols:[['id','s'],['by','s'],['at','n'],['atStr','s']]}
};
var TYPES = {vyvoz:'Вывоз', mesto:'По месту', samo:'Самовывоз', polivo:'Поливомоечная'};

function doGet() {
  return ContentService.createTextOutput('Заявки на вывоз: сервер работает').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var out;
  try {
    var b = JSON.parse(e.postData.contents);
    var me = auth_(b.t);
    if (!me) out = {ok: false, error: 'auth'};
    else if (b.op === 'boot') out = boot_(me, b);
    else if (b.op === 'w') out = write_(me, b);
    else out = {ok: false, error: 'op'};
  } catch (err) {
    out = {ok: false, error: String(err && err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- чтение/запись листов ---------- */
function sheet_(coll) { return SpreadsheetApp.openById(SS_ID).getSheetByName(SPEC[coll].sheet); }

function fromCell_(type, v) {
  if (type === 'n') return v === '' ? undefined : Number(v);
  if (type === 'b') return /^(да|true|1|yes|истина)$/i.test(String(v).trim());
  if (type === 'csv') return v ? String(v).split(',').map(function (x) { return x.trim(); }).filter(String) : [];
  if (type === 'j') { if (!v) return undefined; try { return JSON.parse(v); } catch (e) { return undefined; } }
  return v;
}
function toCell_(type, v) {
  if (v === undefined || v === null) return '';
  if (type === 'b') return v === false ? 'Нет' : 'Да';
  if (type === 'csv') return (v || []).join(', ');
  if (type === 'j') return JSON.stringify(v);
  return String(v);
}
function readColl_(coll) {
  var sh = sheet_(coll), n = sh.getLastRow();
  if (n < 2) return [];
  var cols = SPEC[coll].cols, vals = sh.getRange(2, 1, n - 1, cols.length).getDisplayValues(), res = [];
  vals.forEach(function (row) {
    if (!row[0]) return;
    var d = {};
    cols.forEach(function (c, i) {
      if (c[1] === 'd') { d[c[0]] = row[i]; return; }
      var x = fromCell_(c[1], row[i]);
      if (x !== undefined) d[c[0]] = x;
    });
    res.push(d);
  });
  return res;
}
function findRow_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return -1;
  var ids = sh.getRange(1, 1, n, 1).getDisplayValues();
  for (var i = 1; i < n; i++) if (ids[i][0] === id) return i + 1;
  return -1;
}
function summary_(r) {
  return (r.items || []).map(function (it) {
    var s = (TYPES[it.type] || it.type) + (it.material ? ' ' + it.material : '');
    if (it.allOwn) s += ' · все свои'; else if (it.count) s += ' × ' + it.count;
    var a = it.assigned || {}, p = [];
    if (a.own && a.own.length) p.push('свои ' + a.own.join(','));
    (a.hired || []).forEach(function (h) { p.push(h.carrier + ' × ' + h.count); });
    if (a.agreed) p.push('согласовано ' + a.agreed);
    if (it.note) s += ' (' + it.note + ')';
    var f = it.fact;
    if (f) {
      var v = 0, tr = 0;
      (f.own || []).forEach(function (o) { v += Number(o.vol) || 0; tr += Number(o.trips) || 0; });
      (f.hired || []).forEach(function (h) { v += Number(h.vol) || 0; });
      p.push('ФАКТ ' + Math.round(v * 10) / 10 + ' м³, рейсов Рено ' + tr + (f.dest ? ', ' + f.dest : '') + (f.km ? ', ' + f.km + ' км' : ''));
    }
    return s + (p.length ? ' → ' + p.join('; ') : ' → не назначено');
  }).join(' | ');
}
function writeDoc_(coll, doc) {
  var sh = sheet_(coll), cols = SPEC[coll].cols;
  if (coll === 'requests') doc.summary = summary_(doc);
  var row = cols.map(function (c) { return toCell_(c[1], doc[c[0]]); });
  var r = findRow_(sh, doc.id);
  if (r < 0) r = Math.max(sh.getLastRow(), 1) + 1;
  var rng = sh.getRange(r, 1, 1, cols.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
}

/* ---------- доступ ---------- */
function tokenMap_() {
  var cache = CacheService.getScriptCache(), s = cache.get('tm');
  if (s) return JSON.parse(s);
  var m = {};
  readColl_('foremen').forEach(function (f) { if (f.token && f.active !== false) m[f.token] = {kind: 'foreman', id: f.id, name: f.name}; });
  readColl_('staff').forEach(function (f) { if (f.token) m[f.token] = {kind: f.role, id: f.id, name: f.name}; });
  cache.put('tm', JSON.stringify(m), 21600);
  return m;
}
function auth_(t) { if (!t || typeof t !== 'string') return null; return tokenMap_()[t] || null; }
function getVer_() { var c = CacheService.getScriptCache(), v = c.get('ver'); if (!v) { v = String(Date.now()); c.put('ver', v, 21600); } return v; }
function bump_() { var v = String(Date.now()) + Math.floor(Math.random() * 1000); CacheService.getScriptCache().put('ver', v, 21600); return v; }

function boot_(me, b) {
  var ver = getVer_();
  if (b.v && b.v === ver) return {ok: true, same: true, ver: ver};
  var data = {};
  Object.keys(SPEC).forEach(function (c) { data[c] = readColl_(c); });
  if (me.kind === 'foreman') data.requests = data.requests.filter(function (r) { return r.foremanId === me.id; });
  if (me.kind !== 'admin') { data.foremen.forEach(function (f) { delete f.token; }); data.staff.forEach(function (f) { delete f.token; }); }
  data.foremen = data.foremen.filter(function (f) { return me.kind === 'admin' || me.kind === 'dispatcher' ? true : f.active !== false; });
  return {ok: true, ver: ver, me: me, data: data};
}

function allowed_(me, kind, coll, id, data, existing) {
  var k = me.kind;
  if (k === 'admin') return true;
  if (k === 'director') return false;
  if (k === 'approver') return coll === 'approvals';
  if (k === 'dispatcher') return ['requests', 'carriers', 'fleet', 'foremen', 'objects', 'drivers', 'routes', 'places'].indexOf(coll) >= 0;
  if (k === 'foreman') {
    if (coll !== 'requests') return false;
    if (kind === 'set' && !existing) return data && data.foremanId === me.id;
    return !!existing && existing.foremanId === me.id && (kind === 'delete' || !data || !data.foremanId || data.foremanId === me.id);
  }
  return false;
}
function newToken_() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789', s = '';
  for (var i = 0; i < 18; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}

function write_(me, b) {
  var coll = b.coll, kind = b.kind, id = String(b.id || '');
  if (!SPEC[coll] || !id || ['set', 'update', 'delete'].indexOf(kind) < 0) return {ok: false, error: 'bad'};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(coll), r = findRow_(sh, id), existing = null;
    if (r > 0) { existing = readColl_(coll).filter(function (x) { return x.id === id; })[0] || null; }
    var data = b.data || {};
    if (!allowed_(me, kind, coll, id, data, existing)) return {ok: false, error: 'Нет прав на это действие'};
    if (coll === 'staff' && me.kind !== 'admin') return {ok: false, error: 'Нет прав'};
    if (kind === 'delete') {
      if (r > 0) sh.deleteRow(r);
    } else {
      var doc;
      if (kind === 'update') { if (!existing) return {ok: false, error: 'Запись не найдена'}; doc = existing; Object.keys(data).forEach(function (k) { doc[k] = data[k]; }); }
      else doc = JSON.parse(JSON.stringify(data));
      doc.id = id;
      if (coll === 'foremen' || coll === 'staff') {
        // личный код не меняется с клиента; новым выдаётся сам
        doc.token = (existing && existing.token) || newToken_();
        if (coll === 'staff' && me.kind !== 'admin') return {ok: false, error: 'Нет прав'};
      }
      if (coll === 'requests' && me.kind === 'foreman') { doc.foremanId = me.id; delete doc.byStaff; }
      writeDoc_(coll, doc);
    }
    if (coll === 'foremen' || coll === 'staff') CacheService.getScriptCache().remove('tm');
    return {ok: true, ver: bump_()};
  } finally {
    lock.releaseLock();
  }
}
