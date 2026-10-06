// Недельный отчёт по объекту из паспорта (cfg) и рапортов прорабов.
// weekHtml(cfg, reports, from, to, opts) -> HTML-фрагмент (стили внутри, класс .wk).
// Используется в кабинете руководителя (вкладка «Неделя») и для образцов.
(function (root) {
  const DOW = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const f = (v, d) => {
    if (v === null || v === undefined || v === "" || isNaN(v)) return "—";
    const p = Math.pow(10, d === undefined ? 1 : d); let s = String(Math.round(Number(v) * p) / p);
    const neg = s[0] === "-"; if (neg) s = s.slice(1);
    const sp = s.split("."); sp[0] = sp[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    return (neg ? "−" : "") + sp.join(",");
  };
  const ru = s => s.slice(8, 10) + "." + s.slice(5, 7);
  const ruY = s => s.slice(8, 10) + "." + s.slice(5, 7) + "." + s.slice(0, 4);
  const addD = (iso, n) => { const d = new Date(iso + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const dow = iso => DOW[new Date(iso + "T12:00:00Z").getUTCDay()];
  const nrm = s => String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
  const isMach = w => String(w.src || "").indexOf("mach:") === 0;
  const num = v => (typeof v === "number" && !isNaN(v) ? v : 0);

  function weekHtml(cfg, reports, from, to, opts) {
    opts = opts || {};
    const hl = cfg.hl || {}, haulT = hl.t || "Вывоз", carT = hl.c || "Перевозчик";
    const works = cfg.works || [];
    const days = []; for (let d = from; d <= to; d = addD(d, 1)) days.push(d);
    const rs = (reports || []).filter(r => r && r.pid === cfg.pid);
    const inP = rs.filter(r => r.date >= from && r.date <= to);
    const byDay = {}; inP.forEach(r => (byDay[r.date] = byDay[r.date] || []).push(r));
    const d0date = cfg.d0date || "";
    // накопительно на конец даты X: d0 + рапорты после d0date и до X включительно
    const cumAt = (code, x) => {
      const w = works.find(z => z.c === code) || {};
      return num(w.d0) + rs.filter(r => r.date > d0date && r.date <= x).reduce((a, r) => a + num(r.works && r.works[code]), 0);
    };
    const wSum = (code, list) => list.reduce((a, r) => a + num(r.works && r.works[code]), 0);
    const before = addD(from, -1);
    const repDays = days.filter(d => byDay[d]);
    const missDays = days.filter(d => !byDay[d] && d <= (opts.today || to));

    // ---- работы: было → стало
    const rows = works.filter(w => !isMach(w)).map(w => {
      const was = cumAt(w.c, before), wk = wSum(w.c, inP), now = was + wk;
      return { w, was, wk, now, left: w.t ? w.t - now : null, pct: w.t ? now / w.t * 100 : null };
    }).filter(x => x.wk || x.now || x.w.t);
    const extra = {}; inP.forEach(r => (r.extra || []).forEach(e => { const k = nrm(e.n) + "|" + nrm(e.u); extra[k] = extra[k] || { n: e.n, u: e.u, v: 0 }; extra[k].v += num(e.v); }));

    // ---- вывоз
    const hm = {}, hc = {}; let hv = 0, ht = 0, hw = 0;
    inP.forEach(r => (r.haul || []).forEach(h => {
      const m = h.m || "—", c = h.c || "—";
      hm[m] = hm[m] || { v: 0, t: 0, w: 0 }; hm[m].v += num(h.v); hm[m].t += num(h.t); hm[m].w += num(h.w);
      const k = c + "|" + m; hc[k] = hc[k] || { c, m, v: 0, t: 0, w: 0 }; hc[k].v += num(h.v); hc[k].t += num(h.t); hc[k].w += num(h.w);
      hv += num(h.v); ht += num(h.t); hw += num(h.w);
    }));
    const haulDay = d => (byDay[d] || []).reduce((a, r) => a + (r.haul || []).reduce((b, h) => b + num(h.v), 0), 0);

    // ---- люди
    const pplDay = d => { const l = byDay[d]; if (!l) return null; let n = 0, any = false; l.forEach(r => { if (r.people && r.people.own != null) { any = true; n += num(r.people.own); } (r.subs || []).forEach(s => { any = true; n += num(s.p); }); }); return any ? n : null; };
    const orgs = {}; inP.forEach(r => { if (r.people && r.people.own != null) { orgs["Свои"] = orgs["Свои"] || {}; orgs["Свои"][r.date] = (orgs["Свои"][r.date] || 0) + num(r.people.own); } (r.subs || []).forEach(s => { const n = s.n || "—"; orgs[n] = orgs[n] || {}; orgs[n][r.date] = (orgs[n][r.date] || 0) + num(s.p); }); });
    const manDays = days.reduce((a, d) => a + (pplDay(d) || 0), 0);

    // ---- техника
    const units = {}; inP.forEach(r => (r.mach || []).forEach(m => { const k = nrm(m.n); units[k] = units[k] || { n: m.n, d: {} }; if (m.h != null) units[k].d[r.date] = (units[k].d[r.date] || 0) + num(m.h); else if (units[k].d[r.date] == null) units[k].d[r.date] = null; }));
    const machTotal = Object.keys(units).reduce((a, k) => a + Object.values(units[k].d).reduce((b, v) => b + num(v), 0), 0);
    const machCum = n => { const w = works.find(z => isMach(z) && z.src.slice(5).split("|").map(nrm).length === 1 && nrm(z.src.slice(5)) === nrm(n)); return w ? cumAt(w.c, to) : null; };
    const techDay = d => { let s = 0, any = false; (byDay[d] || []).forEach(r => (r.mach || []).forEach(m => { if (m.h != null) { any = true; s += num(m.h); } })); return any ? s : null; };

    // ---- простои, примечания
    const dts = []; inP.forEach(r => (r.downtime || []).forEach(x => dts.push({ d: r.date, ...x })));
    const notes = inP.filter(r => r.other).map(r => ({ d: r.date, t: r.other, who: r.foreman }));
    const dtH = dts.reduce((a, x) => a + num(x.h), 0);

    // ---- автоматический контроль
    const ctl = [];
    if (missDays.length) ctl.push(["warn", "Нет рапорта: " + missDays.map(d => ru(d) + " (" + dow(d) + ")").join(", ") + ". Эти дни не учтены и не приравнены к нулю."]);
    rows.filter(x => x.w.t && x.now > x.w.t * 1.0001).forEach(x => ctl.push(["bad", x.w.n + ": выполнено " + f(x.now) + " " + x.w.u + " — больше плана " + f(x.w.t) + ". Проверьте объём или паспорт."]));
    rows.filter(x => x.w.t && x.left > 0 && !x.wk && !x.w.ro).slice(0, 6).forEach(x => ctl.push(["info", x.w.n + ": за неделю без движения, остаток " + f(x.left) + " " + x.w.u + "."]));
    Object.keys(units).forEach(k => { const u = units[k], v = Object.values(u.d); if (v.length && v.every(h => !num(h))) ctl.push(["info", u.n + ": 0 маш.-ч всю неделю (простой или не указаны часы)."]); });
    if (!ctl.length) ctl.push(["ok", "Замечаний нет: рапорты за все дни, объёмы в пределах плана."]);

    // ---- главный объём для графика: вывоз, если есть; иначе первая ручная работа с объёмом
    let chart = null;
    if (hv) chart = { t: haulT + " по дням, м³", u: "м³", val: d => byDay[d] ? haulDay(d) : null };
    else { const m = rows.find(x => x.wk && !x.w.src); if (m) chart = { t: m.w.n + " по дням, " + m.w.u, u: m.w.u, val: d => byDay[d] ? wSum(m.w.c, byDay[d]) : null }; }

    // ---- KPI
    const kpi = [
      ["Рапорты", repDays.length + " из " + days.length, missDays.length ? "дней без рапорта: " + missDays.length : "все дни закрыты", missDays.length ? "warn" : "ok"],
      [haulT, f(hv) + " м³", (ht ? f(ht, 0) + " рейс." : "") + (hw ? (ht ? " · " : "") + f(hw) + " т" : "") || "за неделю", ""],
      ["Техника", f(machTotal) + " маш.-ч", Object.keys(units).length + " ед. в рапортах", ""],
      ["Люди", manDays ? f(manDays, 0) + " чел.-см." : "не указаны", manDays && repDays.length ? "в среднем " + f(manDays / repDays.length) + " чел./смену" : "", ""],
    ];
    if (dtH) kpi.push(["Простои", f(dtH) + " ч", dts.length + " случ.", "bad"]);

    // ================= HTML =================
    const H = [];
    H.push('<div class="wk">');
    H.push('<header class="wk-hd"><div><div class="wk-sup">ГК «КРАШМАШ» · недельный отчёт</div><h1>' + esc(cfg.name || cfg.pid) + '</h1><div class="wk-per">' + ruY(from) + " – " + ruY(to) + (d0date ? ' · выполнено ранее учтено на ' + ruY(d0date) : "") + "</div></div>" + (opts.logo ? '<img src="' + opts.logo + '" alt="ГК КРАШМАШ" class="wk-logo">' : "") + "</header>");
    if (opts.banner) H.push('<div class="wk-banner">' + opts.banner + "</div>");
    H.push('<section class="wk-kpis">' + kpi.map(k => '<div class="wk-kpi ' + k[3] + '"><small>' + esc(k[0]) + "</small><strong>" + esc(k[1]) + "</strong><span>" + esc(k[2]) + "</span></div>").join("") + "</section>");

    // контроль
    H.push('<section class="wk-card"><h2>Контроль данных</h2><ul class="wk-ctl">' + ctl.map(c => '<li class="' + c[0] + '"><b aria-hidden="true">' + ({ warn: "!", bad: "×", info: "i", ok: "✓" }[c[0]]) + "</b>" + esc(c[1]) + "</li>").join("") + "</ul></section>");

    // было → стало
    if (rows.length || Object.keys(extra).length) {
      H.push('<section class="wk-card"><h2>Объёмы: было → стало</h2><div class="wk-scroll"><table><thead><tr><th>Работа</th><th>Ед.</th><th class="n">План</th><th class="n">Было на ' + ru(before) + '</th><th class="n">За неделю</th><th class="n">Стало на ' + ru(to) + '</th><th class="n">Остаток</th><th>Готовность</th></tr></thead><tbody>');
      rows.forEach(x => {
        const p = x.pct == null ? null : Math.max(0, Math.min(100, x.pct));
        H.push('<tr' + (x.w.ro ? ' class="ro"' : "") + "><td>" + esc(x.w.n) + (x.w.src ? '<small>считается из ' + (String(x.w.src).indexOf("trips:") === 0 ? "рейсов" : String(x.w.src).indexOf("tons:") === 0 ? "тонн" : "раздела «" + esc(haulT) + "»") + "</small>" : "") + "</td><td>" + esc(x.w.u) + '</td><td class="n">' + (x.w.t ? f(x.w.t) : "—") + '</td><td class="n">' + f(x.was) + '</td><td class="n"><b>' + (x.wk ? "+" + f(x.wk) : "0") + '</b></td><td class="n">' + f(x.now) + '</td><td class="n">' + (x.left == null ? "—" : f(x.left)) + "</td><td>" + (p == null ? '<span class="mut">план не задан</span>' : '<div class="wk-bar' + (x.pct > 100.01 ? " over" : "") + '" title="' + f(x.pct) + '%"><i style="width:' + p.toFixed(1) + '%"></i></div><span class="wk-pc">' + f(x.pct) + "%</span>") + "</td></tr>");
      });
      Object.keys(extra).forEach(k => H.push('<tr class="ex"><td>' + esc(extra[k].n) + "<small>вне бюджета</small></td><td>" + esc(extra[k].u) + '</td><td class="n">—</td><td class="n">—</td><td class="n"><b>+' + f(extra[k].v) + '</b></td><td class="n">—</td><td class="n">—</td><td><span class="mut">доп. работа</span></td></tr>'));
      H.push("</tbody></table></div></section>");
    }

    // график
    if (chart) {
      const vals = days.map(chart.val), mx = Math.max(1, ...vals.filter(v => v != null)), have = vals.filter(v => v != null);
      const avg = have.length ? have.reduce((a, b) => a + b, 0) / have.length : 0;
      const W = 720, Hh = 230, pl = 44, pr = 12, pt = 26, pb = 40, cw = (W - pl - pr) / days.length, bw = Math.min(46, cw * 0.56), sc = v => (Hh - pb) - (v / topV()) * (Hh - pb - pt), topV = () => top;
      const nice = v => { const e = Math.pow(10, Math.floor(Math.log10(v))), m = v / e; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * e; };
      const step = nice(mx / 3), top = step * Math.ceil(mx * 1.08 / step), ticks = []; for (let t = 0; t <= top + 1e-9; t += step) ticks.push(t);
      let s = '<svg viewBox="0 0 ' + W + " " + Hh + '" class="wk-svg" role="img" aria-label="' + esc(chart.t) + '">';
      ticks.forEach(t => { const y = sc(t); s += '<line x1="' + pl + '" x2="' + (W - pr) + '" y1="' + y + '" y2="' + y + '" class="gr"/><text x="' + (pl - 6) + '" y="' + (y + 4) + '" class="ax" text-anchor="end">' + f(t, 0) + "</text>"; });
      days.forEach((d, i) => {
        const x = pl + cw * i + (cw - bw) / 2, v = vals[i];
        if (v == null) s += '<rect x="' + x + '" y="' + (sc(0) - 22) + '" width="' + bw + '" height="22" rx="4" class="na"><title>' + ru(d) + ": нет рапорта</title></rect><text x=\"" + (x + bw / 2) + '" y="' + (sc(0) - 28) + '" class="nal" text-anchor="middle">нет</text>';
        else if (v > 0) { const y = sc(v); s += '<path d="M' + x + "," + sc(0) + "V" + (y + 4) + "q0,-4 4,-4h" + (bw - 8) + "q4,0 4,4V" + sc(0) + 'Z" class="bar"><title>' + ru(d) + ": " + f(v) + " " + esc(chart.u) + "</title></path>" + '<text x="' + (x + bw / 2) + '" y="' + (y - 6) + '" class="vl" text-anchor="middle">' + f(v, 0) + "</text>"; }
        else s += '<text x="' + (x + bw / 2) + '" y="' + (sc(0) - 6) + '" class="vl" text-anchor="middle">0</text>';
        s += '<text x="' + (x + bw / 2) + '" y="' + (Hh - pb + 16) + '" class="ax" text-anchor="middle">' + ru(d) + '</text><text x="' + (x + bw / 2) + '" y="' + (Hh - pb + 30) + '" class="ax dw" text-anchor="middle">' + dow(d) + "</text>";
      });
      if (avg) { const y = sc(avg); s += '<line x1="' + pl + '" x2="' + (W - pr) + '" y1="' + y + '" y2="' + y + '" class="avg"/><text x="' + (pl + 4) + '" y="' + (y - 5) + '" class="avl" text-anchor="start">среднее по рапортам ' + f(avg) + "</text>"; }
      s += "</svg>";
      H.push('<section class="wk-card"><h2>' + esc(chart.t) + "</h2>" + s + "</section>");
    }

    // по дням
    const dayWorks = works.filter(w => !isMach(w) && inP.some(r => num(r.works && r.works[w.c]))).filter(w => !/^в т\.ч\./i.test(w.n)).slice(0, 6);
    H.push('<section class="wk-card"><h2>По дням</h2><div class="wk-scroll"><table class="days"><thead><tr><th>Дата</th>' + dayWorks.map(w => '<th class="n">' + esc(w.n.length > 40 ? w.n.slice(0, 38) + "…" : w.n) + ", " + esc(w.u) + "</th>").join("") + '<th class="n">' + esc(haulT) + ', м³</th><th class="n">Люди</th><th class="n">Техника, ч</th><th>Прораб</th></tr></thead><tbody>');
    days.forEach(d => {
      const l = byDay[d];
      if (!l) { H.push('<tr class="miss"><td>' + ru(d) + " " + dow(d) + '</td><td colspan="' + (dayWorks.length + 4) + '">нет рапорта</td></tr>'); return; }
      const p = pplDay(d), t = techDay(d);
      H.push("<tr><td>" + ru(d) + " " + dow(d) + "</td>" + dayWorks.map(w => '<td class="n">' + (wSum(w.c, l) ? f(wSum(w.c, l)) : '<span class="mut">0</span>') + "</td>").join("") + '<td class="n">' + (haulDay(d) ? f(haulDay(d)) : '<span class="mut">0</span>') + '</td><td class="n">' + (p == null ? '<span class="mut">не указ.</span>' : f(p, 0)) + '</td><td class="n">' + (t == null ? "—" : f(t)) + "</td><td>" + esc(l.map(r => (r.shift === "N" ? "ночь: " : "") + (r.foreman || "?")).join("; ")) + "</td></tr>");
    });
    H.push('<tr class="tot"><td>Итого</td>' + dayWorks.map(w => '<td class="n">' + f(wSum(w.c, inP)) + "</td>").join("") + '<td class="n">' + f(hv) + '</td><td class="n">' + (manDays ? f(manDays, 0) + " чел.-см." : "—") + '</td><td class="n">' + f(machTotal) + "</td><td></td></tr>");
    H.push("</tbody></table></div></section>");

    // техника
    const uk = Object.keys(units);
    if (uk.length) {
      const mxh = Math.max(1, ...uk.map(k => Math.max(0, ...Object.values(units[k].d).map(num))));
      H.push('<section class="wk-card"><h2>Техника, маш.-ч</h2><div class="wk-scroll"><table class="heat"><thead><tr><th>Единица</th>' + days.map(d => '<th class="n">' + ru(d) + "</th>").join("") + '<th class="n">Неделя</th><th class="n">С начала</th></tr></thead><tbody>');
      uk.forEach(k => {
        const u = units[k], tot = Object.values(u.d).reduce((a, v) => a + num(v), 0), c = machCum(u.n);
        H.push("<tr><td>" + esc(u.n) + "</td>" + days.map(d => { if (!byDay[d]) return '<td class="n na">—</td>'; const v = u.d[d]; if (v === undefined) return '<td class="n mut">·</td>'; if (v === null) return '<td class="n mut" title="часы не указаны">?</td>'; const a = num(v) / mxh; return '<td class="n" style="--h:' + a.toFixed(2) + '"><span>' + f(v) + "</span></td>"; }).join("") + '<td class="n"><b>' + f(tot) + '</b></td><td class="n">' + (c == null ? '<span class="mut">—</span>' : f(c)) + "</td></tr>");
      });
      H.push('</tbody></table></div><p class="wk-note">«—» — нет рапорта за день, «·» — единицы не было в рапорте, «?» — в рапорте без часов. «С начала» — для техники, которая есть в паспорте объекта.</p></section>');
    }

    // вывоз по перевозчикам
    const hk = Object.keys(hc);
    if (hk.length) {
      H.push('<section class="wk-card"><h2>' + esc(haulT) + " по материалам и " + (hl.c ? "покупателям" : "перевозчикам") + '</h2><div class="wk-scroll"><table><thead><tr><th>' + esc(carT) + "</th><th>Материал</th>" + (ht ? '<th class="n">Рейсов</th>' : "") + '<th class="n">м³</th>' + (hw ? '<th class="n">т</th>' : "") + "</tr></thead><tbody>");
      hk.sort((a, b) => hc[b].v - hc[a].v).forEach(k => { const x = hc[k]; H.push("<tr><td>" + esc(x.c) + "</td><td>" + esc(x.m) + "</td>" + (ht ? '<td class="n">' + f(x.t, 0) + "</td>" : "") + '<td class="n">' + f(x.v) + "</td>" + (hw ? '<td class="n">' + f(x.w) + "</td>" : "") + "</tr>"); });
      H.push('<tr class="tot"><td>Итого</td><td></td>' + (ht ? '<td class="n">' + f(ht, 0) + "</td>" : "") + '<td class="n">' + f(hv) + "</td>" + (hw ? '<td class="n">' + f(hw) + "</td>" : "") + "</tr></tbody></table></div></section>");
    }

    // люди по организациям
    const ok = Object.keys(orgs);
    if (ok.length) {
      H.push('<section class="wk-card"><h2>Люди по организациям</h2><div class="wk-scroll"><table><thead><tr><th>Организация</th>' + days.map(d => '<th class="n">' + ru(d) + "</th>").join("") + '<th class="n">Чел.-смен</th></tr></thead><tbody>');
      ok.forEach(o => H.push("<tr><td>" + esc(o) + "</td>" + days.map(d => '<td class="n">' + (!byDay[d] ? '<span class="mut">—</span>' : orgs[o][d] == null ? '<span class="mut">·</span>' : f(orgs[o][d], 0)) + "</td>").join("") + '<td class="n"><b>' + f(Object.values(orgs[o]).reduce((a, b) => a + b, 0), 0) + "</b></td></tr>"));
      H.push("</tbody></table></div></section>");
    }

    if (dts.length) H.push('<section class="wk-card"><h2>Простои — ' + f(dtH) + ' ч</h2><ul class="wk-list">' + dts.map(x => "<li><b>" + ru(x.d) + "</b> " + esc(x.what || "") + " — " + f(x.h) + " ч · " + esc(x.why || "") + ' <span class="mut">(вина: ' + esc(x.fault || "?") + ")</span></li>").join("") + "</ul></section>");
    if (notes.length) H.push('<section class="wk-card"><h2>Примечания прорабов</h2><ul class="wk-list">' + notes.map(x => "<li><b>" + ru(x.d) + "</b> " + esc(x.t) + "</li>").join("") + "</ul></section>");

    H.push('<p class="wk-foot">Собрано автоматически из рапортов прорабов' + (opts.built ? " · " + esc(opts.built) : "") + ". Дни без рапорта не приравниваются к нулю.</p></div>");
    return H.join("");
  }

  const WEEK_CSS = `
.wk{--g:#2E671F;--g2:#e7efe3;--ink:#1d1d1b;--mut:#6b6f66;--line:#dde3d8;--card:#fff;--bg:#f3f5f1;--warn:#9a5b00;--warnbg:#fdf3e1;--bad:#a1271c;--badbg:#fdecea;--info:#2f5e8c;--infobg:#e9f0f7;color:var(--ink);font:15px/1.45 Arial,Helvetica,sans-serif;max-width:1040px;margin:0 auto}
.wk *{box-sizing:border-box}
.wk-hd{background:var(--g);color:#fff;border-radius:12px;padding:18px 20px;display:flex;justify-content:space-between;gap:16px;align-items:flex-start}
.wk-hd h1{margin:4px 0 6px;font-size:24px;line-height:1.2}.wk-sup{font-size:13px;opacity:.85}.wk-per{font-size:14px;opacity:.95}
.wk-logo{height:44px;border-radius:4px;flex:none}
.wk-banner{margin:12px 0 0;padding:10px 14px;border-radius:10px;background:var(--infobg);color:var(--info);font-size:14px}
.wk-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:12px 0}
.wk-kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.wk-kpi small{display:block;color:var(--mut);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
.wk-kpi strong{display:block;font-size:22px;margin:4px 0 2px}.wk-kpi span{color:var(--mut);font-size:13px}
.wk-kpi.warn{border-left:4px solid var(--warn)}.wk-kpi.ok{border-left:4px solid var(--g)}.wk-kpi.bad{border-left:4px solid var(--bad)}
.wk-card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin:0 0 12px}
.wk-card h2{margin:0 0 10px;font-size:17px;color:var(--g)}
.wk-scroll{overflow-x:auto}
.wk table{width:100%;border-collapse:collapse;font-size:14px}
.wk th{font-size:12px;color:var(--mut);font-weight:600;text-align:left;padding:7px 8px;border-bottom:2px solid var(--line);white-space:nowrap;background:#fafbf9}
.wk td{padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}
.wk td small{display:block;color:var(--mut);font-size:12px}
.wk .n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.wk .mut{color:var(--mut)}
.wk table.days th{white-space:normal;vertical-align:bottom;min-width:64px}.wk table.days td:first-child{white-space:nowrap}
.wk tr.tot td{font-weight:700;background:#f6f8f4;border-top:2px solid var(--line)}
.wk tr.miss td{color:var(--mut);font-style:italic;background:repeating-linear-gradient(135deg,#fafafa 0 6px,#f1f1ef 6px 12px)}
.wk tr.ro td{color:var(--mut)}.wk tr.ex td:first-child{border-left:3px solid var(--warn)}
.wk-bar{display:inline-block;width:90px;height:8px;background:#e6e9e3;border-radius:4px;overflow:hidden;vertical-align:middle;margin-right:6px}
.wk-bar i{display:block;height:100%;background:var(--g);border-radius:4px}.wk-bar.over i{background:var(--bad)}
.wk-pc{font-variant-numeric:tabular-nums;font-size:13px}
.wk-ctl{list-style:none;margin:0;padding:0}.wk-ctl li{display:flex;gap:10px;align-items:flex-start;padding:8px 10px;border-radius:8px;margin-bottom:6px;font-size:14px}
.wk-ctl b{flex:none;width:20px;height:20px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:12px;color:#fff}
.wk-ctl .warn{background:var(--warnbg)}.wk-ctl .warn b{background:var(--warn)}
.wk-ctl .bad{background:var(--badbg)}.wk-ctl .bad b{background:var(--bad)}
.wk-ctl .info{background:var(--infobg)}.wk-ctl .info b{background:var(--info)}
.wk-ctl .ok{background:var(--g2)}.wk-ctl .ok b{background:var(--g)}
.wk-svg{width:100%;height:auto;display:block}
.wk-svg .gr{stroke:var(--line);stroke-width:1}.wk-svg .ax{fill:var(--mut);font-size:12px}.wk-svg .dw{font-size:11px}
.wk-svg .bar{fill:var(--g)}.wk-svg .bar:hover{opacity:.8}.wk-svg .vl{fill:var(--ink);font-size:12px;font-weight:600}
.wk-svg .na{fill:#eceeea;stroke:#c9cec4;stroke-dasharray:3 3}.wk-svg .nal{fill:var(--mut);font-size:11px;font-style:italic}
.wk-svg .avg{stroke:var(--warn);stroke-width:1.5;stroke-dasharray:5 4}.wk-svg .avl{fill:var(--warn);font-size:12px}
.wk table.heat td[style]{background:rgba(46,103,31,calc(var(--h)*.55 + .06))}
.wk table.heat td[style] span{font-weight:600}.wk table.heat td.na{background:repeating-linear-gradient(135deg,#fafafa 0 6px,#f1f1ef 6px 12px);color:var(--mut)}
.wk-note,.wk-foot{color:var(--mut);font-size:12px;margin:8px 0 0}
.wk-list{margin:0;padding-left:18px}.wk-list li{margin:4px 0}
@media (max-width:640px){.wk-hd{flex-direction:column}.wk-hd h1{font-size:20px}.wk-logo{height:36px}}
@media print{.wk-card,.wk-kpi{break-inside:avoid}.wk-hd{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
`;
  const api = { weekHtml, WEEK_CSS };
  if (typeof module !== "undefined") module.exports = api; else root.RaportWeek = api;
})(this);
