/* Вход по логину и паролю для страниц «Рапорт прораба» (Руководитель проектов, Светофор, Доступы).
   Подключается одной строкой <script src="auth.js"></script> в начале <head>.
   Страницы продолжают читать адрес базы и ключ из адресной строки (#u=…&k=…): после входа сюда подставляется k=tk:ТОКЕН.
   Если в ссылке старый ключ (без tk:), страница работает как раньше. */
(function () {
  "use strict";
  var LS = "raport2-auth";
  var hs = new URLSearchParams(location.hash.slice(1));
  var U = hs.get("u") || "", K = hs.get("k") || "";
  function ld() { try { return JSON.parse(localStorage.getItem(LS) || "null"); } catch (e) { return null; } }
  function sv(o) { try { localStorage.setItem(LS, JSON.stringify(o)); } catch (e) {} }
  var st = ld() || {};
  if (!U && st.u) U = st.u;
  var origFetch = window.fetch ? window.fetch.bind(window) : null;
  function setHash(k) {
    var p = new URLSearchParams(location.hash.slice(1));
    if (U) p.set("u", U);
    if (k) p.set("k", k); else p.delete("k");
    try { history.replaceState(null, "", location.pathname + location.search + "#" + p.toString()); } catch (e) {}
  }
  function logout(reload) {
    st.t = ""; st.name = ""; st.role = ""; sv(st);
    setHash("");
    if (reload !== false) location.reload();
  }
  var legacy = !!K && K.indexOf("tk:") !== 0;
  var haveToken = false;
  if (!legacy) {
    if (K.indexOf("tk:") === 0) { st.t = K.slice(3); st.u = U; sv(st); haveToken = true; }
    else if (st.t && U && st.u === U) { setHash("tk:" + st.t); haveToken = true; }
  }
  window.RaportAuth = {
    u: function () { return U; },
    token: function () { return haveToken ? st.t : ""; },
    name: function () { return st.name || ""; },
    role: function () { return st.role || ""; },
    legacy: legacy,
    logout: logout
  };

  // токен просрочен или отозван — вернуть на экран входа
  if (origFetch && !legacy) {
    window.fetch = function () {
      var a = arguments, url = String(a[0] && a[0].url || a[0]);
      return origFetch.apply(window, a).then(function (r) {
        try {
          if (U && url.indexOf(U.split("?")[0]) === 0) {
            r.clone().json().then(function (j) { if (j && j.error === "auth" && st.t) logout(); }).catch(function () {});
          }
        } catch (e) {}
        return r;
      });
    };
  }
  if (legacy || haveToken) {
    if (haveToken) {
      document.addEventListener("DOMContentLoaded", function () {
        var b = document.createElement("button");
        b.type = "button"; b.id = "auth-out";
        b.textContent = (st.name ? st.name + " · " : "") + "Выйти";
        b.setAttribute("style", "position:fixed;right:8px;bottom:6px;z-index:40;font:12px/1 Arial,sans-serif;padding:6px 10px;border-radius:14px;border:1px solid rgba(128,128,128,.4);background:rgba(255,255,255,.88);color:#333;cursor:pointer");
        b.onclick = function () { logout(); };
        document.body.appendChild(b);
      });
    }
    return;
  }

  /* ---------- экран входа ---------- */
  var root = document.documentElement;
  var lock = document.createElement("style");
  lock.textContent = "html.auth-lock body>*:not(#auth-ov){display:none!important}#auth-ov{position:fixed;inset:0;z-index:100;display:flex;align-items:center;justify-content:center;padding:16px;background:#2E671F;font-family:Arial,sans-serif}#auth-ov .c{background:#fff;color:#1d1d1b;border-radius:14px;max-width:380px;width:100%;padding:22px 20px}#auth-ov h2{margin:0 0 4px;font-size:22px}#auth-ov p{margin:0 0 14px;color:#555;font-size:14px;line-height:1.4}#auth-ov label{display:block;font-size:13px;color:#555;margin:10px 0 4px}#auth-ov input{width:100%;box-sizing:border-box;font-size:17px;padding:11px;border:1px solid #bbb;border-radius:8px;background:#fff;color:#111}#auth-ov button{width:100%;margin-top:16px;min-height:48px;font-size:17px;font-weight:700;border:0;border-radius:10px;background:#2E671F;color:#fff;cursor:pointer}#auth-ov button:disabled{opacity:.6}#auth-ov .e{color:#b3362c;font-size:14px;margin-top:10px;min-height:18px}";
  document.head.appendChild(lock);
  root.classList.add("auth-lock");

  function post(body) {
    return origFetch(U, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(body), redirect: "follow" }).then(function (r) { return r.json(); });
  }
  function mk(tag, txt, attrs) {
    var e = document.createElement(tag);
    if (txt != null) e.textContent = txt;
    if (attrs) Object.keys(attrs).forEach(function (k) { e.setAttribute(k, attrs[k]); });
    return e;
  }
  function finish(j) {
    st = { u: U, t: j.token, name: j.name, role: j.role }; sv(st);
    setHash("tk:" + j.token);
    setTimeout(function () { location.reload(); }, 600);
  }
  function showLogin(box) {
    box.innerHTML = "";
    box.appendChild(mk("h2", "Вход"));
    box.appendChild(mk("p", "ГК «КРАШМАШ». Введите логин и пароль, которые выдал администратор."));
    if (!U) { box.appendChild(mk("p", "Откройте страницу по ссылке, которую вам прислали (в ней адрес базы).")); return; }
    var l1 = mk("label", "Логин"), i1 = mk("input", null, { type: "text", autocomplete: "username", autocapitalize: "none", autocorrect: "off", spellcheck: "false" });
    var l2 = mk("label", "Пароль"), i2 = mk("input", null, { type: "password", autocomplete: "current-password" });
    i1.setAttribute("name", "username"); i2.setAttribute("name", "password");
    var go = mk("button", "Войти", { type: "submit" }), er = mk("div", "", {}); er.className = "e";
    var fm = mk("form", null, { autocomplete: "on" }); box.appendChild(fm);
    [l1, i1, l2, i2, go, er].forEach(function (x) { fm.appendChild(x); });
    fm.addEventListener("submit", function (e) { e.preventDefault(); submit(); });
    function submit() {
      if (!i1.value.trim() || !i2.value) { er.textContent = "Введите логин и пароль"; return; }
      go.disabled = true; er.textContent = "";
      post({ a: "login", u: i1.value.trim(), p: i2.value }).then(function (j) {
        go.disabled = false;
        if (j.ok && j.must) { showChange(box, j, i1.value.trim()); return; }
        if (j.ok) { finish(j); return; }
        er.textContent = j.error === "locked" ? "Слишком много попыток. Подождите 15 минут." : j.error === "bad" ? "Неверный логин или пароль" : "Ошибка: " + (j.error || "нет связи");
      }).catch(function () { go.disabled = false; er.textContent = "Нет связи. Проверьте интернет и повторите."; });
    }
    i1.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); i2.focus(); } });
    setTimeout(function () { try { i1.focus(); } catch (e) {} }, 50);
  }
  function showChange(box, j, lg) {
    box.innerHTML = "";
    box.appendChild(mk("h2", "Задайте свой пароль"));
    box.appendChild(mk("p", "Вы вошли с временным паролем. Придумайте свой — не короче 8 символов. Дальше входить нужно будет с ним."));
    var l1 = mk("label", "Новый пароль"), i1 = mk("input", null, { type: "password", autocomplete: "new-password" });
    var l2 = mk("label", "Повторите пароль"), i2 = mk("input", null, { type: "password", autocomplete: "new-password" });
    var un = mk("input", null, { type: "text", name: "username", autocomplete: "username", value: lg || j.login || "" }); un.style.display = "none";
    var go = mk("button", "Сохранить и войти", { type: "submit" }), er = mk("div", ""); er.className = "e";
    var fm = mk("form", null, { autocomplete: "on" }); box.appendChild(fm);
    [un, l1, i1, l2, i2, go, er].forEach(function (x) { fm.appendChild(x); });
    fm.addEventListener("submit", function (e) { e.preventDefault(); submit(); });
    function submit() {
      if (i1.value.length < 8) { er.textContent = "Пароль короче 8 символов"; return; }
      if (i1.value !== i2.value) { er.textContent = "Пароли не совпадают"; return; }
      go.disabled = true; er.textContent = "";
      post({ a: "chpass", t: j.token, p: i1.value }).then(function (r) {
        go.disabled = false;
        if (r.ok) { finish(r); return; }
        er.textContent = r.error === "short" ? "Пароль короче 8 символов" : r.error === "same" ? "Новый пароль должен отличаться от временного" : "Ошибка: " + (r.error || "нет связи");
      }).catch(function () { go.disabled = false; er.textContent = "Нет связи. Повторите."; });
    }
    setTimeout(function () { try { i1.focus(); } catch (e) {} }, 50);
  }
  document.addEventListener("DOMContentLoaded", function () {
    var ov = mk("div", null, { id: "auth-ov" }), box = mk("div"); box.className = "c";
    ov.appendChild(box); document.body.appendChild(ov);
    showLogin(box);
  });
})();
