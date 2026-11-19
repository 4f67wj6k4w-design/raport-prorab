/* Вход по логину и паролю — ГК КРАШМАШ, версия 2.
   Подключается в <head> после config.js: <script src="config.js"></script><script src="auth.js"></script>
   - адрес сервера берётся из config.js (window.KRM_API), ключей в адресной строке нет;
   - токен хранится только на этом устройстве (localStorage) и отправляется в теле POST-запроса;
   - страница получает доступ через window.RaportAuth.db() → { u: адрес, k: "tk:ТОКЕН" }. */
(function () {
  "use strict";
  var LS = "raport2-auth-v2";
  var U = window.KRM_API || "";
  function ld() { try { return JSON.parse(localStorage.getItem(LS) || "null"); } catch (e) { return null; } }
  function sv(o) { try { localStorage.setItem(LS, JSON.stringify(o)); } catch (e) {} }
  var st = ld() || {};
  if (st.u !== U) st = {};                                  // другой сервер — войти заново
  // старые ссылки v1 с ключом в адресе: ключ из адресной строки убрать
  try { if (/[#&]k=/.test(location.hash)) history.replaceState(null, "", location.pathname + location.search); } catch (e) {}
  var DEV = (function () { var ua = navigator.userAgent || ""; var m = /(iPhone|iPad|Android[^;)]*|Windows NT [\d.]+|Mac OS X [\d_]+)/.exec(ua); return (m ? m[1] : "устройство") + (/Edg\//.test(ua) ? " Edge" : /Chrome\//.test(ua) ? " Chrome" : /Safari\//.test(ua) ? " Safari" : /Firefox\//.test(ua) ? " Firefox" : ""); })();
  var origFetch = window.fetch ? window.fetch.bind(window) : null;
  function logout(reload) {
    st = { u: U }; sv(st);
    if (reload !== false) location.reload();
  }
  var haveToken = !!(st.t && U);
  window.RaportAuth = {
    u: function () { return U; },
    token: function () { return haveToken ? st.t : ""; },
    db: function () { return haveToken ? { u: U, k: "tk:" + st.t } : null; },
    dev: DEV,
    name: function () { return st.name || ""; },
    role: function () { return st.role || ""; },
    legacy: false,
    logout: logout
  };

  // токен просрочен или отозван — вернуть на экран входа
  if (origFetch) {
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
  if (haveToken) {
    document.addEventListener("DOMContentLoaded", function () {
      var b = document.createElement("button");
      b.type = "button"; b.id = "auth-out";
      b.textContent = (st.name ? st.name + " · " : "") + "Выйти";
      b.setAttribute("style", "position:fixed;right:8px;bottom:6px;z-index:40;font:12px/1 Arial,sans-serif;padding:6px 10px;border-radius:14px;border:1px solid rgba(128,128,128,.4);background:rgba(255,255,255,.88);color:#333;cursor:pointer");
      b.onclick = function () { logout(); };
      document.body.appendChild(b);
    });
    return;
  }

  /* ---------- экран входа ---------- */
  var root = document.documentElement;
  var lock = document.createElement("style");
  lock.textContent = "html.auth-lock body>*:not(#auth-ov){display:none!important}#auth-ov{position:fixed;inset:0;z-index:100;display:flex;align-items:center;justify-content:center;padding:16px;background:#2E671F;font-family:Arial,sans-serif}#auth-ov .c{background:#fff;color:#1d1d1b;border-radius:14px;max-width:380px;width:100%;padding:22px 20px}#auth-ov h2{margin:0 0 4px;font-size:22px}#auth-ov p{margin:0 0 14px;color:#555;font-size:14px;line-height:1.4}#auth-ov label{display:block;font-size:13px;color:#555;margin:10px 0 4px}#auth-ov input{width:100%;box-sizing:border-box;font-size:17px;padding:11px;border:1px solid #bbb;border-radius:8px;background:#fff;color:#111}#auth-ov button{width:100%;margin-top:16px;min-height:48px;font-size:17px;font-weight:700;border:0;border-radius:10px;background:#2E671F;color:#fff;cursor:pointer}#auth-ov button:disabled{opacity:.6}#auth-ov .e{color:#b3362c;font-size:14px;margin-top:10px;min-height:18px}";
  document.head.appendChild(lock);
  root.classList.add("auth-lock");

  function post(body) {
    body.dev = DEV;
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
    setTimeout(function () { location.reload(); }, 400);
  }
  function showLogin(box) {
    box.innerHTML = "";
    box.appendChild(mk("h2", "Вход"));
    box.appendChild(mk("p", "ГК КРАШМАШ. Введите логин и пароль, которые выдал администратор."));
    if (!U) { box.appendChild(mk("p", "Сервер не настроен (config.js). Сообщите администратору.")); return; }
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
        er.textContent = j.error === "locked" ? "Слишком много попыток. Вход заблокирован на 15 минут." : j.error === "bad" ? "Неверный логин или пароль" : j.error === "rate" ? "Слишком много попыток со стороны. Подождите 10 минут и повторите." : "Ошибка: " + (j.error || "нет связи");
      }).catch(function () { go.disabled = false; er.textContent = "Нет связи. Проверьте интернет и повторите."; });
    }
    i1.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); i2.focus(); } });
    setTimeout(function () { try { i1.focus(); } catch (e) {} }, 50);
  }
  function showChange(box, j, lg) {
    box.innerHTML = "";
    box.appendChild(mk("h2", "Задайте свой пароль"));
    box.appendChild(mk("p", "Вы вошли с временным паролем. Придумайте свой: не короче 12 символов, буквы и цифры (без «password», «qwerty», названия компании и вашего логина)."));
    var l1 = mk("label", "Новый пароль"), i1 = mk("input", null, { type: "password", autocomplete: "new-password" });
    var l2 = mk("label", "Повторите пароль"), i2 = mk("input", null, { type: "password", autocomplete: "new-password" });
    var un = mk("input", null, { type: "text", name: "username", autocomplete: "username", value: lg || "" }); un.style.display = "none";
    var go = mk("button", "Сохранить и войти", { type: "submit" }), er = mk("div", ""); er.className = "e";
    var fm = mk("form", null, { autocomplete: "on" }); box.appendChild(fm);
    [un, l1, i1, l2, i2, go, er].forEach(function (x) { fm.appendChild(x); });
    fm.addEventListener("submit", function (e) { e.preventDefault(); submit(); });
    function submit() {
      var p = i1.value;
      if (p.length < 12) { er.textContent = "Пароль короче 12 символов"; return; }
      if (!/[0-9]/.test(p) || !/[A-Za-zА-Яа-яЁё]/.test(p)) { er.textContent = "Слишком простой пароль: нужны буквы и цифры, без «password», «qwerty», названия компании и логина"; return; }
      if (p !== i2.value) { er.textContent = "Пароли не совпадают"; return; }
      go.disabled = true; er.textContent = "";
      post({ a: "chpass", t: j.token, p: p }).then(function (r) {
        go.disabled = false;
        if (r.ok) { finish(r); return; }
        er.textContent = r.error === "short" ? "Пароль короче 12 символов" : r.error === "weak" ? "Слишком простой пароль: нужны буквы и цифры, без «password», «qwerty», названия компании и логина" : "Ошибка: " + (r.error || "нет связи");
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
