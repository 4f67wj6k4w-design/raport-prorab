/* Фоновый модуль приложения ГК КРАШМАШ: хранит страницы, чтобы они открывались без интернета.
   Данные (запросы к серверу Google) НЕ кэшируются — они всегда идут напрямую. */
var VER = "krm-v2-3";
var SHELL = ["./index.html", "./dsk.html", "./director.html", "./exec.html", "./auth.js", "./pwa.js", "./config.js", "./dostup.html",
  "./tema.css", "./tema.js", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"];
self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(VER).then(function (c) {
    return Promise.all(SHELL.map(function (u) { return c.add(new Request(u, { cache: "reload" })).catch(function () {}); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== VER; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener("fetch", function (e) {
  var req = e.request, url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== self.location.origin) return; // сервер Google, шрифты — мимо
  // сначала сеть (свежая версия), без сети — сохранённая копия
  // при слабом сигнале не ждём сеть дольше 4 секунд: открываем сохранённую копию
  var net = Promise.race([fetch(req), new Promise(function (_, rej) { setTimeout(function () { rej(new Error("timeout")); }, 4000); })]);
  e.respondWith(net.then(function (res) {
    if (res && res.ok) { var cp = res.clone(); caches.open(VER).then(function (c) { c.put(req, cp); }); }
    return res;
  }).catch(function () {
    return caches.match(req, { ignoreSearch: true }).then(function (r) {
      return r || caches.match("./index.html");
    });
  }));
});
