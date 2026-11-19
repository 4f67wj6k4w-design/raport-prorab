/* Приложение на экране телефона и компьютера (ГК КРАШМАШ).
   1) Регистрирует фоновый модуль sw.js: страница открывается без интернета.
   2) Запоминает личную ссылку (#p=… или #m=…) и восстанавливает её,
      когда приложение запущено со значка без ссылки. */
(function () {
  "use strict";
  var page = (location.pathname.split("/").pop() || "index.html").toLowerCase();
  var KEY = "raport2-lasthash-" + page;
  try {
    var h = location.hash || "";
    if (/[#&](p|m)=/.test(h) || h === "#menu") {
      localStorage.setItem(KEY, h);
    } else if ((!h || h === "#") && !/admin/.test(location.search)) {
      var saved = localStorage.getItem(KEY);
      if (saved) history.replaceState(null, "", location.pathname + location.search + saved);
    }
  } catch (e) {}
  if ("serviceWorker" in navigator && location.protocol === "https:") {
    window.addEventListener("load", function () {
      navigator.serviceWorker.register("sw.js").catch(function () {});
    });
  }
  // постоянное хранилище: браузер не будет сам удалять очередь рапортов
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {}
})();
