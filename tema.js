/* ГК КРАШМАШ — переключатель светлой/тёмной темы (общий для страниц). Выбор запоминается на устройстве. */
(function(){
  var K="raport2-theme";
  function get(){try{return localStorage.getItem(K);}catch(e){return null;}}
  function set(v){try{localStorage.setItem(K,v);}catch(e){}}
  var root=document.documentElement,t=get();
  if(t==="dark"||t==="light")root.setAttribute("data-theme",t);
  function isDark(){var a=root.getAttribute("data-theme");return a?a==="dark":!!(window.matchMedia&&matchMedia("(prefers-color-scheme: dark)").matches);}
  function label(b){b.textContent=isDark()?"\u2600":"\u263E";}
  function init(){
    var row=document.querySelector("header.top .brandrow");if(!row||document.getElementById("thm"))return;
    var b=document.createElement("button");b.type="button";b.id="thm";b.title="Тёмный / светлый вид";b.setAttribute("aria-label","Переключить тёмный или светлый вид");label(b);
    var img=row.querySelector("img");row.insertBefore(b,img||null);
    b.onclick=function(){var n=isDark()?"light":"dark";root.setAttribute("data-theme",n);set(n);label(b);};
  }
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init);else init();
})();
