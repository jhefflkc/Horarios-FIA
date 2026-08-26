/* Cambio de tema y persistencia en localStorage. */

function applyTheme(t){
  THEME_ORDER.forEach(function(x){if(x!=="dark") document.body.classList.remove(x);});
  if(t!=="dark") document.body.classList.add(t);
  document.documentElement.style.background=THEME_PANEL[t]||"#130e08";
  localStorage.setItem("theme",t);
  var icon=THEME_ICON[t]||"dark";
  ["dark","light","stitch","google"].forEach(function(k){
    var el=document.getElementById("theme-icon-"+k);
    if(el) el.style.display=(k===icon)?"":"none";
  });
  document.getElementById("theme-label").textContent=THEME_LABELS[t]||t;
  /* Las etiquetas llevan el color en línea, así que hay que repintarlas
     cuando cambia el tema (cada tema tiene su propia paleta). */
  if(typeof drawTags==="function"&&typeof sel!=="undefined") drawTags();
}


function toggleTheme(){
  var cur=localStorage.getItem("theme")||"dark";
  var idx=THEME_ORDER.indexOf(cur);
  applyTheme(THEME_ORDER[(idx+1)%THEME_ORDER.length]);
}


/* Onda al pulsar, como en los componentes nativos de Material.
   Solo actúa con el tema Google; los demás no la usan. */
var _G_RIPPLE_SEL=".btn,.hbtn,.chip,.c-row,.sec,.rate-btn,.modal-x,.ann-x,.ann-btn-ok,.ann-btn-skip";

function _gRipple(e){
  if(!document.body.classList.contains("google")) return;
  if(e.button) return;                      /* solo el botón principal */
  var t=e.target.closest(_G_RIPPLE_SEL);
  if(!t||t.classList.contains("conf")) return;
  var r=t.getBoundingClientRect();
  if(!r.width) return;
  var d=Math.max(r.width,r.height)*2;
  var s=document.createElement("span");
  s.className="g-ripple";
  s.style.width=s.style.height=d+"px";
  s.style.left=(e.clientX-r.left-d/2)+"px";
  s.style.top=(e.clientY-r.top-d/2)+"px";
  t.appendChild(s);
  setTimeout(function(){s.remove();},560);
}

document.addEventListener("pointerdown",_gRipple);
