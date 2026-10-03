/* Cambio de tema y persistencia en localStorage. */

var _themingTimer=null;
var currentTheme=null;

/* Todas las clases que puede dejar puesto un tema, activo o retirado, para
   que al cambiar no quede ninguna pegada (p.ej. «google» guardado de antes) */
function _allThemeClasses(){
  var fam=Object.keys(THEME_FAMILY).map(function(k){return THEME_FAMILY[k];});
  return THEME_ORDER.concat(THEMES_LEGACY,fam).filter(function(x){return x!=="dark";});
}

/* Aplica un tema. `persist` solo es true cuando el usuario lo elige con el
   botón: así quien nunca lo ha tocado sigue la preferencia del sistema. */
function applyTheme(t,persist){
  /* Funde los colores mientras dura el cambio; fuera de esa ventana la
     clase se retira para no ralentizar el resto de interacciones. */
  document.body.classList.add("theming");
  clearTimeout(_themingTimer);
  _themingTimer=setTimeout(function(){document.body.classList.remove("theming");},450);
  _allThemeClasses().forEach(function(x){document.body.classList.remove(x);});
  if(t!=="dark") document.body.classList.add(t);
  if(THEME_FAMILY[t]) document.body.classList.add(THEME_FAMILY[t]);
  currentTheme=t;
  document.documentElement.style.background=THEME_PANEL[t]||"#130e08";
  var meta=document.querySelector('meta[name="theme-color"]');
  if(meta) meta.setAttribute("content",THEME_PANEL[t]||"#130e08");
  if(persist){try{localStorage.setItem("theme",t);}catch(e){}}
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

/* Tema guardado por el usuario, solo si sigue siendo uno de los activos */
function savedTheme(){
  var t=null;
  try{t=localStorage.getItem("theme");}catch(e){}
  if(t&&THEME_ORDER.indexOf(t)>=0) return t;
  /* Un valor de un tema retirado ya no sirve: se olvida para que mande el
     sistema en lugar de quedarse con un tema que no está en la rotación */
  if(t){try{localStorage.removeItem("theme");}catch(e){}}
  return null;
}

/* Tema de entrada cuando no hay elección guardada */
function defaultTheme(){
  if(DEFAULT_THEME!=="auto") return DEFAULT_THEME;
  var dark=window.matchMedia&&window.matchMedia("(prefers-color-scheme: dark)").matches;
  return dark?AUTO_THEMES.dark:AUTO_THEMES.light;
}

/* Mientras no haya elección guardada, seguir los cambios del sistema en vivo */
if(window.matchMedia){
  var _mq=window.matchMedia("(prefers-color-scheme: dark)");
  var _onScheme=function(){if(!savedTheme()&&DEFAULT_THEME==="auto") applyTheme(defaultTheme(),false);};
  if(_mq.addEventListener) _mq.addEventListener("change",_onScheme);
  else if(_mq.addListener) _mq.addListener(_onScheme);
}


function toggleTheme(){
  var cur=currentTheme||savedTheme()||defaultTheme();
  var idx=THEME_ORDER.indexOf(cur);
  applyTheme(THEME_ORDER[(idx+1)%THEME_ORDER.length],true);
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
