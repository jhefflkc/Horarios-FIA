/* Arranque de la app, navegación móvil y avisos emergentes. */

/* ── Mobile viewport + safe-area fix (iOS Safari) ─────────────────────────
   Two bugs on iOS first load:
   1. 100dvh may differ from window.innerHeight → body height wrong → gap below tab bar
   2. env(safe-area-inset-bottom) not ready before first layout → --sab = 0
   Fix: set body height and --sab from JS, update on resize/orientationchange.
────────────────────────────────────────────────────────────────────────── */
(function(){
  var probe=document.getElementById('sab-probe');

  function applySAB(){
    if(!probe) return;
    var h=probe.getBoundingClientRect().height||0;
    document.documentElement.style.setProperty('--sab', h+'px');
  }

  function fixBodyHeight(){
    document.body.style.height=window.innerHeight+'px';
  }

  function fixAll(){
    fixBodyHeight();
    applySAB();
  }

  /* Run after first layout (double rAF ensures browser has computed env()) */
  requestAnimationFrame(function(){ requestAnimationFrame(fixAll); });
  window.addEventListener('load', fixAll);
  window.addEventListener('resize', fixAll);
  window.addEventListener('orientationchange', function(){
    setTimeout(fixAll, 100);
    setTimeout(fixAll, 400);
  });
})();


/* ── Mobile tab navigation ─────────────────────────── */
function showMobTab(tab){
  const sidebar=document.querySelector('.sidebar');
  const main=document.querySelector('.main');
  const tCourses=document.getElementById('mob-tab-courses');
  const tSched=document.getElementById('mob-tab-schedule');
  if(tab==='courses'){
    sidebar.classList.add('mob-active');
    main.classList.remove('mob-active');
    tCourses.classList.add('on');
    tSched.classList.remove('on');
  } else {
    main.classList.add('mob-active');
    sidebar.classList.remove('mob-active');
    tSched.classList.add('on');
    tCourses.classList.remove('on');
  }
  tCourses.setAttribute('aria-selected',tab==='courses'?'true':'false');
  tSched.setAttribute('aria-selected',tab==='courses'?'false':'true');
}

/* Init mobile panels */
if(window.innerWidth<=720){
  document.querySelector('.sidebar').classList.add('mob-active');
}
/* Si la página pasa a la maqueta móvil después de cargar (redimensionar,
   girar el móvil, zoom) y ninguna columna está activa, se muestra la que
   toca: antes la pantalla se quedaba en blanco */
function ensureMobPanel(){
  if(window.innerWidth>720) return;
  if(!document.querySelector('.sidebar.mob-active,.main.mob-active'))
    showMobTab(document.getElementById('mob-tab-schedule').classList.contains('on')?'schedule':'courses');
}
window.addEventListener('resize',ensureMobPanel);
window.addEventListener('orientationchange',function(){setTimeout(ensureMobPanel,100);});


/* ── Ventanas: entrada y salida ──────────────────────
   Se muestran con display:flex; al cerrar, la clase «is-closing» reproduce
   la salida (styles.css) y después se ocultan. Con «menos movimiento» el
   cierre es inmediato. */
function reduceMotion(){
  return !!(window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches);
}
function showOverlay(el){
  clearTimeout(el._hide);
  if(el._done){el.removeEventListener("animationend",el._done);el._done=null;}
  var abierta=el.style.display==="flex"&&!el.classList.contains("is-closing");
  el.classList.remove("is-closing");
  el.style.display="flex";
  /* El foco entra en la ventana y, al cerrarla, vuelve a donde estaba. Si se
     reabre antes de devolverlo, se conserva el origen que ya tenía. */
  if(!abierta&&!el._prev) el._prev=document.activeElement;
  var d=el.querySelector("[role=dialog]");
  if(d&&!d.contains(document.activeElement)) d.focus({preventScroll:true});
}
/* Devuelve el foco al control que abrió la ventana. Se hace al acabar la
   tarea en curso: quien cierra suele repintar justo después (quitar un curso,
   guardar un horario editado) y el control de origen puede haber sido
   sustituido; entonces se busca el nuevo por el código de su curso. */
function _restoreFocus(el){
  var p=el._prev;
  if(!p) return;
  setTimeout(function(){
    if(el._prev!==p) return;
    var o=_openOverlay();
    if(o===el) return;            /* se reabrió: el origen sirve para el próximo cierre */
    el._prev=null;
    if(o||p===document.body||!p.focus) return;   /* con otra ventana abierta, el foco es suyo */
    if(!document.contains(p)){
      var cod=p.dataset&&p.dataset.cod,q=null;
      if(cod&&window.CSS&&CSS.escape){
        var s='[data-cod="'+CSS.escape(cod)+'"]';
        if(p.classList.contains("blk")) q=document.querySelector("#printable .blk"+s);
        if(!q||q.offsetParent===null) q=document.querySelector(".c-row"+s);
      }
      p=q;
    }
    if(p&&document.contains(p)) p.focus({preventScroll:true});
  },0);
}
/* La ventana se oculta cuando termina su animación de salida, no a un plazo
   fijo: si el equipo va cargado la animación empieza tarde y un temporizador
   la cortaba a medias. El plazo queda solo como red por si no llega a correr. */
function hideOverlay(el){
  if(el.style.display==="none"||el.classList.contains("is-closing")) return;
  _restoreFocus(el);
  if(reduceMotion()){el.style.display="none";return;}
  var d=el.querySelector("[role=dialog]");
  el._done=function(e){
    if(e&&(e.target!==d||!/-out$/.test(e.animationName))) return;
    clearTimeout(el._hide);
    el.removeEventListener("animationend",el._done);el._done=null;
    el.classList.remove("is-closing");el.style.display="none";
  };
  el.addEventListener("animationend",el._done);
  el.classList.add("is-closing");
  el._hide=setTimeout(el._done,600);
}
/* Las entradas de lo nuevo (bloques, etiquetas, barra de cruce) llevan la
   clase «fx-in»; se retira al terminar para que no se repitan al volver a
   mostrar su panel (pestañas del móvil) */
document.addEventListener("animationend",function(e){
  var t=e.target;
  if(t.classList&&t.classList.contains("fx-in")) t.classList.remove("fx-in");
});
/* Safari en iOS solo aplica :active (el hundirse al pulsar) si hay un
   oyente táctil; vacío y pasivo no frena el desplazamiento */
document.addEventListener("touchstart",function(){},{passive:true});
/* ── Teclado ─────────────────────────────────────────
   Esc cierra lo que esté abierto; Enter y Espacio activan lo que no es un
   botón nativo (filas, bloques, secciones, estrellas, Score, cargar
   Excel); y con una ventana abierta, Tab no sale de ella. */
function _openOverlay(){
  var a=document.getElementById("ann-overlay"),m=document.getElementById("modal");
  if(a.style.display==="flex"&&!a.classList.contains("is-closing")) return a;
  if(m.style.display==="flex"&&!m.classList.contains("is-closing")) return m;
  return null;
}
document.addEventListener("keydown",function(e){
  var o;
  if(e.key==="Escape"){
    var pop=document.getElementById("score-pop");
    if(pop&&pop.style.display==="block"){pop.style.display="none";return;}
    o=_openOverlay();
    if(o) (o.id==="modal"?modalClose():closeAnnounce(false));
    return;
  }
  var t=e.target;
  if((e.key==="Enter"||e.key===" ")&&t&&t.matches&&t.matches(".c-row,.sec,.blk,.score-pill,.star-inp,.upload-row")){
    if(t.getAttribute("aria-disabled")==="true") return;
    e.preventDefault();t.click();
    return;
  }
  if(e.key==="Tab"&&(o=_openOverlay())){
    var f=[].filter.call(o.querySelectorAll('button,[href],input,select,textarea,[tabindex]:not([tabindex="-1"])'),
      function(x){return !x.disabled&&x.offsetParent!==null;});
    if(!f.length){e.preventDefault();return;}
    var first=f[0],last=f[f.length-1],act=document.activeElement;
    if(!o.contains(act)||(!e.shiftKey&&act===last)){e.preventDefault();first.focus();}
    else if(e.shiftKey&&(act===first||f.indexOf(act)<0)){e.preventDefault();last.focus();}
  }
});
/* El popover del Score se coloca una sola vez: al cambiar el tamaño de la
   ventana quedaría descolocado o fuera de pantalla, así que se cierra */
window.addEventListener("resize",function(){
  var pop=document.getElementById("score-pop");
  if(pop) pop.style.display="none";
});


/* ms: duración (3,5 s por defecto); 0 lo deja fijo hasta el siguiente aviso */
function toast(msg,type,ms){
  const t=document.getElementById("toast");
  document.getElementById("tmsg").textContent=msg;
  const tabs=document.querySelector(".mob-tabs");
  if(tabs&&tabs.offsetHeight>0){
    /* Se mide desde el borde superior real de la barra: en Vuelo la barra
       flota separada del borde, así que su altura no basta */
    t.style.bottom=(window.innerHeight-tabs.getBoundingClientRect().top+12)+"px";
  } else {
    t.style.bottom="";
  }
  t.className="toast on "+(type||"ok");
  clearTimeout(t._t);
  if(ms!==0) t._t=setTimeout(function(){t.classList.remove("on");},ms||3500);
}


document.addEventListener("mousemove",function(e){
  /* Solo los temas Stitch usan el foco que sigue al ratón; con los demás
     no se toca :root para no recalcular estilos en cada movimiento */
  if(!/\bstitch-/.test(document.body.className)) return;
  document.documentElement.style.setProperty("--dot-x",e.clientX+"px");
  document.documentElement.style.setProperty("--dot-y",e.clientY+"px");
});


window.addEventListener("DOMContentLoaded",function(){
  /* Si el usuario eligió un tema (y sigue activo), manda ese; si no, el de
     entrada, que por defecto sigue la preferencia claro/oscuro del sistema.
     No se guarda: solo se guarda una elección hecha con el botón. */
  applyTheme(savedTheme()||defaultTheme(),false);
  initFacultySelector();
  switchFaculty(Object.keys(ALL_DATA)[0]);
  fetchAnnounce();
  document.addEventListener("click",function(){
    var pop=document.getElementById("score-pop");
    if(pop) pop.style.display="none";
  });
});
