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
}

/* Init mobile panels */
if(window.innerWidth<=720){
  document.querySelector('.sidebar').classList.add('mob-active');
}


function toast(msg,type){
  const t=document.getElementById("toast");
  document.getElementById("tmsg").textContent=msg;
  const tabs=document.querySelector(".mob-tabs");
  if(tabs&&tabs.offsetHeight>0){
    t.style.bottom=(tabs.offsetHeight+12)+"px";
  } else {
    t.style.bottom="";
  }
  t.className="toast on "+(type||"ok");
  clearTimeout(t._t);t._t=setTimeout(function(){t.classList.remove("on");},3500);
}


document.addEventListener("mousemove",function(e){
  document.documentElement.style.setProperty("--dot-x",e.clientX+"px");
  document.documentElement.style.setProperty("--dot-y",e.clientY+"px");
});


window.addEventListener("DOMContentLoaded",function(){
  var saved=localStorage.getItem("theme");
  if(!saved){
    saved=window.matchMedia("(prefers-color-scheme: light)").matches?"light":"dark";
  }
  applyTheme(saved);
  initFacultySelector();
  switchFaculty(Object.keys(ALL_DATA)[0]);
  fetchAnnounce();
  document.addEventListener("click",function(){
    var pop=document.getElementById("score-pop");
    if(pop) pop.style.display="none";
  });
});
