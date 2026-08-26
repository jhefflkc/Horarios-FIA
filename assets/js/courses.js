/* Modelo de cursos: agrupa las filas del Excel en secciones, reparte los
   docentes por tipo de sesión, detecta cruces y gestiona la selección. */

function _tOrder(t){return _TORD[t]===undefined?9:_TORD[t];}

/* "T"+"P" \u2192 "Teor\u00eda \u00b7 Pr\u00e1ctica" */
function _docRole(tipos){
  return (tipos||[]).map(function(t){return TN[t]||t;}).join(" \u00b7 ");
}

/* Etiqueta de rol: solo aporta si la secci\u00f3n tiene m\u00e1s de un docente */
function _docRoleHTML(sec,d){
  if(!sec.docs||sec.docs.length<2||!d.tipos.length) return "";
  return "<span class=\"doc-role\">"+_docRole(d.tipos)+"</span>";
}

/* Agrupa las sesiones de una secci\u00f3n por docente y de qu\u00e9 tipos se encarga cada uno.
   Una secci\u00f3n puede tener varios (p.ej. uno dicta teor\u00eda y otro el laboratorio). */
function _rebuildDocs(s){
  const dm={},order=[];
  s.ss.forEach(function(x){
    const n=(x.dc||"").trim();
    if(!n) return;
    if(!dm[n]){dm[n]={name:n,tipos:[]};order.push(n);}
    if(dm[n].tipos.indexOf(x.t)<0) dm[n].tipos.push(x.t);
  });
  s.docs=order.map(function(n){
    dm[n].tipos.sort(function(a,b){return _tOrder(a)-_tOrder(b);});
    return dm[n];
  }).sort(function(a,b){return _tOrder(a.tipos[0])-_tOrder(b.tipos[0]);});
  /* docente "principal" (calificaciones, exportaciones): el que dicta teor\u00eda */
  const th=s.docs.filter(function(d){return d.tipos.indexOf("T")>=0;})[0];
  s.docente=(th||s.docs[0]||{}).name||"";
  return s;
}

function fmtEsps(c){
  var list=c.esps&&c.esps.length?c.esps:[c.esp];
  return list.map(function(e){return ESP_SH[e]||e;}).join(" \u00b7 ");
}

function boot(rows){
  loadRatings();
  const map={};
  rows.forEach(function(r){
    const k=r.cod+"|"+r.secc;
    if(!map[k]) map[k]={cod:r.cod,curso:r.curso,secc:r.secc,esp:r.esp,docente:r.docente,ss:[]};
    map[k].ss.push({t:r.tipo,d:r.dia,h0:r.hIni,h1:r.hFin,rm:r.salon,dc:(r.docente||"").trim()});
  });
  Object.values(map).forEach(_rebuildDocs);
  const cm={};
  Object.values(map).forEach(function(s){
    if(!cm[s.cod]) cm[s.cod]={cod:s.cod,curso:s.curso,esp:s.esp,esps:[],secs:[]};
    if(!cm[s.cod].esps.includes(s.esp)) cm[s.cod].esps.push(s.esp);
    cm[s.cod].secs.push(s);
  });
  courses=Object.values(cm).sort(function(a,b){return a.curso.localeCompare(b.curso);});
  sel={};pal={};palIdx=0;
  drawList();drawSched();drawTags();
}



function toggle(cod,fromCbox){
  if(sel[cod]){delete sel[cod];delete pal[cod];sync();return;}
  const c=courses.find(function(x){return x.cod===cod;});
  if(!c) return;
  if(c.secs.length===1) pick(c.secs[0],cod,fromCbox);
  else openModal(c);
}


function pick(sec,cod,stayOnCourses){
  _pendingCod=null;_pendingSec=null;
  var cf=checkSectionConflict(sec);
  if(cf.hasPractice){
    toast("No se puede agregar "+cod+": cruza práctica con "+cf.practiceWith.join(", "),"er");
    document.getElementById("modal").style.display="none";
    return;
  }
  if(!pal[cod]){pal[cod]=PAL[palIdx%PAL.length];palIdx++;}
  sel[cod]=sec;
  document.getElementById("modal").style.display="none";
  if(window.innerWidth<=720&&!stayOnCourses) showMobTab('schedule');
  if(cf.hasTheory) toast("Cruce de teoría con "+cf.theoryWith.join(", ")+" — aparecerán juntos","warn");
  sync();
}


function sync(){drawList();drawSched();drawTags();updateScorePill();}


/* Returns {hasPractice, hasTheory, practiceWith, theoryWith} for a given section */
function checkSectionConflict(sec){
  var practiceWith=[],theoryWith=[];
  sec.ss.forEach(function(ss){
    Object.keys(sel).forEach(function(cod){
      sel[cod].ss.forEach(function(x){
        if(x.d===ss.d&&!(ss.h1<=x.h0||ss.h0>=x.h1)){
          if(ss.t==='P'&&x.t==='P'){if(!practiceWith.includes(cod)) practiceWith.push(cod);}
          else{if(!theoryWith.includes(cod)) theoryWith.push(cod);}
        }
      });
    });
  });
  return{hasPractice:practiceWith.length>0,hasTheory:theoryWith.length>0,practiceWith:practiceWith,theoryWith:theoryWith};
}

/* True if ALL sections of course c are blocked by a practice conflict */
function allHardConflict(c){
  return c.secs.every(function(s){return checkSectionConflict(s).hasPractice;});
}

/* True if at least one section has a theory overlap (no practice block) */
function anySoftConflict(c){
  if(allHardConflict(c)) return false;
  return c.secs.some(function(s){var cf=checkSectionConflict(s);return cf.hasTheory&&!cf.hasPractice;});
}


function clearAll(){sel={};pal={};palIdx=0;sync();}

function toggleNames(){
  showNames=!showNames;
  const btn=document.getElementById("btn-names");
  const lbl=document.getElementById("btn-names-lbl");
  if(btn) btn.classList.toggle("active",showNames);
  if(lbl) lbl.textContent=showNames?"Código":"Nombre";
  drawSched();
}


function getCiclo(cod){var c=String(cod)[2];return(c&&!isNaN(c))?parseInt(c):11;}

function parseHorario(h){var p=String(h).trim().split(" ");if(p.length!==2) return null;var r=p[1].split("-");if(r.length!==2) return null;var h0=parseInt(r[0]),h1=parseInt(r[1]);if(isNaN(h0)||isNaN(h1)) return null;return{dia:p[0].toUpperCase(),hIni:h0,hFin:h1};}
