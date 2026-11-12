/* Pintado de la interfaz: listado de cursos, etiquetas y tabla de horario. */

/* Cada fila se genera por separado; si la lista es la misma (mismo filtro y
   mismos cursos) solo se sustituyen las filas cuyo HTML cambió. Rehacer las
   164 filas en cada clic costaba ~55 ms de estilo+layout (CPU x4) y movía la
   lista: el ancla de desplazamiento del navegador se perdía. */
var _listKey=null,_listRowHTML={};
function _courseRowHTML(c){
    const isSel=!!sel[c.cod];
    let isConf=false,isSoft=false;
    if(!isSel){   /* una sola pasada: antes cada sección se comprobaba hasta 3 veces */
      let allHard=true,anySoft=false;
      c.secs.forEach(function(s){const cf=checkSectionConflict(s);if(!cf.hasPractice){allHard=false;if(cf.hasTheory)anySoft=true;}});
      isConf=allHard;isSoft=!allHard&&anySoft;
    }
    const ss=sel[c.cod];
    const meta=ss?"Secc. "+ss.secc+" \u00b7 "+ss.ss.length+" ses.":c.secs.length+(c.secs.length>1?" secc.":"\u00a0secc.");
    /* «multi»: al pulsarla se abre la ventana de secciones (el tema muestra un chevrón) */
    /* Con teclado: cada fila es una casilla que se alcanza con Tab y se
       marca con Enter o Espacio (main.js); las bloqueadas quedan fuera */
    let h="<div class=\"c-row cy"+c.esp+(c.secs.length>1?" multi":"")+(isSel?" sel":"")+(isConf?" conf":"")+"\" data-cod=\""+c.cod+"\""+
      " role=\"checkbox\" aria-checked=\""+(isSel?"true":"false")+"\""+(isConf?" aria-disabled=\"true\"":" tabindex=\"0\"")+">";
    h+="<div class=\"cbox\"></div>";
    h+="<div class=\"c-info\">";
    h+="<div class=\"c-name\">"+c.curso+"</div>";
    h+="<div class=\"c-meta\"><span class=\"c-cod\">"+c.cod+"</span><span class=\"c-secs\">"+meta+"</span>";
    if(isConf) h+="<span class=\"c-conf-tag\">práctica bloqueada</span>";
    else if(isSoft) h+="<span class=\"c-conf-tag c-conf-tag-soft\">cruce teor\u00eda</span>";
    h+="</div></div></div>";
    return h;
}
document.getElementById("clist").addEventListener("click",function(e){
  const r=e.target.closest(".c-row");
  if(r) toggle(r.dataset.cod,e.target.closest(".cbox")!==null);
});

function drawList(){
  const q=document.getElementById("q").value.toLowerCase().trim();
  const list=courses.filter(function(c){
    return !q||c.cod.toLowerCase().includes(q)||c.curso.toLowerCase().includes(q);
  });
  const el=document.getElementById("clist");
  if(!list.length){
    _listKey=null;
    var emptyMsg="Sin resultados";
    el.innerHTML="<div style=\"padding:2rem;text-align:center;color:var(--tx4);font-size:0.8rem\">"+emptyMsg+"</div>";
    return;
  }
  /* Agrupado por ciclo, activable desde GROUP_BY_CYCLE (config.js). Con el
     interruptor apagado va todo en una sola lista, sin cabeceras.
     Un ciclo fuera de 1–10 (o ausente) significa que el Excel no lo trae y
     no se pudo deducir del código: van juntos al final. */
  if(!GROUP_BY_CYCLE){
    const key=list.map(function(c){return c.cod;}).join(",");
    const html=list.map(_courseRowHTML);
    if(key===_listKey&&el.children.length===list.length){
      list.forEach(function(c,i){
        if(_listRowHTML[c.cod]===html[i]) return;
        const t=document.createElement("template");t.innerHTML=html[i];
        const row=t.content.firstChild,old=el.children[i];
        const hadFocus=document.activeElement===old;
        old.replaceWith(row);
        if(hadFocus&&row.focus) row.focus({preventScroll:true});   /* el foco sigue en la fila repintada */
        _listRowHTML[c.cod]=html[i];
      });
      return;
    }
    _listKey=key;_listRowHTML={};
    list.forEach(function(c,i){_listRowHTML[c.cod]=html[i];});
    el.innerHTML=html.join("");
    return;
  }
  _listKey=null;
  const grupos={};
  list.forEach(function(c){
    if(!GROUP_BY_CYCLE){(grupos["todos"]=grupos["todos"]||[]).push(c);return;}
    const n=parseInt(c.ciclo);
    const k=(n>=1&&n<=10)?n:"otros";
    (grupos[k]=grupos[k]||[]).push(c);
  });
  const orden=Object.keys(grupos).filter(function(k){return k!=="otros"&&k!=="todos";})
    .map(Number).sort(function(a,b){return a-b;});
  if(grupos["otros"]) orden.push("otros");
  if(grupos["todos"]) orden.push("todos");

  let h="";
  orden.forEach(function(k){
    if(k!=="todos"){
      h+="<div class=\"cy-head\">"+(k==="otros"?"Otros cursos":"Ciclo "+k)+
         "<span class=\"cy-head-n\">"+grupos[k].length+"</span></div>";
    }
    grupos[k].forEach(function(c){h+=_courseRowHTML(c);
    });
  });
  el.innerHTML=h;
}


function blkInfo(e){
  e.stopPropagation();
  const cod=e.currentTarget.dataset.cod;
  openInfoModal(cod);
}


var _tagsPrev={};
function drawTags(){
  const keys=Object.keys(sel);
  const prevT=_tagsPrev;_tagsPrev={};
  keys.forEach(function(cod){_tagsPrev[cod]=sel[cod].secc;});
  const tz=document.getElementById("tags");
  const bc=document.getElementById("bclear");
  const st=document.getElementById("stats");
  if(!keys.length){
    tz.innerHTML="<span class=\"ph\">&#8592; Selecciona cursos para armar tu horario</span>";
    bc.style.display="none";st.style.display="none";
    document.getElementById("btn-names").style.display="none";return;
  }
  let h="";
  keys.forEach(function(cod){
    const s=sel[cod];const p=pal[cod]||"p0";const hex=palHex(p);
    h+="<div class=\"tag "+p+(prevT[cod]!==s.secc?" fx-in":"")+"\" style=\"background:"+hex+"1e;border-color:"+hex+"50;color:"+hex+"\" data-cod=\""+cod+"\">";
    h+=cod+" <span style=\"opacity:0.6\">"+s.secc+"</span><button type=\"button\" class=\"tag-x\" aria-label=\"Quitar "+cod+"\">\u00d7</button></div>";
  });
  tz.innerHTML=h;
  tz.querySelectorAll(".tag").forEach(function(t){
    t.querySelector(".tag-x").addEventListener("click",function(e){e.stopPropagation();toggle(t.dataset.cod);});
  });
  bc.style.display="";st.style.display="";   /* "" deja mandar al CSS (flex, o rejilla en Vuelo móvil) */
  document.getElementById("btn-names").style.display="";
  let hrs=0;
  Object.values(sel).forEach(function(s){s.ss.forEach(function(x){hrs+=x.h1-x.h0;});});
  document.getElementById("sn").textContent=keys.length;
  document.getElementById("sh").textContent=hrs;
  const isFIA=facultyName==="Fac. Ing. Ambiental";
  const scEl=document.getElementById("sc").parentElement;
  scEl.style.display=isFIA?"":"none";
  if(isFIA){
    const cy=[...new Set(Object.keys(sel).map(function(cod){return sel[cod].esp;}).filter(Boolean))].sort(function(a,b){return ESP_ORDER.indexOf(a)-ESP_ORDER.indexOf(b);});
    document.getElementById("sc").textContent=cy.map(function(n){return ESP_SH[n]||n;}).join(", ")||"\u2014";
  }
}


/* La carpeta del estado vacío del tema Vuelo: dos billetes con el periodo
   y la facultad, y el frente con el rótulo. Es decorativa (aria-hidden) y
   los temas anteriores la ocultan (styles.css). */
function emptyFolderHTML(){
  var esc=function(t){return String(t).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c];});};
  var plane="<svg viewBox=\"0 0 24 24\" width=\"14\" height=\"14\" fill=\"currentColor\"><path d=\"M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5z\"/></svg>";
  return "<div class=\"v-folder\" aria-hidden=\"true\">"+
    "<div class=\"v-ticket\"><span class=\"v-t-date\">Semestre "+esc(getPeriod())+"</span><b class=\"v-badge\">En armado</b>"+
      "<span class=\"v-t-route\">"+plane+"<i></i><em></em><span class=\"v-t-code\">LIM</span></span></div>"+
    "<div class=\"v-ticket v-ticket-2\"><span class=\"v-t-date\">Sin cursos a\u00fan</span></div>"+
    "<div class=\"v-folder-front\">Horario</div></div>";
}

/* Bloques que entran con animación: solo los de cursos nuevos o cuya
   sección/sesiones cambiaron. Antes todos volvían a «aparecer» (opacidad 0)
   en cada cambio y el horario parpadeaba entero. */
var _schedPrev={};
function drawSched(){
  const keys=Object.keys(sel);
  const dp=document.getElementById("dp");
  const prevS=_schedPrev;_schedPrev={};
  keys.forEach(function(cod){const s=sel[cod];_schedPrev[cod]=s.secc+"|"+s.ss.map(function(x){return x.t+x.d+x.h0+"-"+x.h1+x.rm;}).join(",");});
  const blkIn=function(cod){return prevS[cod]!==_schedPrev[cod]?" fx-in":"";};
  if(!keys.length){
    /* sin cursos no puede haber cruces: el aviso se retira (antes seguía a
       la vista tras «Limpiar todo» o al cambiar de facultad) */
    document.getElementById("conf").classList.remove("on");
    dp.innerHTML="<div class=\"empty\"><div class=\"empty-box\"><svg width=\"36\" height=\"36\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.3\" opacity=\"0.3\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"18\" rx=\"2\"/><line x1=\"16\" x2=\"16\" y1=\"2\" y2=\"6\"/><line x1=\"8\" x2=\"8\" y1=\"2\" y2=\"6\"/><line x1=\"3\" x2=\"21\" y1=\"10\" y2=\"10\"/></svg>"+emptyFolderHTML()+"</div><h3>Tu horario aparecer\u00e1 aqu\u00ed</h3><p>Selecciona tus cursos del panel izquierdo para comenzar.</p></div>";
    return;
  }
  const cells={};
  keys.forEach(function(cod){
    const s=sel[cod];const p=pal[cod]||"p0";
    s.ss.forEach(function(x){
      for(let h=x.h0;h<x.h1;h++){
        const k=x.d+"-"+h;
        if(!cells[k]) cells[k]=[];
        if(!cells[k].find(function(e){return e.cod===cod;}))
          cells[k].push({cod:cod,sec:s.secc,tipo:x.t,rm:x.rm,p:p});
      }
    });
  });
  const conflicts=[];
  Object.keys(cells).forEach(function(k){
    if(cells[k].length>1){
      const parts=k.split("-");const d=parts[0];const hr=parts[1];
      conflicts.push(cells[k].map(function(e){return e.cod;}).join(" y ")+" \u2014 "+(DN[d]||d)+" "+hr+":00");
    }
  });
  const ca=document.getElementById("conf");
  if(conflicts.length){
    const nuevo=!ca.classList.contains("on");
    ca.classList.add("on");
    document.getElementById("cmsg").textContent="Cruce de teor\u00eda: "+conflicts[0]+(conflicts.length>1?" (+"+(conflicts.length-1)+" m\u00e1s)":"");
    if(nuevo){
      /* la barra abre su hueco poco a poco (fx-bar) en vez de empujar el
         horario de golpe: el margen parte de su alto en negativo */
      ca.style.setProperty("--conf-pull",-(ca.offsetHeight+parseFloat(getComputedStyle(ca).marginTop))+"px");
      ca.classList.add("fx-in");
    }
  } else ca.classList.remove("on","fx-in");

  const allH=[];
  Object.values(sel).forEach(function(s){s.ss.forEach(function(x){allH.push(x.h0);allH.push(x.h1);});});
  const mn=Math.min.apply(null,allH),mx=Math.max.apply(null,allH);
  /* Se incluye la hora de fin como última fila: sin ella el bloque que
     termina a las 21:00 se quedaba sin marca de cierre en la rejilla. */
  const hrs=[];for(let i=mn;i<=mx;i++) hrs.push(i);

  const todayIdx=new Date().getDay();
  const todayKey=["DO","LU","MA","MI","JU","VI","SA"][todayIdx];

  let selCount=keys.length;
  let h="<div class=\"sched-wrap\" id=\"printable\">";
  h+="<table class=\"sch\"><thead><tr><th></th>";
  DAYS.forEach(function(d){
    h+="<th><div class=\"dhead"+(d===todayKey?" now":"")+"\">"+(DN[d]||d)+"</div></th>";
  });
  h+="</tr></thead><tbody>";
  hrs.forEach(function(hr){
    h+="<tr><td>"+hr+":00</td>";
    DAYS.forEach(function(d){
      const arr=cells[d+"-"+hr];
      if(!arr||!arr.length){h+="<td><div class=\"cell-mt\"></div></td>";return;}
      if(arr.length===1){
        const c=arr[0];
        h+="<td><div class=\"blk "+c.p+blkIn(c.cod)+"\" data-cod=\""+c.cod+"\" onclick=\"blkInfo(event)\" role=\"button\" tabindex=\"0\" aria-label=\""+c.cod+" secci\u00f3n "+c.sec+": ver informaci\u00f3n\">";
        if(showNames){
          const cn=(sel[c.cod]||{}).curso||c.cod;
          h+="<div class=\"blk-name\">"+cn+"</div>";
          h+="<div class=\"blk-cod blk-cod-sm\">"+c.cod+" <span class=\"blk-secc\">"+c.sec+"</span></div>";
        } else {
          h+="<div class=\"blk-cod\">"+c.cod+" <span class=\"blk-secc\">"+c.sec+"</span></div>";
        }
        h+="<div class=\"blk-type\">"+(TN[c.tipo]||c.tipo)+"</div>";
        if(c.rm) h+="<div class=\"blk-room\">"+c.rm+"</div>";
        h+="</div></td>";
      } else {
        h+="<td><div class=\"cell-blocks\">";
        arr.forEach(function(c){
          h+="<div class=\"blk "+c.p+blkIn(c.cod)+"\" data-cod=\""+c.cod+"\" onclick=\"blkInfo(event)\" role=\"button\" tabindex=\"0\" aria-label=\""+c.cod+" secci\u00f3n "+c.sec+": ver informaci\u00f3n\">";
          if(showNames){
            const cn=(sel[c.cod]||{}).curso||c.cod;
            h+="<div class=\"blk-name\">"+cn+"</div>";
            h+="<div class=\"blk-cod blk-cod-sm\">"+c.cod+" <span class=\"blk-secc\">"+c.sec+"</span></div>";
          } else {
            h+="<div class=\"blk-cod\">"+c.cod+" <span class=\"blk-secc\">"+c.sec+"</span></div>";
          }
          h+="<div class=\"blk-type\">"+(TN[c.tipo]||c.tipo)+"</div>";
          h+="</div>";
        });
        h+="</div></td>";
      }
    });
    h+="</tr>";
  });
  h+="</tbody></table>";
  h+="<div class=\"sched-legend\">";
  keys.forEach(function(cod){
    const s=sel[cod];
    const p=pal[cod]||"p0";
    h+="<div class=\"leg-item\">";
    h+="<div class=\"leg-dot "+p+"\"></div>";
    h+="<span class=\"leg-cod\">"+cod+"</span>";
    h+="<span class=\"leg-name\">"+s.curso+"</span>";
    h+="</div>";
  });
  h+="</div></div>";
  dp.innerHTML=h;
}
