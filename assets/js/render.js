/* Pintado de la interfaz: listado de cursos, etiquetas y tabla de horario. */

function drawList(){
  const q=document.getElementById("q").value.toLowerCase().trim();
  const list=courses.filter(function(c){
    return !q||c.cod.toLowerCase().includes(q)||c.curso.toLowerCase().includes(q);
  });
  const el=document.getElementById("clist");
  if(!list.length){
    var emptyMsg="Sin resultados";
    el.innerHTML="<div style=\"padding:2rem;text-align:center;color:var(--tx4);font-size:0.8rem\">"+emptyMsg+"</div>";
    return;
  }
  let h="";
  list.forEach(function(c){
    const isSel=!!sel[c.cod];
    const isConf=!isSel&&allHardConflict(c);
    const isSoft=!isSel&&!isConf&&anySoftConflict(c);
    const ss=sel[c.cod];
    const meta=ss?"Secc. "+ss.secc+" \u00b7 "+ss.ss.length+" ses.":c.secs.length+(c.secs.length>1?" secc.":"\u00a0secc.");
    h+="<div class=\"c-row cy"+c.esp+(isSel?" sel":"")+(isConf?" conf":"")+"\" data-cod=\""+c.cod+"\">";
    h+="<div class=\"cbox\"></div>";
    h+="<div class=\"c-info\">";
    h+="<div class=\"c-name\">"+c.curso+"</div>";
    h+="<div class=\"c-meta\"><span class=\"c-cod\">"+c.cod+"</span><span class=\"c-secs\">"+meta+"</span>";
    if(isConf) h+="<span class=\"c-conf-tag\">práctica bloqueada</span>";
    else if(isSoft) h+="<span class=\"c-conf-tag c-conf-tag-soft\">cruce teor\u00eda</span>";
    h+="</div></div></div>";
  });
  el.innerHTML=h;
  el.querySelectorAll(".c-row").forEach(function(r){
    r.addEventListener("click",function(e){toggle(r.dataset.cod,e.target.closest(".cbox")!==null);});
  });
}


function blkInfo(e){
  e.stopPropagation();
  const cod=e.currentTarget.dataset.cod;
  openInfoModal(cod);
}


function drawTags(){
  const keys=Object.keys(sel);
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
    h+="<div class=\"tag\" style=\"background:"+hex+"1e;border-color:"+hex+"50;color:"+hex+"\" data-cod=\""+cod+"\">";
    h+=cod+" <span style=\"opacity:0.6\">"+s.secc+"</span><i class=\"tag-x\">\u00d7</i></div>";
  });
  tz.innerHTML=h;
  tz.querySelectorAll(".tag").forEach(function(t){
    t.querySelector(".tag-x").addEventListener("click",function(e){e.stopPropagation();toggle(t.dataset.cod);});
  });
  bc.style.display="";st.style.display="flex";
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


function drawSched(){
  const keys=Object.keys(sel);
  const dp=document.getElementById("dp");
  if(!keys.length){
    dp.innerHTML="<div class=\"empty\"><div class=\"empty-box\"><svg width=\"36\" height=\"36\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.3\" opacity=\"0.3\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"18\" rx=\"2\"/><line x1=\"16\" x2=\"16\" y1=\"2\" y2=\"6\"/><line x1=\"8\" x2=\"8\" y1=\"2\" y2=\"6\"/><line x1=\"3\" x2=\"21\" y1=\"10\" y2=\"10\"/></svg></div><h3>Tu horario aparecer\u00e1 aqu\u00ed</h3><p>Selecciona tus cursos del panel izquierdo para comenzar.</p></div>";
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
    ca.classList.add("on");
    document.getElementById("cmsg").textContent="Cruce de teor\u00eda: "+conflicts[0]+(conflicts.length>1?" (+"+(conflicts.length-1)+" m\u00e1s)":"");
  } else ca.classList.remove("on");

  const allH=[];
  Object.values(sel).forEach(function(s){s.ss.forEach(function(x){allH.push(x.h0);allH.push(x.h1);});});
  const mn=Math.min.apply(null,allH),mx=Math.max.apply(null,allH);
  const hrs=[];for(let i=mn;i<mx;i++) hrs.push(i);

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
        h+="<td><div class=\"blk "+c.p+"\" data-cod=\""+c.cod+"\" onclick=\"blkInfo(event)\">";
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
          h+="<div class=\"blk "+c.p+"\" data-cod=\""+c.cod+"\" onclick=\"blkInfo(event)\">";
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
