/* Ventanas modales: elegir sección, ver un curso, editar sesiones y periodo. */

function openModal(c){
  document.getElementById("mttl").textContent=c.curso;
  document.getElementById("msub").textContent=c.cod+" \u00b7 "+facultyName+" \u00b7 Elige una secci\u00f3n";
  const body=document.getElementById("mbody");
  let h="";
  c.secs.forEach(function(sec,i){
    const cf=checkSectionConflict(sec);
    const isPractice=cf.hasPractice;const isTheory=!cf.hasPractice&&cf.hasTheory;
    h+="<div class=\"sec"+(isPractice?" conf":isTheory?" soft-conf":"")+"\" data-i=\""+i+"\">";
    h+="<div class=\"sec-top\"><span class=\"sec-name\">Secci\u00f3n "+sec.secc+"</span>";
    if(isPractice) h+="<span class=\"sec-conf\">Práctica bloqueada</span>";
    else if(isTheory) h+="<span class=\"sec-conf sec-warn\">Cruce de teoría</span>";
    h+="</div>";
    if(sec.docs&&sec.docs.length){
      h+="<div class=\"sec-docs\">";
      sec.docs.forEach(function(d){
        h+="<div class=\"sec-doc\" style=\"display:flex;align-items:center;gap:7px;flex-wrap:wrap\">";
        h+="<span>"+d.name+"</span>";
        h+=_docRoleHTML(sec,d);
        h+="<span style=\"display:inline-flex;align-items:center;gap:5px\">"+_ratingHTML(d.name)+"</span>";
        h+="</div>";
      });
      h+="</div>";
    }
    sec.ss.forEach(function(s){
      h+="<div class=\"ses\"><span class=\"ses-"+s.t+"\">"+(TN[s.t]||s.t)+"</span>";
      h+="<span>"+(DN[s.d]||s.d)+" "+s.h0+":00\u2013"+s.h1+":00</span>";
      h+="<span style=\"color:var(--tx3)\">"+s.rm+"</span></div>";
    });
    h+="</div>";
  });
  body.innerHTML=h;
  body.querySelectorAll(".sec").forEach(function(el){
    el.addEventListener("click",function(){pick(c.secs[parseInt(el.dataset.i)],c.cod,true);});
  });
  document.getElementById("modal").style.display="flex";
}


function modalClose(){
  if(_pendingCod&&_pendingSec){sel[_pendingCod]=_pendingSec;_pendingCod=null;_pendingSec=null;sync();}
  _selectedRating=0;_ratingTeacher=null;
  document.getElementById("modal").style.display="none";
}

function closeModal(e){if(e.target===document.getElementById("modal")) modalClose();}


function openPeriodModal(){
  document.getElementById("mttl").textContent="Informaci\u00f3n del horario";
  document.getElementById("msub").textContent="Facultad, a\u00f1o y semestre que aparecer\u00e1n en el horario";
  var inp="background:var(--surface);border:1.5px solid var(--line2);border-radius:9px;"+
    "padding:10px 14px;font-family:'Plus Jakarta Sans',sans-serif;font-size:0.9rem;font-weight:600;"+
    "color:var(--tx);outline:none;width:100%;transition:border-color 0.2s;box-sizing:border-box";
  var inpMono="background:var(--surface);border:1.5px solid var(--line2);border-radius:9px;"+
    "padding:11px 14px;font-family:'JetBrains Mono',monospace;font-size:1.3rem;font-weight:800;"+
    "color:var(--tx);outline:none;width:100%;text-align:center;transition:border-color 0.2s;box-sizing:border-box";
  var lbl="font-size:0.6rem;font-family:'JetBrains Mono',monospace;color:var(--tx3);"+
    "letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;display:block";
  var presetBtns=FACULTY_PRESETS.map(function(p){
    return "<button onclick=\"document.getElementById('p-faculty').value='"+p.value+"'\" "+
      "style=\"background:var(--surface);border:1.5px solid var(--line2);border-radius:7px;"+
      "padding:5px 10px;font-size:0.68rem;font-weight:700;color:var(--tx2);cursor:pointer;"+
      "font-family:'JetBrains Mono',monospace;white-space:nowrap;transition:all 0.15s\""+
      " onmouseover=\"this.style.borderColor='var(--accent2)';this.style.color='var(--accent3)'\""+
      " onmouseout=\"this.style.borderColor='var(--line2)';this.style.color='var(--tx2)'\">"+
      p.label+"</button>";
  }).join("");
  document.getElementById("mbody").innerHTML=
    "<div style=\"margin-bottom:16px\">"+
      "<span style=\""+lbl+"\">Facultad</span>"+
      "<input id=\"p-faculty\" type=\"text\" value=\""+facultyName+"\" style=\""+inp+"\" placeholder=\"Ej: Fac. Ing. Ambiental\""+
      " onfocus=\"this.style.borderColor='var(--accent)'\" onblur=\"this.style.borderColor='var(--line2)'\">"+
      "<div style=\"display:flex;gap:6px;flex-wrap:wrap;margin-top:8px\">"+presetBtns+"</div>"+
    "</div>"+
    "<div style=\"display:flex;align-items:flex-end;gap:10px;margin-bottom:16px\">"+
      "<div style=\"flex:2\">"+
        "<span style=\""+lbl+"\">A\u00f1o</span>"+
        "<input id=\"p-year\" type=\"text\" maxlength=\"4\" value=\""+periodYear+"\" style=\""+inpMono+"\""+
        " onfocus=\"this.style.borderColor='var(--accent)'\" onblur=\"this.style.borderColor='var(--line2)'\">"+
      "</div>"+
      "<div style=\"color:var(--tx3);font-size:1.4rem;font-weight:300;padding-bottom:10px;flex-shrink:0\">&mdash;</div>"+
      "<div style=\"flex:1\">"+
        "<span style=\""+lbl+"\">Per\u00edodo</span>"+
        "<input id=\"p-num\" type=\"text\" maxlength=\"1\" value=\""+periodNum+"\" style=\""+inpMono+"\""+
        " onfocus=\"this.style.borderColor='var(--accent)'\" onblur=\"this.style.borderColor='var(--line2)'\">"+
      "</div>"+
    "</div>"+
    "<button class=\"btn btn-dl\" style=\"width:100%;justify-content:center\" onclick=\"savePeriod()\">Guardar</button>";
  document.getElementById("modal").style.display="flex";
  setTimeout(function(){document.getElementById("p-faculty").focus();},100);
}


function savePeriod(){
  var f=document.getElementById("p-faculty").value.trim();
  var y=document.getElementById("p-year").value.trim();
  var n=document.getElementById("p-num").value.trim();
  if(!f||!y||!n){toast("Completa todos los campos","er");return;}
  facultyName=f;periodYear=y;periodNum=n;
  var match=FACULTY_PRESETS.find(function(p){return p.value===f;});
  facultyLabel=match?match.label:f;
  courses=[];sel={};pal={};palIdx=0;
  drawList();drawSched();drawTags();
  updatePeriodUI();
  modalClose();
  toast("Per\u00edodo actualizado: "+facultyName+" \u00b7 "+getPeriod(),"ok");
}


function openInfoModal(cod){
  const c=courses.find(function(x){return x.cod===cod;});
  const sec=sel[cod];
  if(!c||!sec) return;
  document.getElementById("mttl").textContent=c.curso;
  document.getElementById("msub").textContent=cod+" \u00b7 "+facultyName+" \u00b7 Secci\u00f3n "+sec.secc;
  const body=document.getElementById("mbody");
  const p=pal[cod]||"p0";
  let h="<div class=\"sec\" style=\"cursor:default;pointer-events:none\">";
  h+="<div class=\"sec-top\"><span class=\"sec-name\">Secci\u00f3n "+sec.secc+"</span>";
  h+="<span class=\"leg-dot "+p+"\" style=\"width:12px;height:12px;display:inline-block;border-radius:3px;margin-left:6px\"></span></div>";
  if(sec.docs&&sec.docs.length){
    sec.docs.forEach(function(d,di){
      h+="<div class=\"sec-doc\" style=\"display:flex;align-items:center;gap:8px;flex-wrap:wrap;pointer-events:auto;margin-bottom:3px\">";
      h+="<span style=\"flex:1;min-width:0\">"+d.name+"</span>";
      h+=_docRoleHTML(sec,d);
      h+="</div>";
      const rBtn=_ratingBtnHTML(cod,di);
      const rHtml=_ratingHTML(d.name);
      if(rBtn||rHtml){
        h+="<div style=\"display:flex;align-items:center;justify-content:space-between;gap:8px;margin:2px 0 8px;pointer-events:auto\">";
        h+="<span style=\"display:inline-flex;align-items:center;gap:5px\">"+rHtml+"</span>";
        h+=rBtn+"</div>";
      }
    });
  }
  sec.ss.forEach(function(s){
    h+="<div class=\"ses\"><span class=\"ses-"+s.t+"\">"+(TN[s.t]||s.t)+"</span>";
    h+="<span>"+(DN[s.d]||s.d)+" "+s.h0+":00\u2013"+s.h1+":00</span>";
    h+="<span style=\"color:var(--tx3)\">"+s.rm+"</span></div>";
  });
  h+="</div>";
  h+="<div style=\"display:flex;gap:8px;margin-top:14px;flex-wrap:wrap\">";
  h+="<button class=\"btn btn-clear\" style=\"flex:1;justify-content:center\" onclick=\"removeFromInfo('"+cod+"')\">";
  h+="<svg width=\"13\" height=\"13\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.5\"><polyline points=\"3 6 5 6 21 6\"/><path d=\"M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2\"/></svg>";
  h+=" Quitar curso</button>";
  if(c.secs.length>1){
    h+="<button class=\"btn\" style=\"flex:1;justify-content:center;background:var(--surface);border-color:var(--line2);color:var(--tx2)\" onclick=\"changeSecFromInfo('"+cod+"')\">";
    h+="<svg width=\"13\" height=\"13\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.5\"><path d=\"M1 4v6h6\"/><path d=\"M23 20v-6h-6\"/><path d=\"M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4-4.64 4.36A9 9 0 0 1 3.51 15\"/></svg>";
    h+=" Cambiar secci\u00f3n</button>";
  }
  h+="</div>";
  h+="<button class=\"btn\" style=\"width:100%;justify-content:center;margin-top:8px;background:var(--surface);border-color:var(--line2);color:var(--tx2)\" onclick=\"openEditScheduleModal('"+cod+"')\">";
  h+="<svg width=\"13\" height=\"13\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.5\"><path d=\"M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7\"/><path d=\"M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z\"/></svg>";
  h+=" Editar horario</button>";
  body.innerHTML=h;
  document.getElementById("modal").style.display="flex";
}


function openEditScheduleModal(cod){
  _editingCod=cod;
  const c=courses.find(function(x){return x.cod===cod;});
  const sec=sel[cod];
  if(!c||!sec) return;
  document.getElementById("mttl").textContent=c.curso;
  document.getElementById("msub").textContent="Editar sesiones \u00b7 Secci\u00f3n "+sec.secc;
  const body=document.getElementById("mbody");
  let h="<div id=\"ses-edit-list\" style=\"margin-bottom:10px\">";
  sec.ss.forEach(function(s,i){h+=_buildSesRow(i,s);});
  h+="</div>";
  h+="<button style=\"width:100%;padding:7px;background:transparent;border:1.5px dashed var(--line3);border-radius:8px;color:var(--tx3);font-size:0.78rem;cursor:pointer;margin-bottom:12px;font-family:'Plus Jakarta Sans',sans-serif\" onclick=\"addSesRow()\">";
  h+="+ A\u00f1adir sesi\u00f3n</button>";
  h+="<div style=\"display:flex;gap:8px\">";
  h+="<button class=\"btn\" style=\"flex:1;justify-content:center;background:var(--surface);border-color:var(--line2);color:var(--tx2)\" onclick=\"openInfoModal('"+cod+"')\">";
  h+="Cancelar</button>";
  h+="<button class=\"btn btn-dl\" style=\"flex:1;justify-content:center\" onclick=\"saveEditedSchedule()\">";
  h+="<svg width=\"13\" height=\"13\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.5\"><path d=\"M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z\"/><polyline points=\"17 21 17 13 7 13 7 21\"/><polyline points=\"7 3 7 8 15 8\"/></svg>";
  h+=" Guardar cambios</button>";
  h+="</div>";
  body.innerHTML=h;
  document.getElementById("modal").style.display="flex";
}


function _buildSesRow(i,s){
  var tOpts=["T","P","L","S"].map(function(v){return"<option value=\""+v+"\""+(s.t===v?" selected":"")+">"+(TN[v]||v)+"</option>";}).join("");
  var dOpts=["LU","MA","MI","JU","VI","SA"].map(function(v){return"<option value=\""+v+"\""+(s.d===v?" selected":"")+">"+(DN[v]||v)+"</option>";}).join("");
  return "<div class=\"ses-edit\" id=\"ses-row-"+i+"\" data-dc=\""+String(s.dc||"").replace(/"/g,"&quot;")+"\">"+
    "<select data-f=\"t\">"+tOpts+"</select>"+
    "<select data-f=\"d\">"+dOpts+"</select>"+
    "<input type=\"number\" data-f=\"h0\" min=\"6\" max=\"22\" value=\""+s.h0+"\" title=\"Hora inicio\">"+
    "<span class=\"ses-edit-sep\">\u2013</span>"+
    "<input type=\"number\" data-f=\"h1\" min=\"7\" max=\"23\" value=\""+s.h1+"\" title=\"Hora fin\">"+
    "<input type=\"text\" data-f=\"rm\" value=\""+(s.rm||"")+"\" placeholder=\"Aula\" maxlength=\"14\">"+
    "<button class=\"ses-edit-del\" onclick=\"removeSesRow("+i+")\">&#215;</button>"+
    "</div>";
}


function addSesRow(){
  const list=document.getElementById("ses-edit-list");
  if(!list) return;
  const i=list.querySelectorAll(".ses-edit").length;
  const div=document.createElement("div");
  div.innerHTML=_buildSesRow(i,{t:"T",d:"LU",h0:8,h1:10,rm:"",dc:(sel[_editingCod]||{}).docente||""});
  list.appendChild(div.firstChild);
}


function removeSesRow(i){
  const list=document.getElementById("ses-edit-list");
  if(!list) return;
  const rows=list.querySelectorAll(".ses-edit");
  if(rows.length<=1){toast("Debe haber al menos una sesi\u00f3n","er");return;}
  const row=document.getElementById("ses-row-"+i);
  if(row) row.remove();
  list.querySelectorAll(".ses-edit").forEach(function(r,idx){
    r.id="ses-row-"+idx;
    const btn=r.querySelector(".ses-edit-del");
    if(btn) btn.setAttribute("onclick","removeSesRow("+idx+")");
  });
}


function saveEditedSchedule(){
  const cod=_editingCod;
  if(!cod||!sel[cod]) return;
  const list=document.getElementById("ses-edit-list");
  if(!list) return;
  const rows=list.querySelectorAll(".ses-edit");
  const newSS=[];
  for(var i=0;i<rows.length;i++){
    const r=rows[i];
    const t=r.querySelector("[data-f='t']").value;
    const d=r.querySelector("[data-f='d']").value;
    const h0=parseInt(r.querySelector("[data-f='h0']").value);
    const h1=parseInt(r.querySelector("[data-f='h1']").value);
    const rm=r.querySelector("[data-f='rm']").value.trim();
    if(isNaN(h0)||isNaN(h1)||h1<=h0){toast("Hora inv\u00e1lida en sesi\u00f3n "+(i+1),"er");return;}
    if(h0<6||h1>23){toast("Hora fuera de rango en sesi\u00f3n "+(i+1),"er");return;}
    newSS.push({t:t,d:d,h0:h0,h1:h1,rm:rm,dc:r.getAttribute("data-dc")||""});
  }
  sel[cod].ss=newSS;
  _rebuildDocs(sel[cod]);
  _editingCod=null;
  document.getElementById("modal").style.display="none";
  sync();
  toast("Horario actualizado","ok");
}


function removeFromInfo(cod){
  delete sel[cod];delete pal[cod];
  document.getElementById("modal").style.display="none";
  sync();
}


function changeSecFromInfo(cod){
  const c=courses.find(function(x){return x.cod===cod;});
  if(!c) return;
  _pendingCod=cod;_pendingSec=sel[cod];
  delete sel[cod];
  openModal(c);
}
