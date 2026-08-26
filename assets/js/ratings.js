/* Calificaciones de docentes (Google Apps Script) y píldora de Score. */

// ─── Rating helpers ────────────────────────────────────────────────────────
function _isFIA(){return currentFaculty==="FIA";}


function loadRatings(){
  if(!_isFIA()) return Promise.resolve();
  if(_ratingsCache!==null) return Promise.resolve();
  if(!RATINGS_CFG.webAppUrl){_ratingsCache={};return Promise.resolve();}
  if(!_ratingsPromise){
    _ratingsPromise=fetch(RATINGS_CFG.webAppUrl)
      .then(function(r){return r.json();})
      .catch(function(){return {};})
      .then(function(data){_ratingsCache=data;updateScorePill();});
  }
  return _ratingsPromise;
}


function _ratingHTML(teacher){
  if(!_isFIA()||!teacher||!_ratingsCache) return "";
  const r=_ratingsCache[teacher];
  if(!r) return "";
  return "<span class=\"rate-star\"><svg width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" fill=\"currentColor\"><path d=\"M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z\"/></svg>"+r.avg+"</span>"+
    "<span class=\"rate-count\">("+r.count+")</span>";
}


function updateScorePill(){
  const pill=document.getElementById("score-pill");
  if(!pill) return;
  if(!_isFIA()||!_ratingsCache){pill.style.display="none";return;}
  const keys=Object.keys(sel);
  if(!keys.length){pill.style.display="none";return;}
  let sum=0,count=0;
  keys.forEach(function(cod){
    const sec=sel[cod]||{};
    const names=(sec.docs&&sec.docs.length)
      ?sec.docs.map(function(d){return d.name;})
      :(sec.docente?[sec.docente]:[]);
    /* Un curso con varios docentes aporta el promedio de los que estén evaluados */
    let s=0,n=0;
    names.forEach(function(t){
      const r=_ratingsCache[t];
      if(r&&r.avg){s+=parseFloat(r.avg);n++;}
    });
    if(n){sum+=s/n;count++;}
  });
  if(!count){pill.style.display="none";return;}
  document.getElementById("score-val").textContent=(sum/count).toFixed(1);
  document.getElementById("score-pop-count").textContent=count+(count===1?" curso evaluado":" cursos evaluados")+" para el Score";
  pill.style.display="inline-flex";
}


function toggleScorePop(e){
  e.stopPropagation();
  const pop=document.getElementById("score-pop");
  if(pop.style.display==="block"){pop.style.display="none";return;}
  /* Mostrar fuera de pantalla para leer dimensiones reales */
  pop.style.visibility="hidden";pop.style.left="-9999px";pop.style.top="-9999px";pop.style.bottom="auto";pop.style.display="block";
  const pw=pop.offsetWidth;const ph=pop.offsetHeight;
  pop.style.visibility="";pop.style.top="auto";
  const r=document.getElementById("score-pill").getBoundingClientRect();
  const margin=8;
  let left=r.left+r.width/2-pw/2;
  if(left+pw>window.innerWidth-margin) left=window.innerWidth-pw-margin;
  if(left<margin) left=margin;
  const bottomSpace=window.innerHeight-r.top;
  pop.style.left=left+"px";
  pop.style.bottom=bottomSpace+margin+"px";
}


/* Docente nº di de la sección elegida para `cod` (0 por defecto) */
function _docOf(cod,di){
  const sec=sel[cod];if(!sec) return "";
  return ((sec.docs||[])[di||0]||{}).name||sec.docente||"";
}


function _ratingBtnHTML(cod,di){
  if(!_isFIA()) return "";
  if(!RATINGS_CFG.webAppUrl&&!RATINGS_CFG.formUrl) return "";
  const teacher=_docOf(cod,di);if(!teacher) return "";
  if(_hasRated(teacher)) return "<span class=\"rate-count\">&#10003; Ya calificaste</span>";
  return "<button class=\"rate-btn\" onclick=\"openRatingModal('"+cod+"',"+(di||0)+")\">&#9733; Calificar</button>";
}


function _hasRated(teacher){
  try{return!!(JSON.parse(localStorage.getItem("fia_rated")||"{}")[teacher]);}catch(e){return false;}
}

function _markRated(teacher){
  try{const r=JSON.parse(localStorage.getItem("fia_rated")||"{}");r[teacher]=Date.now();localStorage.setItem("fia_rated",JSON.stringify(r));}catch(e){}
}


const _STAR_LABELS=["","Muy malo","Malo","Regular","Bueno","Excelente"];


function openRatingModal(cod,di){
  const sec=sel[cod];if(!sec) return;
  const teacher=_docOf(cod,di);if(!teacher) return;
  _ratingTeacher=teacher;_selectedRating=0;
  document.getElementById("mttl").textContent="Calificar docente";
  const d=(sec.docs||[])[di||0];
  const role=(d&&sec.docs.length>1&&d.tipos.length)?" · "+_docRole(d.tipos):"";
  document.getElementById("msub").textContent=teacher+role;
  const body=document.getElementById("mbody");
  const r=(_ratingsCache||{})[teacher];
  let h="";
  if(r) h+="<div style=\"text-align:center;margin-bottom:10px;font-size:0.77rem;color:var(--tx3)\">Promedio actual: <strong style=\"color:#f0c870\">&#9733; "+r.avg+"</strong> <span style=\"color:var(--tx3)\">("+r.count+" votos)</span></div>";
  h+="<div style=\"text-align:center;padding:10px 0 4px\">";
  h+="<div style=\"font-size:0.78rem;color:var(--tx3);margin-bottom:16px\">&#191;C&#243;mo calificar&#237;as al docente?</div>";
  h+="<div id=\"star-inp\" style=\"display:flex;justify-content:center;gap:12px;font-size:2.4rem;margin-bottom:8px\">";
  for(var i=1;i<=5;i++) h+="<span class=\"star-inp\" data-v=\""+i+"\" onclick=\"selectStar("+i+")\" onmouseover=\"hoverStar("+i+")\" onmouseout=\"hoverStar(0)\">&#9734;</span>";
  h+="</div>";
  h+="<div id=\"star-lbl\" style=\"font-size:0.78rem;color:var(--accent);font-weight:700;height:1.3em;margin-bottom:16px\"></div>";
  h+="</div>";
  h+="<div style=\"display:flex;gap:8px\">";
  h+="<button class=\"btn\" style=\"flex:1;justify-content:center;background:var(--surface);border-color:var(--line2);color:var(--tx2)\" onclick=\"openInfoModal('"+cod+"')\">Cancelar</button>";
  h+="<button id=\"star-submit\" class=\"btn btn-dl\" style=\"flex:1;justify-content:center;opacity:0.4;pointer-events:none\" onclick=\"confirmRating('"+cod+"')\">";
  h+="<svg width=\"12\" height=\"12\" viewBox=\"0 0 24 24\" fill=\"currentColor\"><path d=\"M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z\"/></svg> Enviar calificaci&#243;n</button>";
  h+="</div>";
  body.innerHTML=h;
  document.getElementById("modal").style.display="flex";
}


function hoverStar(v){
  const active=v||_selectedRating;
  document.querySelectorAll(".star-inp").forEach(function(s,i){
    s.textContent=i<active?"\u2605":"\u2606";
    s.className="star-inp"+(i<active?" on":"");
    s.style.transform=(i===v-1)?"scale(1.28)":"scale(1)";
  });
  const lbl=document.getElementById("star-lbl");
  if(lbl) lbl.textContent=v?_STAR_LABELS[v]:(_selectedRating?_STAR_LABELS[_selectedRating]:"");
}


function selectStar(v){
  _selectedRating=v;hoverStar(v);
  const btn=document.getElementById("star-submit");
  if(btn){btn.style.opacity="1";btn.style.pointerEvents="auto";}
}


async function confirmRating(cod){
  const teacher=_ratingTeacher;if(!_selectedRating||!teacher) return;
  try{
    if(RATINGS_CFG.formUrl){
      const f=RATINGS_CFG.fields;
      const body=new URLSearchParams();
      body.append(f.docente,teacher);
      body.append(f.puntuacion,String(_selectedRating));
      body.append(f.curso,cod);
      await fetch(RATINGS_CFG.formUrl,{method:"POST",mode:"no-cors",body});
    } else if(RATINGS_CFG.webAppUrl){
      await fetch(RATINGS_CFG.webAppUrl+"?action=rate&docente="+encodeURIComponent(teacher)+"&puntuacion="+_selectedRating+"&curso="+encodeURIComponent(cod));
    }
  }catch(e){}
  // Actualiza caché local de forma optimista
  if(_ratingsCache){
    const old=_ratingsCache[teacher]||{avg:0,count:0};
    const nc=old.count+1;
    _ratingsCache[teacher]={avg:Math.round(((old.avg*old.count)+_selectedRating)/nc*10)/10,count:nc};
  }
  _markRated(teacher);
  _ratingTeacher=null;_selectedRating=0;
  toast("\u2713 \u00a1Gracias por tu calificaci\u00f3n!","ok");
  openInfoModal(cod);
}
