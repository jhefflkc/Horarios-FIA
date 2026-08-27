/* Facultad activa, periodo y carga de un Excel desde el navegador. */

function getPeriod(){return ALL_DATA[currentFaculty].period;}

function updatePeriodUI(){
  var p=getPeriod();
  document.getElementById("period-label").textContent="\u25c6 Cursos disponibles \u2014 "+p;
  document.getElementById("logo-sub-period").textContent="Horarios \u00b7 "+p;
  var lt=document.getElementById("logo-text-faculty");
  if(lt) lt.innerHTML=facultyLabel.replace(" \u00b7 "," <span>&middot;</span> ");
  var fs=document.getElementById("faculty-selector");
  if(fs) fs.querySelectorAll("button[data-sigla]").forEach(function(btn){
    var active=btn.dataset.sigla===currentFaculty;
    btn.style.borderColor=active?"var(--accent)":"var(--line2)";
    btn.style.color=active?"var(--accent3)":"var(--tx2)";
  });
}

function initFacultySelector(){
  var container=document.getElementById("faculty-selector");
  if(!container) return;
  var siglas=Object.keys(ALL_DATA);
  if(siglas.length<=1){container.style.display="none";return;}
  container.style.display="flex";
  var h="";
  siglas.forEach(function(sigla){
    h+="<button data-sigla=\""+sigla+"\" onclick=\"switchFaculty('"+sigla+"')\" "+
      "style=\"background:var(--surface);border:1.5px solid var(--line2);border-radius:7px;"+
      "padding:4px 10px;font-size:0.68rem;font-weight:700;color:var(--tx2);cursor:pointer;"+
      "font-family:'JetBrains Mono',monospace;white-space:nowrap;transition:all 0.15s\""+
      " onmouseover=\"this.style.borderColor='var(--accent2)';this.style.color='var(--accent3)'\""+
      " onmouseout=\"if(this.dataset.sigla!==currentFaculty){this.style.borderColor='var(--line2)';this.style.color='var(--tx2)'}\">"+
      sigla+"</button>";
  });
  container.innerHTML=h;
}

function switchFaculty(sigla){
  if(!ALL_DATA[sigla]) return;
  currentFaculty=sigla;
  var meta=ALL_DATA[sigla];
  facultyName=meta.fullName;
  facultyLabel=meta.label;
  boot(meta.rows);
  updatePeriodUI();
}


function loadExcel(event){
  const file=event.target.files[0];if(!file) return;
  event.target.value="";
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const rows=[];
      const TM={TEORIA:"T","TEOR\u00cdA":"T",PRACTICA:"P","PR\u00c1CTICA":"P",LABORATORIO:"L",LAB:"L",SEMINARIO:"S"};
      wb.SheetNames.forEach(function(name){
        const ws=wb.Sheets[name];
        const json=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        let hi=-1;
        for(let i=0;i<json.length;i++){
          const r=json[i].map(function(c){return String(c).toUpperCase().trim();});
          if(r.indexOf("COD")>=0&&r.indexOf("CURSO")>=0){hi=i;break;}
        }
        if(hi<0) return;
        /* Compara cabeceras sin tildes, puntos ni espacios: «Vac. máximas»,
           «VAC MAXIMAS» y «Vacantes Máximas» son la misma columna. */
        const norm=function(t){return String(t).normalize("NFD").replace(/[\u0300-\u036f]/g,"")
          .toUpperCase().replace(/[^A-Z0-9]/g,"");};
        const hs=json[hi].map(norm);
        const ix=function(){for(let k=0;k<arguments.length;k++){
          const j=hs.indexOf(norm(arguments[k])); if(j>=0) return j;} return -1;};
        const iC=ix("COD"),iCu=ix("CURSO"),iS=ix("SECC"),iTi=ix("TIPO"),iH=ix("HORARIO"),iD=ix("DIA"),iHi=ix("H INI"),iHf=ix("H FIN"),iSa=ix("AULA")>=0?ix("AULA"):ix("SALON"),iDo=ix("DOCENTE"),iCy=ix("CICLO","CICLO(S)","CICLOS"),
              iVm=ix("VAC. MAXIMAS","VACANTES MAXIMAS","VAC MAXIMAS","VACANTES");
        const fiaFmt=iH>=0;
        for(let i=hi+1;i<json.length;i++){
          const r=json[i];
          if(!r[iC]||!r[iCu]) continue;
          const cod=String(r[iC]).trim(),secc=String(r[iS]||"").trim().toUpperCase();
          const tipo=TM[String(r[iTi]||"").trim().toUpperCase()]||"T";
          const ciclo=parseCiclo(iCy>=0?r[iCy]:null,cod);
          const salon=iSa>=0?String(r[iSa]||"").trim():"";
          const docente=iDo>=0?String(r[iDo]||"").trim():"";
          const curso=String(r[iCu]).trim();
          const vm=iVm>=0?parseInt(r[iVm]):NaN;
          const extra=isNaN(vm)?{}:{vacMax:vm};
          if(fiaFmt){
            const ph=parseHorario(r[iH]);if(!ph) continue;
            rows.push(Object.assign({cod:cod,curso:curso,secc:secc,tipo:tipo,ciclo:ciclo,dia:ph.dia,hIni:ph.hIni,hFin:ph.hFin,salon:salon,docente:docente},extra));
          } else {
            if(!r[iD]) continue;
            const h0=parseInt(r[iHi]),h1=parseInt(r[iHf]);
            if(isNaN(h0)||isNaN(h1)) continue;
            rows.push(Object.assign({cod:cod,curso:curso,secc:secc,tipo:tipo,ciclo:ciclo,dia:String(r[iD]).trim().toUpperCase(),hIni:h0,hFin:h1,salon:salon,docente:docente},extra));
          }
        }
      });
      if(!rows.length){toast("No se encontraron datos v\u00e1lidos","er");return;}
      // Detectar facultad y periodo desde el nombre del archivo (ej: FIA2026-1.xlsx)
      const m=file.name.match(/^([A-Za-z]+)(\d{4})-(\d+)\.xlsx?$/i);
      if(m){
        const sigla=m[1].toUpperCase();
        const period=m[2]+"-"+m[3];
        const fmeta=FACULTY_MAP_JS[sigla];
        if(fmeta){
          ALL_DATA[sigla]={label:fmeta.label,fullName:fmeta.fullName,period:period,rows:rows};
          initFacultySelector();
          switchFaculty(sigla);
          toast("\u2713 "+rows.length+" sesiones \u2014 "+fmeta.label+" "+period,"ok");
          return;
        }
      }
      // Fallback: reemplaza la facultad actual sin cambiar selector
      ALL_DATA[currentFaculty].rows=rows;
      boot(rows);
      toast("\u2713 "+rows.length+" sesiones cargadas","ok");
    }catch(e){toast("Error: "+e.message,"er");}
  };
  reader.readAsBinaryString(file);
}
