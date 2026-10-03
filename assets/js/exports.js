/* Exportar el horario a PDF y a calendario (.ics). */

function downloadPDF(){
  const el=document.getElementById("printable");
  if(!el){toast("Primero selecciona cursos para generar el horario","er");return;}
  /* El tema que se ve en pantalla manda. No basta con lo guardado: en modo
     automático no se guarda nada y el PDF salía con el marco de otro tema. */
  const pdfC=PDF_THEME[currentTheme]||PDF_THEME[defaultTheme()]||PDF_THEME[THEME_ORDER[0]];
  const bgCanvas=pdfC.canvas;
  const bgOuter=pdfC.outer;
  const bgInner=pdfC.inner;
  const txtPri=pdfC.pri;
  const txtSec=pdfC.sec;
  toast("Generando PDF\u2026","ok");
  /* La maqueta de exportación (clase «exporting»: horario compacto y plano,
     sin cabecera fija ni leyenda) se aplica solo a la copia del documento
     que dibuja html2canvas, así la pantalla no cambia mientras se genera. */
  function prepare(doc){
    doc.body.classList.add("exporting");
    const p=doc.getElementById("printable");
    if(!p) return;
    p.querySelectorAll("thead th").forEach(function(th){th.style.position="static";});
    const lg=p.querySelector(".sched-legend");
    if(lg) lg.style.display="none";
  }
  /* Tamaño de esa maqueta, medido sin que llegue a pintarse (la clase se
     pone y se quita en la misma tarea; «v-measure» evita transiciones al
     volver). Con él se limita el lienzo a 16 MP, el tope de Safari en iOS. */
  const legend=el.querySelector(".sched-legend");
  document.body.classList.add("v-measure","exporting");
  if(legend) legend.style.display="none";
  const r=el.getBoundingClientRect();
  document.body.classList.remove("exporting");
  if(legend) legend.style.display="";
  void el.offsetWidth;
  document.body.classList.remove("v-measure");
  const sc=Math.min(2.8,Math.sqrt(16e6/Math.max(1,r.width*r.height)));
  /* Se espera a las tipografías: si alguna aún no ha cargado, la captura
     saldría con la de respaldo. La mono 500 se pide aparte porque es la que
     usa html2canvas para medir la línea base de la mono. */
  const fonts=document.fonts?Promise.all([document.fonts.ready,
    document.fonts.load('500 10px "IBM Plex Mono"',"A0").catch(function(){})]):Promise.resolve();
  fonts.then(function(){
    return html2canvas(el,{backgroundColor:bgCanvas,scale:sc,useCORS:true,logging:false,scrollX:0,scrollY:0,onclone:prepare});
  }).then(function(canvas){
    const W=canvas.width/sc, H=canvas.height/sc;
    /* Cabecera: título y subtítulo a tamaño legible (las unidades de jsPDF
       en «px» son puntos ×1,333); la tarjeta baja 6 px más que la imagen
       para que se vean sus cuatro esquinas redondeadas */
    const PH=H+98;
    const pdf=new window.jspdf.jsPDF({orientation:W>H?"landscape":"portrait",unit:"px",format:[W+56,PH]});
    pdf.setFillColor(...bgOuter);pdf.rect(0,0,W+56,PH,"F");
    pdf.setFillColor(...bgInner);pdf.roundedRect(28,16,W,H+54,6,6,"F");
    /* En Vuelo el nombre va en serif, como en la cabecera de la página */
    const serif=THEME_FAMILY[currentTheme]==="vuelo";
    pdf.setTextColor(...txtPri);pdf.setFontSize(20);pdf.setFont(serif?"times":"helvetica","bold");
    pdf.text(facultyLabel,38,38);
    pdf.setTextColor(...txtSec);pdf.setFontSize(13);pdf.setFont("helvetica","normal");
    pdf.text("HORARIO "+getPeriod()+"  \u00b7  Generado desde Horarios FIA "+getPeriod(),38,55);
    /* "FAST": la imagen va comprimida (antes pesaba 15–40 MB) */
    pdf.addImage(canvas.toDataURL("image/png"),"PNG",28,64,W,H,undefined,"FAST");
    pdf.save("horario-"+facultyLabel.split(" \u00b7 ")[0].toLowerCase()+"-"+getPeriod()+".pdf");
    toast("\u2713 PDF descargado correctamente","ok");
  }).catch(function(e){toast("Error al generar: "+e.message,"er");});
}


function exportICS(){
  var keys=Object.keys(sel);
  if(!keys.length){toast("Primero selecciona cursos para exportar el horario","er");return;}
  var dayMap={LU:"MO",MA:"TU",MI:"WE",JU:"TH",VI:"FR",SA:"SA"};
  var dayIdx={LU:1,MA:2,MI:3,JU:4,VI:5,SA:6};
  function pad(n){return String(n).padStart(2,"0");}
  function nextWeekday(code){
    var today=new Date();today.setHours(0,0,0,0);
    var diff=dayIdx[code]-today.getDay();
    if(diff<=0)diff+=7;
    var d=new Date(today);d.setDate(today.getDate()+diff);
    return d;
  }
  function fmtDT(date,hour){
    return date.getFullYear()+pad(date.getMonth()+1)+pad(date.getDate())+"T"+pad(hour)+"0000";
  }
  var lines=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Horarios FIA//ES","CALSCALE:GREGORIAN","X-WR-CALNAME:Horario "+currentFaculty+" "+getPeriod()];
  keys.forEach(function(cod){
    var s=sel[cod];
    s.ss.forEach(function(x,i){
      var d=nextWeekday(x.d);
      lines.push("BEGIN:VEVENT");
      lines.push("UID:"+cod+"-"+x.d+"-"+x.h0+"-"+i+"@horarios-fia");
      lines.push("DTSTART:"+fmtDT(d,x.h0));
      lines.push("DTEND:"+fmtDT(d,x.h1));
      lines.push("RRULE:FREQ=WEEKLY;BYDAY="+dayMap[x.d]);
      lines.push("SUMMARY:"+(s.curso||cod)+" ("+s.secc+") \u2013 "+(TN[x.t]||x.t));
      if(x.rm)lines.push("LOCATION:"+x.rm);
      lines.push("DESCRIPTION:Docente: "+(x.dc||s.docente||"\u2013")+"\\nC\u00f3digo: "+cod);
      lines.push("END:VEVENT");
    });
  });
  lines.push("END:VCALENDAR");
  var blob=new Blob([lines.join("\r\n")],{type:"text/calendar;charset=utf-8"});
  var url=URL.createObjectURL(blob);
  var a=document.createElement("a");a.href=url;a.download="horario-"+currentFaculty.toLowerCase()+"-"+getPeriod()+".ics";a.click();
  URL.revokeObjectURL(url);
  toast("\u2713 Archivo .ics descargado \u2014 \u00e1brelo para importar a tu calendario","ok");
}
