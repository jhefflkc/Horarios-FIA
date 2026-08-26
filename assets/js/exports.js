/* Exportar el horario a PDF y a calendario (.ics). */

function downloadPDF(){
  const el=document.getElementById("printable");
  if(!el){toast("Primero selecciona cursos para generar el horario","er");return;}
  /* El tema activo manda: antes se comprobaban clases sueltas y cualquier
     tema no contemplado —como Google— salía con el marco oscuro. */
  const pdfC=PDF_THEME[localStorage.getItem("theme")||"dark"]||PDF_THEME.dark;
  const bgCanvas=pdfC.canvas;
  const bgOuter=pdfC.outer;
  const bgInner=pdfC.inner;
  const txtPri=pdfC.pri;
  const txtSec=pdfC.sec;
  toast("Generando PDF\u2026","ok");
  /* Temporarily remove sticky positioning so html2canvas captures the full table correctly */
  const stickyEls=el.querySelectorAll("thead th");
  stickyEls.forEach(function(th){th.style.position="static";});
  /* Hide legend \u2014 not needed in PDF since name mode is available */
  const legend=el.querySelector(".sched-legend");
  if(legend) legend.style.display="none";
  /* Congela las animaciones: html2canvas captura el estado del momento y
     un bloque a medio aparecer saldr\u00eda trasl\u00facido o desplazado en el PDF. */
  document.body.classList.add("exporting");
  html2canvas(el,{backgroundColor:bgCanvas,scale:2.8,useCORS:true,logging:false,scrollX:0,scrollY:0}).then(function(canvas){
    document.body.classList.remove("exporting");
    stickyEls.forEach(function(th){th.style.position="";});
    if(legend) legend.style.display="";
    const W=canvas.width/2.8, H=canvas.height/2.8;
    const pdf=new window.jspdf.jsPDF({orientation:W>H?"landscape":"portrait",unit:"px",format:[W+56,H+72]});
    pdf.setFillColor(...bgOuter);pdf.rect(0,0,W+56,H+72,"F");
    pdf.setFillColor(...bgInner);pdf.roundedRect(28,16,W,H+28,6,6,"F");
    pdf.setTextColor(...txtPri);pdf.setFontSize(9);pdf.setFont("helvetica","bold");
    pdf.text(facultyLabel,36,28);
    pdf.setTextColor(...txtSec);pdf.setFontSize(7);pdf.setFont("helvetica","normal");
    pdf.text("HORARIO "+getPeriod()+"  \u00b7  Generado desde Horarios FIA "+getPeriod(),36,36);
    pdf.addImage(canvas.toDataURL("image/png"),"PNG",28,44,W,H);
    pdf.save("horario-"+facultyLabel.split(" \u00b7 ")[0].toLowerCase()+"-"+getPeriod()+".pdf");
    toast("\u2713 PDF descargado correctamente","ok");
  }).catch(function(e){document.body.classList.remove("exporting");stickyEls.forEach(function(th){th.style.position="";});if(legend) legend.style.display="";toast("Error al generar: "+e.message,"er");});
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
