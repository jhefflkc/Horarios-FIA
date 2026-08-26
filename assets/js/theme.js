/* Cambio de tema y persistencia en localStorage. */

function applyTheme(t){
  document.body.classList.remove("light","stitch-dark","stitch-light");
  if(t!=="dark") document.body.classList.add(t);
  document.documentElement.style.background=THEME_PANEL[t]||"#130e08";
  localStorage.setItem("theme",t);
  var dark=THEME_IS_DARK[t];
  var stitch=t==="stitch-dark"||t==="stitch-light";
  document.getElementById("theme-icon-dark").style.display=(!dark&&!stitch)?"":"none";
  document.getElementById("theme-icon-light").style.display=(dark&&!stitch)?"":"none";
  document.getElementById("theme-icon-stitch").style.display=stitch?"":"none";
  document.getElementById("theme-label").textContent=THEME_LABELS[t]||t;
}


function toggleTheme(){
  var cur=localStorage.getItem("theme")||"dark";
  var idx=THEME_ORDER.indexOf(cur);
  applyTheme(THEME_ORDER[(idx+1)%THEME_ORDER.length]);
}
