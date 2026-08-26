/* Ajustes y constantes.
   Aquí se añade una facultad nueva, se cambia la paleta de colores
   o la URL del servicio de calificaciones. */

const DAYS=["LU","MA","MI","JU","VI","SA"];

const DN={LU:"Lunes",MA:"Martes",MI:"Mi\u00e9rcoles",JU:"Jueves",VI:"Viernes",SA:"S\u00e1bado"};

const TN={T:"Teor\u00eda",P:"Pr\u00e1ctica",L:"Laboratorio",S:"Seminario"};

const _TORD={T:0,P:1,L:2,S:3};

const PAL=["p0","p1","p2","p3","p4","p5","p6","p7"];

const ESP_ORDER=["CB","IA","IS","IH"];

const ESP_LAB={CB:"CIENCIAS B\u00c1SICAS",IA:"INGENIER\u00cdA AMBIENTAL",IS:"INGENIER\u00cdA SANITARIA",IH:"ING. HIGIENE Y SEGURIDAD"};

const ESP_SH={CB:"B\u00e1sicas",IA:"I. Ambiental",IS:"I. Sanitaria",IH:"I. Higiene y Seg."};

const PAL_HEX={p0:"#c86432",p1:"#46883c",p2:"#c8a01e",p3:"#b45028",p4:"#788c28",p5:"#288c6e",p6:"#b4821e",p7:"#508232"};

/* Tema Google (claro): las etiquetas llevan el color como texto sobre fondo
   blanco, así que usan los tonos 600/700 de la paleta de Google —los mismos
   matices que los bloques pastel del horario, pero legibles. */
const PAL_HEX_GOOGLE={p0:"#1a73e8",p1:"#1e8e3e",p2:"#e37400",p3:"#d93025",p4:"#007b83",p5:"#9334e6",p6:"#e8710a",p7:"#00796b"};

/* Color de una etiqueta según el tema activo */
function palHex(p){
  if(document.body.classList.contains("google")) return PAL_HEX_GOOGLE[p]||PAL_HEX_GOOGLE.p0;
  return PAL_HEX[p]||PAL_HEX.p0;
}


// ─── Calificaciones de docentes (Google Apps Script) ───────────────────────
// Pega aquí la URL del Web App de Google Apps Script tras desplegarlo
const RATINGS_CFG={
  webAppUrl:"https://script.google.com/macros/s/AKfycbxrfYVwyuG7cjv9iP1QPOc1Fff4LY_2Lw9ep4FznyTAumFqCJJN_6NIBZ-BCXPh5-T5/exec",
  formUrl:"https://docs.google.com/forms/d/e/1FAIpQLSe6MiBUwLFktcGNHW6ZWDbVvXUqNa0pzpZWVpICGtM2_myhIA/formResponse",
  fields:{docente:"entry.1423576837",puntuacion:"entry.787681983",curso:"entry.1242329604"}
};

var FACULTY_MAP_JS={
  "FIEECS":{label:"FIEECS \u00b7 UNI",fullName:"Fac. Ing. El\u00e9ctrica y Electr\u00f3nica"},
  "FIGMM": {label:"FIGMM \u00b7 UNI", fullName:"Fac. Ing. Geol\u00f3gica, Minera y Metal\u00fargica"},
  "FIQT":  {label:"FIQT \u00b7 UNI",  fullName:"Fac. Ing. Qu\u00edmica y Textil"},
  "FIIS":  {label:"FIIS \u00b7 UNI",  fullName:"Fac. Ing. Industrial y Sistemas"},
  "FIEE":  {label:"FIEE \u00b7 UNI",  fullName:"Fac. Ing. El\u00e9ctrica y Electr\u00f3nica"},
  "FIPP":  {label:"FIPP \u00b7 UNI",  fullName:"Fac. Ing. Petr\u00f3leo, Gas Natural y Petroqu\u00edmica"},
  "FIM":   {label:"FIM \u00b7 UNI",   fullName:"Fac. Ing. Mec\u00e1nica"},
  "FIC":   {label:"FIC \u00b7 UNI",   fullName:"Fac. Ing. Civil"},
  "FIA":   {label:"FIA \u00b7 UNI",   fullName:"Fac. Ing. Ambiental"},
  "FC":    {label:"FC \u00b7 UNI",    fullName:"Fac. de Ciencias"},
  "FAUA":  {label:"FAUA \u00b7 UNI",  fullName:"Fac. de Arquitectura, Urbanismo y Artes"}
};


var THEME_ORDER=["dark","stitch-dark","light","stitch-light","google"];

var THEME_LABELS={dark:"Ámbar",["stitch-dark"]:"Grafito",light:"Crema",["stitch-light"]:"Glacial",google:"Google"};

var THEME_IS_DARK={dark:true,["stitch-dark"]:true,light:false,["stitch-light"]:false,google:false};

/* Qué icono muestra el botón de tema en cada caso */
var THEME_ICON={dark:"light",light:"dark",["stitch-dark"]:"stitch",["stitch-light"]:"stitch",google:"google"};

/* Panel colors per theme — must match CSS --panel values.
   html { background } can't inherit --panel from body, so we set it directly. */
var THEME_PANEL={dark:"#130e08",["stitch-dark"]:"#0d0d0f",light:"#fffaf4",["stitch-light"]:"#efefef",google:"#f0f4f9"};
