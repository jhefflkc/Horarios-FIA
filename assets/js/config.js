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

/* Tema Vuelo: el color va como texto de la etiqueta, así que cada modo usa
   la versión del tono que contrasta con su fondo (ver theme-vuelo.css). */
const PAL_HEX_VUELO_CLARO={p0:"#574514",p1:"#40436d",p2:"#075246",p3:"#3f552c",p4:"#693840",p5:"#29526d",p6:"#6d422a",p7:"#604369"};
const PAL_HEX_VUELO_OSCURO={p0:"#f0deb3",p1:"#d6dcff",p2:"#b1efe0",p3:"#c1d8b0",p4:"#ffd2d9",p5:"#a9d6f7",p6:"#f1c5ad",p7:"#e4c4ee"};

/* Paleta de etiquetas por tema; los que no figuran usan PAL_HEX */
const PAL_HEX_BY_THEME={google:PAL_HEX_GOOGLE,"vuelo-claro":PAL_HEX_VUELO_CLARO,"vuelo-oscuro":PAL_HEX_VUELO_OSCURO};

/* Color de una etiqueta según el tema activo */
function palHex(p){
  const pal=PAL_HEX_BY_THEME[typeof currentTheme!=="undefined"?currentTheme:""]||PAL_HEX;
  return pal[p]||pal.p0;
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


/* ─── Temas ──────────────────────────────────────────────────────────────
   Activos: solo el tema «Vuelo», en claro y oscuro. Los cinco anteriores
   siguen definidos (CSS y configuración) pero fuera de la rotación; para
   recuperarlos basta añadirlos a THEME_ORDER. Copia de seguridad y
   instrucciones en respaldo/temas-anteriores/. */
var THEME_ORDER=["vuelo-claro","vuelo-oscuro"];
var THEMES_LEGACY=["dark","stitch-dark","light","stitch-light","google"];

/* Familia: clase común que comparten varios temas (estructura compartida;
   cada modo solo cambia los tokens de color) */
var THEME_FAMILY={"vuelo-claro":"vuelo","vuelo-oscuro":"vuelo"};

/* Tema de entrada para quien no haya elegido uno. "auto" sigue la
   preferencia claro/oscuro del sistema, y la sigue en vivo mientras el
   usuario no toque el botón. */
var DEFAULT_THEME="auto";
var AUTO_THEMES={light:"vuelo-claro",dark:"vuelo-oscuro"};

/* Agrupar la lista de cursos por ciclo, con una cabecera por cada uno.
   Desactivado por ahora: ponlo en true para recuperarlo. El ciclo se sigue
   leyendo del Excel, así que no hace falta nada más. */
var GROUP_BY_CYCLE=false;

var THEME_LABELS={"vuelo-claro":"Claro","vuelo-oscuro":"Oscuro",
  dark:"Ámbar",["stitch-dark"]:"Grafito",light:"Crema",["stitch-light"]:"Glacial",google:"Google"};

var THEME_IS_DARK={"vuelo-claro":false,"vuelo-oscuro":true,
  dark:true,["stitch-dark"]:true,light:false,["stitch-light"]:false,google:false};

/* Qué icono muestra el botón de tema en cada caso (sol en claro, luna en oscuro) */
var THEME_ICON={"vuelo-claro":"dark","vuelo-oscuro":"light",
  dark:"light",light:"dark",["stitch-dark"]:"stitch",["stitch-light"]:"stitch",google:"google"};

/* Colores del PDF por tema. html2canvas no hereda el tema de la página, así
   que el marco del PDF se pinta a mano: cada tema necesita su entrada aquí o
   el PDF saldrá con los colores de otro. */
var PDF_THEME={
  ["vuelo-claro"]: {canvas:"#fafbfc",outer:[238,241,244],inner:[250,251,252],pri:[29,35,39],   sec:[90,100,105]},
  ["vuelo-oscuro"]:{canvas:"#363638",outer:[49,49,51],   inner:[54,54,56],   pri:[249,248,248],sec:[158,157,153]},
  dark:            {canvas:"#0c0905",outer:[12,9,5],     inner:[19,14,8],    pri:[200,150,100],sec:[110,100,80]},
  ["stitch-dark"]: {canvas:"#0d0d0f",outer:[13,13,15],   inner:[17,17,19],   pri:[229,229,231],sec:[142,142,147]},
  light:           {canvas:"#fdf6ee",outer:[253,246,238],inner:[255,250,244],pri:[90,58,24],   sec:[138,98,72]},
  ["stitch-light"]:{canvas:"#efefef",outer:[239,239,239],inner:[255,255,255],pri:[28,28,30],   sec:[99,99,102]},
  google:          {canvas:"#f0f4f9",outer:[240,244,249],inner:[255,255,255],pri:[31,31,31],   sec:[95,99,104]}
};

/* Panel colors per theme — must match CSS --panel values.
   html { background } can't inherit --panel from body, so we set it directly. */
var THEME_PANEL={"vuelo-claro":"#eef1f4","vuelo-oscuro":"#313133",
  dark:"#130e08",["stitch-dark"]:"#0d0d0f",light:"#fffaf4",["stitch-light"]:"#efefef",google:"#f0f4f9"};
