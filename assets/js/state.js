/* Estado en memoria de la sesión (cursos cargados, selección actual). */

let courses=[], sel={}, pal={}, palIdx=0;

let _pendingCod=null, _pendingSec=null;

let _editingCod=null;


let _ratingsCache=null;   // {"Dr. García":{avg:4.2,count:15}, ...}

let _ratingsPromise=null;

var _selectedRating=0;

var _ratingTeacher=null;

let currentFaculty=null;

let facultyName="",facultyLabel="";

let showNames=false;
