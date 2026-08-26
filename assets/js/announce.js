/* Anuncios: descarga ANNOUNCE.md y lo convierte de Markdown a HTML. */

// ── Announcements ──────────────────────────────────────────────────────────
function _annHash(s){
  var h=0;
  for(var i=0;i<Math.min(s.length,3000);i++) h=(Math.imul(31,h)+s.charCodeAt(i))|0;
  return h.toString(36)+":"+s.length;
}


function md2html(md){
  function esc(s){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
  function attr(s){return s.replace(/"/g,"&quot;");}
  /* Solo http(s)/mailto/rutas relativas: bloquea javascript: y similares */
  function url(u){
    u=u.trim();
    if(/^[a-z][a-z0-9+.\-]*:/i.test(u)&&!/^(https?:|mailto:|data:image\/)/i.test(u)) return "#";
    return attr(u);
  }
  function inline(s){
    /* \* \_ etc. → literal: se aparcan como marcador y se restauran al final */
    var lit=[];
    s=s.replace(/\\([\\`*_~\[\]()#+\-.!>])/g,function(m,ch){
      lit.push(ch);return "\uE000"+(lit.length-1)+"\uE000";
    });
    s=s.replace(/`([^`]+)`/g,"<code>$1</code>");
    /* El delimitador debe pegar con texto (no " * "), y _ no parte palabras
       (así FIA_2026_1.xlsx o 5 * 3 * 2 quedan intactos) */
    s=s.replace(/\*\*\*(\S(?:[\s\S]*?\S)?)\*\*\*/g,"<strong><em>$1</em></strong>");
    s=s.replace(/\*\*(\S(?:[\s\S]*?\S)?)\*\*/g,"<strong>$1</strong>");
    s=s.replace(/(^|[^\w\\])__(\S(?:[\s\S]*?\S)?)__(?!\w)/g,"$1<strong>$2</strong>");
    s=s.replace(/\*(\S(?:[^*\n]*?\S)?)\*/g,"<em>$1</em>");
    s=s.replace(/(^|[^\w\\])_(\S(?:[^_\n]*?\S)?)_(?!\w)/g,"$1<em>$2</em>");
    s=s.replace(/~~(\S(?:[\s\S]*?\S)?)~~/g,"<del>$1</del>");
    s=s.replace(/!\[([^\]]*)\]\(([^)]+)\)/g,function(m,a,u){
      return '<img alt="'+attr(a)+'" src="'+url(u)+'">';
    });
    s=s.replace(/\[([^\]]+)\]\(([^)]+)\)/g,function(m,t,u){
      return '<a href="'+url(u)+'" target="_blank" rel="noopener">'+t+'</a>';
    });
    return s.replace(/\uE000(\d+)\uE000/g,function(m,i){return lit[i];});
  }
  /* Recorrido línea a línea: un título, una regla o una lista cortan el
     párrafo aunque no haya línea en blanco antes, así nunca se pierde texto. */
  var HR=/^\s*([-*_])\s*(\1\s*){2,}$/;
  var UL=/^\s*[-*+]\s+/;
  var OL=/^\s*\d+[.)]\s+/;
  var QT=/^\s*>\s?/;
  var lines=md.replace(/\r\n/g,"\n").split("\n");
  var out="",para=[];
  function txt(l){return inline(esc(l));}
  function flush(){
    if(!para.length) return;
    out+="<p>"+para.map(txt).join("<br>")+"</p>";
    para=[];
  }
  for(var i=0;i<lines.length;i++){
    var ln=lines[i];
    if(!ln.trim()){flush();continue;}
    var hm=ln.match(/^\s*(#{1,3})\s+(.+?)\s*#*$/);
    if(hm){flush();out+="<h"+hm[1].length+">"+txt(hm[2])+"</h"+hm[1].length+">";continue;}
    if(HR.test(ln)){flush();out+="<hr>";continue;}
    if(QT.test(ln)){
      flush();var q=[];
      while(i<lines.length&&QT.test(lines[i])) q.push(lines[i++].replace(QT,""));
      i--;
      out+="<blockquote>"+q.map(txt).join("<br>")+"</blockquote>";continue;
    }
    if(UL.test(ln)&&!HR.test(ln)){
      flush();var ul=[];
      while(i<lines.length&&UL.test(lines[i])&&!HR.test(lines[i])) ul.push(lines[i++].replace(UL,""));
      i--;
      out+="<ul>"+ul.map(function(l){return"<li>"+txt(l)+"</li>";}).join("")+"</ul>";continue;
    }
    if(OL.test(ln)){
      flush();var ol=[];
      while(i<lines.length&&OL.test(lines[i])) ol.push(lines[i++].replace(OL,""));
      i--;
      out+="<ol>"+ol.map(function(l){return"<li>"+txt(l)+"</li>";}).join("")+"</ol>";continue;
    }
    para.push(ln);
  }
  flush();
  return out;
}


function fetchAnnounce(){
  fetch("ANNOUNCE.md?_="+Date.now())
    .then(function(r){if(!r.ok)throw new Error("no file");return r.text();})
    .then(function(text){
      text=text.trim();if(!text)return;
      var hash=_annHash(text);
      if(localStorage.getItem("fia_ann_seen")===hash)return;
      document.getElementById("ann-body").innerHTML=md2html(text);
      var ol=document.getElementById("ann-overlay");
      ol.dataset.hash=hash;
      ol.style.display="flex";
    })
    .catch(function(){/* no ANNOUNCE.md → skip silently */});
}


function closeAnnounce(remember){
  var ol=document.getElementById("ann-overlay");
  if(remember){
    try{localStorage.setItem("fia_ann_seen",ol.dataset.hash||"");}catch(e){}
  }
  ol.style.display="none";
}
