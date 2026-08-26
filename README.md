# Horarios FIA

Generador de horarios para la Facultad de Ingeniería Ambiental de la UNI.
Es una web estática: no hay servidor ni base de datos, se publica sola en
GitHub Pages cada vez que haces push a `main`.

## Estructura

```
index.html              solo el marcado de la página
assets/
  styles.css            todos los estilos (los temas van al principio)
  data.js               GENERADO — los horarios convertidos a JavaScript
  js/
    config.js           ajustes: facultades, paleta, URL de calificaciones
    state.js            estado en memoria (cursos cargados, selección)
    courses.js          secciones, docentes, cruces de horario
    faculty.js          facultad activa, periodo, subir un Excel
    render.js           listado de cursos, etiquetas y tabla
    modals.js           ventanas modales
    ratings.js          calificaciones de docentes y Score
    announce.js         anuncios (Markdown → HTML)
    exports.js          exportar a PDF y a calendario .ics
    theme.js            cambio de tema
    main.js             arranque, navegación móvil, avisos
data/                   los .xlsx de horarios
docs/                   planes de estudio (referencia, no los usa la web)
build_data.py           lee data/*.xlsx y escribe assets/data.js
ANNOUNCE.md             el anuncio que sale al abrir la web
```

Los `.js` se cargan como scripts normales, en el orden que aparece al final
de `index.html`. No hay compilación ni dependencias que instalar.

## Actualizar los horarios

1. Nombra el Excel `{SIGLA}{AÑO}-{PERIODO}.xlsx` — por ejemplo `FIA2026-2.xlsx`
2. Súbelo a la carpeta `data/`
3. Borra el del periodo anterior si ya no lo necesitas

Al hacer push, GitHub Actions ejecuta `build_data.py` y regenera
`assets/data.js`. **Si subes dos periodos de la misma facultad se usa el más
reciente** y el otro se ignora (queda anotado en el log del despliegue).

Siglas reconocidas: `FIA`, `FIEECS`, `FIGMM`, `FIQT`, `FIIS`, `FIEE`, `FIPP`,
`FIM`, `FIC`, `FC`, `FAUA`. Para añadir otra, agrégala en `FACULTY_MAP`
(`build_data.py`) y en `FACULTY_MAP_JS` (`assets/js/config.js`).

### Columnas del Excel

Obligatorias: `COD` y `CURSO`. Opcionales: `SECC`, `TIPO`, `AULA`/`SALON`,
`DOCENTE`, `CICLO`. Para el horario, o bien una columna `HORARIO` con el
formato `MA 10-12`, o bien `DIA` + `H INI` + `H FIN`.

El docente se lee **de cada fila**, así que si la teoría y el laboratorio los
dictan personas distintas, ambas aparecen con su rol. Si dejas la celda vacía
en las filas de práctica, ese docente no se mostrará.

## Actualizar el anuncio

Edita `ANNOUNCE.md` y súbelo. Se muestra al abrir la web y vuelve a aparecer
para todos en cuanto cambie el contenido; quien pulse «Entendido» no lo verá
otra vez hasta la próxima edición.

Acepta Markdown normal: títulos, **negrita**, *cursiva*, listas, `código`,
enlaces, citas con `>`, `---` como separador, emojis y símbolos. Para dejar un
asterisco o guion bajo literal, escápalo: `\*así\*`.

## Seguridad

El repositorio es público y la web es estática, así que **todo lo que hay en
`assets/js/` lo puede leer cualquiera** — incluida la URL del servicio de
calificaciones. Eso no es un descuido: el navegador de cada visitante tiene
que llamar a ese endpoint, así que la URL no puede mantenerse en secreto.
La seguridad tiene que venir de *qué permite hacer el endpoint*, no de
esconderlo.

En el repositorio **no debe haber nunca** claves de servicio, `client_secret`,
tokens ni archivos `.json` de credenciales de Google. Si alguna vez se sube
uno, no basta con borrarlo: queda en el historial de git y hay que revocarlo
en Google.

Al tocar `assets/js/ratings.js`, ten presente que lo que responde el servicio
de calificaciones es **dato ajeno**: pásalo siempre por `_avgOf()` /
`_countOf()` antes de meterlo en `innerHTML`. Si se concatena en crudo,
cualquiera que consiga escribir en la hoja de cálculo puede ejecutar
JavaScript en el navegador de todos los visitantes.

Recomendado en el Apps Script (esto se configura en Google, no aquí):

- Que valide `action`, limite `puntuacion` al rango 1–5 y descarte cualquier
  otro parámetro
- Que tenga permisos solo sobre la hoja de cálculo de calificaciones
- Idealmente, que viva en una cuenta de Google aparte de la personal

El botón «Ya calificaste» solo mira `localStorage`, así que **no impide votar
varias veces**: quien quiera puede llamar al endpoint directamente. Las
calificaciones son orientativas, no un recuento fiable.

## Probarlo en tu máquina

Hace falta servirlo por HTTP (abrir el archivo directamente no carga los
scripts ni el anuncio):

```bash
python -m http.server 8000
# luego abre http://localhost:8000
```

Para regenerar los datos en local:

```bash
pip install openpyxl
python build_data.py
```
