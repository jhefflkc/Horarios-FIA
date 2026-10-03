# Respaldo de los temas anteriores

Copia de los cinco temas que tenía la web antes del rediseño «Vuelo»:

| Clase CSS      | Nombre en la web |
|----------------|------------------|
| *(ninguna)*    | Ámbar            |
| `stitch-dark`  | Grafito          |
| `light`        | Crema            |
| `stitch-light` | Glacial          |
| `google`       | Google           |

Estado exacto antes del rediseño: commit **`38f8233`** de `main`.

## Los temas no se han borrado

Siguen dentro de `assets/styles.css` y funcionan igual que antes; solo están
fuera de la rotación del botón de tema. Este directorio es una copia de
seguridad adicional por si `styles.css` se modifica más adelante.

| Archivo      | Qué contiene                                             |
|--------------|----------------------------------------------------------|
| `styles.css` | La hoja de estilos completa con los cinco temas          |
| `theme.js`   | `applyTheme()` / `toggleTheme()` y la onda del tema Google |
| `config.js`  | `THEME_ORDER`, etiquetas, iconos, colores de panel y PDF |

Estos archivos **no los carga la web**: están aquí solo como referencia.

## Cómo volver a activar uno o varios

En `assets/js/config.js`, añade sus clases a `THEME_ORDER`. Por ejemplo,
para tener otra vez los siete en rotación:

```js
var THEME_ORDER=["vuelo-claro","vuelo-oscuro","dark","stitch-dark","light","stitch-light","google"];
```

La lista completa de los antiguos está en `THEMES_LEGACY`, en el mismo archivo.
Sus etiquetas, iconos, colores de panel y del PDF siguen configurados.

Falta un paso: sus tipografías (Plus Jakarta Sans, Roboto y Roboto Mono) ya
no se descargan, para no cargar fuentes que nadie usa. En `index.html`, justo
debajo del `<link>` de Google Fonts, hay una segunda línea comentada con
ellas: quítale el comentario `<!-- … -->`. Sin ese paso el tema funciona, pero
con las fuentes de respaldo del sistema.

## Cómo volver al estado exacto anterior

```bash
git checkout 38f8233 -- assets/styles.css assets/js/theme.js assets/js/config.js index.html
```
