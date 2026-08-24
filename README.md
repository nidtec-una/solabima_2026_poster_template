# SOLABIMA 2026 Poster Template

Templates para la elaboración de posters para el SOLABIMA 2026

**XIV Congreso de la Sociedad Latinoamericana de Biología Matemática**
Facultad Politécnica, Universidad Nacional de Asunción (UNA), Paraguay
5–9 de octubre de 2026 · www.solabima2026.pol.una.py

## Contenido

- [`latex/`](latex/) — plantilla en LaTeX (`tikzposter`), A0 vertical,
  la mas trabajada de las tres (ver `latex/README.md`).
- [`pptx/`](pptx/) — plantilla PowerPoint (`.pptx`). Generada a partir
  de un diseño anterior de la version LaTeX; **todavia no incorpora**
  las iteraciones mas recientes (2 columnas con esquinas redondeadas,
  logos oficiales, etc.) — actualizar `pptx/build_pptx.py` si se
  retoma este formato.
- [`keynote/`](keynote/) — instrucciones para obtener la versión Keynote
  a partir del `.pptx`.

Los logos oficiales de SOLABIMA 2026, UNA y Facultad Politécnica ya
están en `latex/logos/`; los dos logos genéricos restantes son para
auspiciantes/colaboradores adicionales.

## Guía para autores

1. **Tamaño**: el tamaño de página es A0. La impresión del póster es
   responsabilidad de cada autor/a.
2. **Idioma**: el póster puede redactarse en español, inglés o
   portugués.
3. **Secciones**: las secciones incluidas (Motivación/Introducción,
   Objetivos, Metodología, Resultados, Referencias, Agradecimientos,
   etc.) son solo una guía orientativa — siéntase libre de adaptarlas,
   renombrarlas, reordenarlas o quitarlas según lo que necesite su
   trabajo.
4. **No mover la caja de título ni la de logos**: no cambie la
   posición de la caja de título ni la de la franja de logos
   institucionales al pie del póster.
5. **Logos**: los primeros tres logos (de izquierda a derecha) no
   deben modificarse ni reordenarse. Puede agregar logos adicionales
   a la derecha del tercero; al hacerlo, asegúrese de que queden
   centrados y de que quepan dentro de la caja.

## Cómo usar las plantillas

### LaTeX

La forma recomendada de trabajar con la plantilla LaTeX es
[Overleaf](https://www.overleaf.com):

1. Descargue la carpeta [`latex/`](latex/) completa (incluye el
   `.tex`, la carpeta `logos/` y el `README.md` de esa carpeta).
2. En Overleaf, cree un proyecto nuevo subiendo esa carpeta
   (**New Project → Upload Project**, como un `.zip` de la carpeta
   `latex/`).
3. En **Menu → Compiler**, seleccione el motor **pdfLaTeX** (motor
   estándar; la plantilla no requiere XeLaTeX ni LuaLaTeX).

También puede compilarla localmente con cualquier distribución TeX
Live / MacTeX reciente (ver instrucciones en `latex/README.md`).

> **Nota:** la comisión organizadora de SOLABIMA 2026 no brindará
> soporte para el uso de LaTeX. Ante dudas de LaTeX en sí (no del
> contenido del póster), consulte la documentación de Overleaf/LaTeX
> o abra un [issue](#issues) en este repositorio.

### PowerPoint

1. Descargue [`pptx/solabima2026_poster_template.pptx`](pptx/).
2. Ábralo con PowerPoint (o con Google Slides/LibreOffice Impress,
   con compatibilidad limitada) y edite el texto, las figuras/tablas
   y los logos de marcador de posición.
3. Exporte el resultado final a PDF para la impresión o el envío.

*(Ver la nota en [Contenido](#contenido): este formato todavía no
incorpora el diseño mas reciente de la versión LaTeX.)*

### Keynote

Keynote no importa `.pptx` con edición directa desde este repositorio
sin pasos previos; siga las instrucciones en
[`keynote/README.md`](keynote/README.md) para convertir
`pptx/solabima2026_poster_template.pptx` a un archivo `.key` nativo.

## Issues

¿Encontró un error en la plantilla o tiene una sugerencia? Abra un
issue en
[github.com/nidtec-una/solabima_2026_poster_template/issues](https://github.com/nidtec-una/solabima_2026_poster_template/issues).
