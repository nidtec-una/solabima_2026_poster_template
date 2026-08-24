# Plantilla LaTeX - SOLABIMA 2026

Poster academico A0 vertical (841 x 1189 mm), dos columnas, construido
con la clase `tikzposter`. Secciones en caja (barra de titulo + cuerpo
con borde, esquinas superior-izquierda e inferior-derecha redondeadas),
fondo blanco, acentos con la paleta provisional de SOLABIMA 2026, y una
franja de logos institucionales anclada cerca del pie de la pagina.

## Compilacion

Requiere una distribucion TeX Live / MacTeX. Ademas de `tikzposter`,
usa paquetes estandar: `babel`, `booktabs`, `graphicx`, `amsmath`, y la
libreria `calc` de TikZ.

```bash
cd latex
pdflatex solabima2026_poster.tex
pdflatex solabima2026_poster.tex   # segunda pasada (referencias cruzadas)
```

Esto genera `solabima2026_poster.pdf`.

## Estructura

- `solabima2026_poster.tex` - archivo principal, editable.
- `logos/` - logos institucionales:
  - `logo_solabima.png`, `logo_una.jpg`, `logo_fpuna.pdf` son los
    logos oficiales reales (SOLABIMA 2026, UNA, Facultad Politecnica).
  - `logo_custom1.png`, `logo_custom2.png` son placeholders genericos
    para logos de auspiciantes/colaboradores adicionales (por ejemplo
    ARASY); reemplacelos por los logos reales cuando esten
    disponibles, o quite esas dos lineas de `\includegraphics` en la
    franja de pie de poster si no las necesita.
  - `fpuna_logo.svg` - fuente vectorial original del logo de FPUNA
    (se convirtio a `logo_fpuna.pdf` porque `pdflatex` no puede
    incluir SVG directamente).
  - `logos/make_placeholders.py` - regenera los placeholders
    genericos (`logo_custom1.png`, `logo_custom2.png`) si hace falta.
    No toca los logos oficiales.

## Personalizacion

- **Tamano de poster**: no hay especificacion oficial publicada al
  momento de crear esta plantilla. Si el comite organizador define un
  tamano distinto (por ejemplo A0 horizontal), cambie las opciones de
  `\documentclass` (`a0paper`, `portrait`/`landscape`).
- **Colores**: definidos cerca del inicio del archivo
  (`solabimaBlue`, `solabimaTeal`, `solabimaGold`, `solabimaGrey`).
  Actualicelos si se publica una guia de marca oficial; tambien
  actualice `\definecolorpalette{solabimaPalette}{...}` para que
  coincida.
- **Tamano de fuente**: la plantilla redefine `\tiny`...`\Huge`
  globalmente (~25% mas grandes que el valor por defecto de
  tikzposter) para que se lea bien a distancia y las cajas ocupen
  buena parte de la pagina. Ajuste esos `\renewcommand` si necesita
  otro tamano base.
- **Secciones**: Motivacion/Introduccion, Objetivos, Metodologia,
  Resultados, Referencias, Agradecimientos, repartidas en 2 columnas
  (`\column{0.5}` cada una -- si cambia el numero de columnas, las
  fracciones deben sumar 1). Agregue o quite bloques `\block{...}{...}`
  segun el contenido de su trabajo.
- **Franja de logos**: esta anclada a una distancia fija del borde
  inferior de la pagina (variable `\TP@blocktop` seteada justo antes
  del segundo `\begin{columns}`), no al final natural del contenido de
  las columnas. Si el contenido crece demasiado puede superponerse con
  la franja; ajuste ese valor o recorte contenido si eso ocurre.
- **Pies de figura/tabla**: sin numeracion automatica y en negrita
  (ver la redefinicion de `tikzfigure`). El pie de figura va abajo;
  el de tabla, por convencion, se escribe a mano arriba de la tabla
  (no usa el entorno `tikzfigure`, ver el bloque de Resultados).
