# Plantilla LaTeX - SOLABIMA 2026

Poster academico A0 vertical (841 x 1189 mm) construido con la clase
`tikzposter`, en espanol, con la paleta de colores provisional de
SOLABIMA 2026 (ver nota de colores en `solabima2026_poster.tex`).

## Compilacion

Requiere una distribucion TeX Live / MacTeX con `tikzposter`,
`beamerposter`... en realidad solo se usa `tikzposter`, `qrcode` y
paquetes estandar (`babel`, `booktabs`, `graphicx`).

```bash
cd latex
pdflatex solabima2026_poster.tex
pdflatex solabima2026_poster.tex   # segunda pasada (referencias/QR)
```

Esto genera `solabima2026_poster.pdf`.

## Estructura

- `solabima2026_poster.tex` - archivo principal, editable.
- `logos/` - logos de marcador de posicion (UNA, Facultad Politecnica,
  NIDTEC, ARASY, SOLABIMA 2026). **Reemplacelos por los logos
  oficiales** antes de imprimir; mantenga los mismos nombres de
  archivo o actualice las rutas en `\titlegraphic`.
- `logos/make_placeholders.py` - script que genero los logos de
  marcador de posicion (no es necesario ejecutarlo salvo que quiera
  regenerarlos).

## Personalizacion

- **Tamano de poster**: no hay especificacion oficial publicada al
  momento de crear esta plantilla. Si el comite organizador define un
  tamano distinto (por ejemplo A0 horizontal), cambie las opciones de
  `\documentclass` (`a0paper`, `portrait`/`landscape`).
- **Colores**: definidos al inicio del archivo
  (`solabimaBlue`, `solabimaTeal`, `solabimaGold`, `solabimaGrey`).
  Actualicelos si se publica una guia de marca oficial.
- **Secciones**: Introduccion, Objetivos, Metodologia, Resultados,
  Discusion, Conclusiones, Referencias y agradecimientos. Agregue o
  quite bloques segun el contenido de su trabajo.
