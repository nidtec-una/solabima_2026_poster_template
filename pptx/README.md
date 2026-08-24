# Plantilla PowerPoint - SOLABIMA 2026

`solabima2026_poster_template.pptx`: poster A0 vertical (841 x 1189 mm)
con el mismo contenido, estructura y paleta de colores que la
plantilla LaTeX en `../latex`.

Abra el archivo directamente en PowerPoint (o en Keynote, ver
`../keynote/README.md`) y edite el titulo, autores, texto de cada
bloque y reemplace las figuras/tablas de marcador de posicion.
Tambien reemplace los logos de marcador de posicion por los oficiales
cuando esten disponibles.

## Regenerar el archivo

El `.pptx` se genera con el script `build_pptx.py` (usa
`python-pptx` y `qrcode`). Si prefiere ajustar el layout por codigo en
lugar de a mano en PowerPoint:

```bash
pip3 install python-pptx "qrcode[pil]"
cd pptx
python3 build_pptx.py
```

Esto sobrescribe `solabima2026_poster_template.pptx`.
