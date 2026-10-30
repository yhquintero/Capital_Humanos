# Plantilla de cargos · fotografías 58622 y 58623

## Entregables

- **58622.xlsx / 58622.pdf**: primera fotografía, 22 cargos.
- **58623.xlsx / 58623.pdf**: segunda fotografía, 16 cargos.
- **Plantilla_Campismo.xlsx**: ambas páginas en hojas separadas, notas y fotos originales de referencia.
- **Plantilla_Campismo.pdf**: ambas tablas en dos páginas A4 horizontales.
- **Conversion_58622_58623.zip**: todos los documentos anteriores y sus notas.

Excel contiene texto y cifras editables, encabezados agrupados, bordes, colores y formato de impresión. Los PDF contienen texto seleccionable y vectorial: no son fotografías pegadas en una página.

**Importante:** consultar `NOTAS.md` y los comentarios de celda. Algunas grafías y escalas requieren confirmación. No se han sustituido los importes por los de otras plantillas ni recalculado salarios. Los marcadores `-0-` son texto. Las observaciones originales permanecen vacías. Las abreviaturas se han normalizado para legibilidad, no se certifican nombres oficiales de cargos. La rúbrica aparece solamente dentro de la fotografía original incluida como referencia en Excel; los documentos recreados no se presentan como firmados.

## Regenerar

Desde la raíz del repositorio, con Python 3.10+:

```bash
python -m venv .venv
.venv/bin/python -m pip install -r images/convertidos/requirements.txt
.venv/bin/python images/convertidos/generar.py
```

En Windows, usar `.venv\Scripts\python.exe`. La transcripción revisable está en `datos.json`. Los originales `images/58622.jpg` y `images/58623.jpg` no se modifican. Para comprobar la extracción de texto del PDF se puede utilizar PyMuPDF (dependencia opcional de verificación).
