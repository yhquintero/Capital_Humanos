<div align="center">

<img src="docs/capital_humano_banner.svg" width="100%" alt="Capital Humano — escala salarial interactiva, mínimo $3,210 CUP, ajuste en vivo 1 a 6 veces y exportación PDF y Excel" />

### Escala salarial interactiva · propuesta de plantilla · exportación PDF y Excel

Herramienta para revisar y ajustar la propuesta salarial de la plantilla, partiendo del salario mínimo de **$3,210.00 CUP** y respetando un máximo de **6× el salario original** de cada cargo.

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](TES/Salariales/Nueva%20Escala%203210/README.md)
[![Excel](https://img.shields.io/badge/Exportaci%C3%B3n-Excel%20%7C%20PDF-16866B?style=flat-square)](TES/Salariales/Nueva%20Escala%203210/)
[![Estado](https://img.shields.io/badge/estado-en%20desarrollo-E6B85C?style=flat-square)](#-alcance-y-validación)

**[Abrir la guía de uso](TES/Salariales/Nueva%20Escala%203210/README.md) · [Ver el Excel de ejemplo](TES/Salariales/Nueva%20Escala%203210/Escala_Salarial_3210_Nueva.xlsx) · [Ver el PDF de ejemplo](TES/Salariales/Nueva%20Escala%203210/Escala_Salarial_3210_Nueva.pdf)**

</div>

---

## 🎯 Qué resuelve

- Lee los salarios y grupos de la plantilla heredada `propuesta de plantilla por la resolucion 87.xls`.
- Permite cambiar en **tiempo real** un multiplicador común de **1.00× a 6.00×**.
- Calcula y muestra para cada cargo el salario original, el multiplicador, el tope y el importe ajustado.
- Aplica un control explícito: **el salario ajustado no puede superar 6× el salario base del cargo en la plantilla fuente**.
- Mantiene anclado el **grupo I en $3,210.00**; los grupos sin propuesta de cargos conservan la escala oficial de 44 horas.
- Exporta la escala y el detalle de todos los cargos a **Excel (.xlsx)** y a **PDF A4 horizontal listo para imprimir**.

## 🚀 Inicio rápido

Requiere Python 3.10 o posterior. Desde la raíz del repositorio:

```bash
python -m venv .venv
# Linux / macOS
source .venv/bin/activate
# Windows: .venv\Scripts\activate
python -m pip install -r "TES/Salariales/Nueva Escala 3210/requirements.txt"
streamlit run "TES/Salariales/Nueva Escala 3210/app_escala.py" --server.address 0.0.0.0
```

Streamlit mostrará una dirección local, normalmente `http://localhost:8501`. Abre el enlace, mueve el control **Multiplicador salarial** y descarga la versión actualizada con los botones de Excel o PDF. En el diálogo de impresión también puedes elegir **Guardar como PDF**.

Para regenerar los archivos de referencia originales desde la fuente:

```bash
python "TES/Salariales/Nueva Escala 3210/generar_escala.py"
```

## 🧭 Cómo usar el ajuste

1. Mueve el multiplicador desde **1.00×** hasta **6.00×**; los importes de la escala y del detalle se actualizan inmediatamente.
2. Comprueba las columnas **Salario original**, **Tope máximo 6×** y **Salario ajustado**.
3. Revisa el detalle de cargos y las observaciones de datos en la interfaz.
4. Descarga el libro Excel o el informe PDF. Ambos archivos identifican el multiplicador usado y mantienen trazabilidad del origen.

El control se aplica a cada salario propuesto de la plantilla; el tope se verifica por fila, no solo sobre un promedio o el total. Los grupos que no aparecen en la plantilla mantienen el valor oficial vigente, para no inventar salarios fuente inexistentes.

## 📚 Fuentes y metodología

| Dato | Fuente |
|---|---|
| Salario mínimo mensual de $3,210.00 | Resolución 14/2026, GOC-2026-393-O60 |
| Escala oficial de 32 grupos, jornada de 44 horas | Resolución 15/2026, GOC-2026-394-O60 |
| Salario histórico | `TES/Salariales/Escala 3210_OK.xlsx` (Res. 29/2020) |
| Grupo, cargo y salario propuesto | `Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls` |

La plantilla incluye 183 registros distribuidos entre los grupos IV–XVIII. Cuando una misma categoría de grupo contiene importes distintos, el generador identifica y documenta las discrepancias; no modifica el archivo fuente.

## 📦 Archivos

```text
TES/Salariales/Nueva Escala 3210/
├── app_escala.py                     # interfaz web interactiva y exportadores
├── generar_escala.py                 # extracción y generador de referencia
├── requirements.txt                  # dependencias
├── README.md                         # guía detallada
├── Escala_Salarial_3210_Nueva.xlsx   # ejemplo Excel
└── Escala_Salarial_3210_Nueva.pdf    # ejemplo PDF
```

## 🧰 Tecnologías

- **Python** — extracción, validaciones y cálculos reproducibles.
- **Streamlit** — interfaz interactiva con actualización inmediata.
- **xlrd** — lectura del libro antiguo `.xls`.
- **openpyxl** — generación de `.xlsx` con hojas, formato, filtros y configuración de impresión.
- **fpdf2** — informes PDF multipágina listos para imprimir.

## ✅ Alcance y validación

La herramienta facilita el análisis y la preparación de una propuesta; no sustituye la revisión de Recursos Humanos, Finanzas ni la validación jurídica que corresponda. Los importes calculados deben aprobarse por las autoridades competentes antes de su aplicación oficial.

---

<div align="center"><sub>Capital Humano · Escala salarial y plantilla · Cuba</sub></div>
