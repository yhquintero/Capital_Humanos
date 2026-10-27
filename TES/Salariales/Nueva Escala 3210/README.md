# Escala Salarial 3210 · interfaz interactiva

Herramienta Python para recalcular la escala salarial y preparar una propuesta imprimible a partir de `propuesta de plantilla por la resolucion 87.xls`.

## Funciones

- Slider en vivo entre **1× y 6×**, en pasos de 0.05.
- Aplicación del factor por cargo a cada salario extraído de la plantilla.
- Validación estricta por fila: `salario ajustado ≤ salario original × 6`.
- Mínimo del grupo I fijado en **$3,210.00 CUP** (Res. 14/2026).
- Grupos con propuesta de plantilla recalculados; grupos sin propuesta conservan la escala oficial Res. 15/2026 (44 h).
- Exportación del cálculo actualizado a **Excel `.xlsx`** (escala y detalle) y **PDF multipágina A4 horizontal**, listo para imprimir.
- Trazabilidad por hoja y fila de origen del libro `.xls`.

## Instalación y ejecución

Desde la raíz del repositorio:

```bash
python -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
python -m pip install -r "TES/Salariales/Nueva Escala 3210/requirements.txt"
streamlit run "TES/Salariales/Nueva Escala 3210/app_escala.py" --server.address 0.0.0.0
```

La interfaz se abre normalmente en `http://localhost:8501`. Al mover el slider, la escala y el detalle de cargos se recalculan de inmediato. Usa los botones de descarga para obtener la versión actual.

## Salidas disponibles

- `Escala_Salarial_3210_Nueva.xlsx` y `.pdf`: referencias generadas con `generar_escala.py`.
- Descargas interactivas: llevan el multiplicador elegido en el nombre de archivo y contienen la escala completa y el detalle actualizado de plantilla.

Para regenerar las salidas de referencia:

```bash
python "TES/Salariales/Nueva Escala 3210/generar_escala.py"
```

## Fuentes

- Resolución 14/2026, GOC-2026-393-O60: salario mínimo mensual de $3,210.00.
- Resolución 15/2026, GOC-2026-394-O60: escala oficial de 32 grupos, jornada de 44 horas.
- `Escala 3210_OK.xlsx`: escala histórica Res. 29/2020.
- `propuesta de plantilla por la resolucion 87.xls`: registros de cargos, grupos propuestos y salarios de origen.

## Notas de cálculo

El multiplicador de la interfaz se aplica a los salarios propuestos que aparecen en la plantilla de origen. Para cada registro se calcula un tope independiente de seis veces su propio salario original. Los cargos sin salario propuesto se muestran con el salario oficial de la Res. 15/2026; el mínimo del grupo I permanece anclado en $3,210.00. Por tanto, no se imputan salarios de plantilla a grupos que no aparecen en el libro fuente.

La fuente presenta dos discrepancias en los importes del grupo XII y del XV. El generador de referencia adopta el valor mayoritario y deja anotadas las ocurrencias; la interfaz conserva los registros individuales para facilitar la revisión. Validar los datos antes de una decisión administrativa.
