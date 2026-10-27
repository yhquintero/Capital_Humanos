# Nueva Escala Salarial — Salario Mínimo $3,210.00

Generador automático de la escala salarial con salario mínimo de **$3,210.00**,
que incorpora en una columna nueva los **salarios propuestos** del archivo
`propuesta de plantilla por la resolucion 87.xls` y exporta los resultados a
**Excel (.xlsx)** y **PDF**.

## ¿Por qué Python?

Se recomienda **Python 3** para este trabajo por estas razones:

1. **Lee el formato legacy `.xls`** (libro de 1997-2003) con `xlrd`, algo que
   muchas herramientas modernas ya no soportan; y escribe `.xlsx` con `openpyxl`.
2. **Exporta a PDF** directamente con `fpdf2`, sin necesidad de tener instalado
   Excel ni Acrobat.
3. **Auditabilidad**: el script recorre las 9 hojas de cargos y deja un
   registro fila por fila (hoja *Detalle Plantilla 87*) de dónde salió cada valor.
4. **Reutilizable**: si la plantilla o la escala oficial cambian, se vuelve a
   ejecutar un solo comando y ambos archivos se regeneran igual.
5. **Es gratuito y funciona igual en Windows, Linux y macOS.**

Alternativas consideradas: un libro Excel con macros VBA obliga a usar Microsoft
Office y sus macros suelen ser bloqueadas; R o Java pueden hacer lo mismo pero
con más código y menos librerías para estos formatos de oficina.

## Requisitos

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

```bash
python generar_escala.py
```

El script lee:

- `Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls`
- `TES/Salariales/Escala 3210_OK.xlsx` (escala anterior, Resolución 29/2020)
- Escala oficial de la **Resolución 15/2026** (GOC-2026-O60, pp. 1839-1840),
  incorporada como constante citada en el código.

## Salidas

| Archivo | Contenido |
|---|---|
| `Escala_Salarial_3210_Nueva.xlsx` | 4 hojas (ver abajo) |
| `Escala_Salarial_3210_Nueva.pdf` | Tabla ejecutiva en A4 horizontal |

Hojas del Excel:

1. **Escala Salarial 3210**: tabla principal de 32 grupos (I a XXXII) con la
   escala anterior, coeficiente de correlación, escala oficial Res. 15/2026,
   diferencia e incremento %, y en **verde la columna nueva con el salario
   propuesto de la plantilla del Decreto 87** y su diferencia contra la escala oficial.
2. **Escala por correlacion**: cálculo automático con **fórmulas vivas de Excel**
   (`salario_n = correlación_n × salario_n-1`), partiendo de $3,210.00.
3. **Detalle Plantilla 87**: los 183 cargos extraídos del `.xls`, con hoja,
   fila, nombre del cargo, grupo y salario propuesto (auditoría).
4. **Metodologia y fuentes**: normas citadas, metodología y anomalías detectadas.

## Hallazgos importantes

- La plantilla del Decreto 87 solo propone salarios para los **grupos IV a XVIII**;
  los grupos I-III y XIX-XXXII aparecen con "—".
- La propuesta de la plantilla **no coincide** con la escala oficial
  Res. 15/2026: en los grupos IV-IX la plantilla paga por debajo de la escala
  oficial y del grupo X en adelante por encima (el Excel muestra la diferencia).
- Anomalías en el `.xls` fuente (resueltas por mayoría, documentadas en las salidas):
  - Grupo **XII**: una fila con $5,570.00 frente a 15 con $5,770.00 → se adopta $5,770.00.
  - Grupo **XV**: una fila con $7,777.00 frente a 5 con $7,201.00 → se adopta $7,201.00.

## Normas citadas

- Resolución **14/2026** (MTSS, GOC-2026-393-O60): salario mínimo $3,210.00.
- Resolución **15/2026** (MTSS, GOC-2026-394-O60): escala salarial del sector
  presupuestado, 32 grupos, regímenes de 44 y 40 horas.
- Decreto **87/2023**: organización de la plantilla de cargos de la empresa.
