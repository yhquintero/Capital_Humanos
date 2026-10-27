# -*- coding: utf-8 -*-
"""
=============================================================================
 GENERADOR DE LA NUEVA ESCALA SALARIAL — SALARIO MÍNIMO $3,210.00
-----------------------------------------------------------------------------
 Entradas:
   1. "Nueva Estructura de Plantilla/Organigrama/
       propuesta de plantilla por la resolucion 87.xls"  (formato .xls viejo)
      -> se extrae, por GRUPO DE ESCALA PROPUESTO (G.E.A / G.E.P), el
         "Salario Propuesto" de TODAS las hojas de plantilla.
   2. "TES/Salariales/Escala 3210_OK.xlsx"
      -> salarios de la escala anterior (Resolución 29/2020).
   3. Escala oficial de la Resolución 15/2026 (GOC-2026-O60, pp. 1839-1840),
      tabla de 44 horas semanales, incorporada como constante citada.

 Salidas (en esta misma carpeta):
   - Escala_Salarial_3210_Nueva.xlsx   (4 hojas, con fórmulas y formato)
   - Escala_Salarial_3210_Nueva.pdf    (tabla ejecutiva, A4 horizontal)

 Requisitos:  pip install xlrd openpyxl fpdf2
 (entorno virtual recomendado)
=============================================================================
"""

from __future__ import annotations

import datetime as dt
import math
from collections import Counter, defaultdict
from pathlib import Path

import xlrd
from fpdf import FPDF
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 0. RUTAS Y CONSTANTES
# ---------------------------------------------------------------------------
CARPETA = Path(__file__).resolve().parent
RAIZ_REPO = CARPETA.parents[2]
RUTA_PLANTILLA = RAIZ_REPO / (
    "Nueva Estructura de Plantilla"
    "/Organigrama/propuesta de plantilla por la resolucion 87.xls"
)
RUTA_ESCALA_BASE = CARPETA.parent / "Escala 3210_OK.xlsx"
SALIDA_XLSX = CARPETA / "Escala_Salarial_3210_Nueva.xlsx"
SALIDA_PDF = CARPETA / "Escala_Salarial_3210_Nueva.pdf"

GRUPOS_ROMANOS = [
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX",
    "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXVI", "XXVII", "XXVIII",
    "XXIX", "XXX", "XXXI", "XXXII",
]
ROMANO_A_ENTERO = {g: i + 1 for i, g in enumerate(GRUPOS_ROMANOS)}

# Escala oficial Resolución 15/2026 (GOC-2026-O60), régimen de 44 horas.
# Fuente: tabla "PRIMERO" de la Resolución 15/2026, Gaceta Oficial No. 60
# Ordinaria de 17/07/2026, páginas 1839-1840.
ESCALA_OFICIAL_RES15 = {
    "I": 3210, "II": 3365, "III": 3515, "IV": 3700, "V": 3885,
    "VI": 4065, "VII": 4295, "VIII": 4525, "IX": 4755, "X": 4985,
    "XI": 5210, "XII": 5520, "XIII": 5825, "XIV": 6130, "XV": 6435,
    "XVI": 6740, "XVII": 7045, "XVIII": 7350, "XIX": 7735, "XX": 8115,
    "XXI": 8500, "XXII": 8880, "XXIII": 9265, "XXIV": 9645, "XXV": 10105,
    "XXVI": 10640, "XXVII": 11175, "XXVIII": 11710, "XXIX": 12245,
    "XXX": 13010, "XXXI": 13770, "XXXII": 14535,
}

# Respaldo por si no se pudiera leer el .xlsx de la escala anterior.
RESPALDO_RES29 = {
    "I": 2100, "II": 2200, "III": 2300, "IV": 2420, "V": 2540,
    "VI": 2660, "VII": 2810, "VIII": 2960, "IX": 3110, "X": 3260,
    "XI": 3410, "XII": 3610, "XIII": 3810, "XIV": 4010, "XV": 4210,
    "XVI": 4410, "XVII": 4610, "XVIII": 4810, "XIX": 5060, "XX": 5310,
    "XXI": 5560, "XXII": 5810, "XXIII": 6060, "XXIV": 6310, "XXV": 6610,
    "XXVI": 6960, "XXVII": 7310, "XXVIII": 7660, "XXIX": 8010,
    "XXX": 8510, "XXXI": 9010, "XXXII": 9510,
}

FORMATO_MONEDA = '"$"#,##0.00'
FORMATO_MONEDA_ENTERO = '"$"#,##0'
FORMATO_PORCIENTO = '0.00%'
FORMATO_FACTOR = '0.0000'


# ---------------------------------------------------------------------------
# 1. ESCALA ANTERIOR (RESOLUCIÓN 29/2020)
# ---------------------------------------------------------------------------
def cargar_salarios_res29() -> dict[str, float]:
    """Lee el salario de la Res. 29 desde Escala 3210_OK.xlsx (columna B)."""
    try:
        wb = load_workbook(RUTA_ESCALA_BASE, data_only=True)
        ws = wb.active
        salida = {}
        for fila in ws.iter_rows(min_row=2, max_row=33, max_col=2):
            grupo, salario = fila[0].value, fila[1].value
            if grupo in ROMANO_A_ENTERO and isinstance(salario, (int, float)):
                salida[grupo] = float(salario)
        if len(salida) == 32:
            return salida
        print("  (aviso) lectura incompleta del .xlsx base; se usa respaldo)")
    except FileNotFoundError:
        print("  (aviso) no se encontró Escala 3210_OK.xlsx; se usa respaldo)")
    return dict(RESPALDO_RES29)


# ---------------------------------------------------------------------------
# 2. EXTRACCIÓN DE LA PROPUESTA DE PLANTILLA (resolución 87.xls)
#    La posición de las columnas varía por hoja: se localiza el encabezado
#    "G.E.A" / "G.E.P" y el salario propuesto está en la columna siguiente.
# ---------------------------------------------------------------------------
def _localizar_columna_ge(hoja: xlrd.sheet.Sheet):
    """Devuelve (fila_encabezado, columna_grupo_propuesto)."""
    for r in range(min(25, hoja.nrows)):
        for c in range(hoja.ncols):
            texto = str(hoja.cell_value(r, c)).upper().replace(" ", "")
            if texto.startswith("G.E."):
                return r, c
    return None, None


def _mejor_cargo_de_la_fila(hoja, fila: int, col_grupo: int) -> str:
    """Selecciona el texto más largo de la fila como descripción del cargo."""
    candidatos = []
    for c in range(hoja.ncols):
        v = hoja.cell_value(fila, c)
        if isinstance(v, str) and len(v.strip()) >= 5 and c != col_grupo:
            candidatos.append(v.strip())
    return max(candidatos, key=len) if candidatos else ""


def extraer_propuesta_plantilla():
    """
    Recorre todas las hojas de cargos y devuelve:
      mapa_canónico : {grupo: salario_propuesto}  (valor más frecuente)
      contadores    : {grupo: Counter({salario: n})}
      registros     : lista auditable (hoja, fila, cargo, grupo, salario)
      anomalias     : grupos con más de un salario distinto
    """
    libro = xlrd.open_workbook(RUTA_PLANTILLA)
    contadores: dict[str, Counter] = defaultdict(Counter)
    registros = []

    for hoja in libro.sheets():
        if hoja.name == "Hoja1":
            continue  # resumen por centro de costos, no es plantilla de cargos
        fila_enc, col_grupo = _localizar_columna_ge(hoja)
        if col_grupo is None:
            continue
        col_salario = col_grupo + 1
        for r in range(fila_enc + 1, hoja.nrows):
            grupo = hoja.cell_value(r, col_grupo)
            if not (isinstance(grupo, str) and grupo.strip() in ROMANO_A_ENTERO):
                continue
            try:
                salario = float(hoja.cell_value(r, col_salario))
            except (TypeError, ValueError):
                continue
            grupo = grupo.strip()
            contadores[grupo][salario] += 1
            registros.append({
                "hoja": hoja.name, "fila": r + 1,
                "cargo": _mejor_cargo_de_la_fila(hoja, r, col_grupo),
                "grupo": grupo, "salario": salario,
            })

    mapa_canónico = {g: cont.most_common(1)[0][0] for g, cont in contadores.items()}
    anomalias = []
    for g in sorted(contadores, key=lambda x: ROMANO_A_ENTERO[x]):
        if len(contadores[g]) > 1:
            valores = ", ".join(
                f"${s:,.2f} ({n} ocurrencias)" for s, n in contadores[g].most_common()
            )
            anomalias.append(
                f"Grupo {g}: valores encontrados -> {valores}. "
                f"Se adopta ${mapa_canónico[g]:,.2f} por ser el mayoritario."
            )
    return mapa_canónico, contadores, registros, anomalias


# ---------------------------------------------------------------------------
# 3. CONSTRUCCIÓN DE LA TABLA FINAL
# ---------------------------------------------------------------------------
def construir_filas(salarios_res29, propuesta):
    filas = []
    for i, grupo in enumerate(GRUPOS_ROMANOS):
        anterior = salarios_res29[grupo]
        oficial = ESCALA_OFICIAL_RES15[grupo]
        factor = (anterior / salarios_res29[GRUPOS_ROMANOS[i - 1]]
                  if i > 0 else None)
        diferencia = oficial - anterior
        incremento_pct = diferencia / anterior
        salario_plantilla = propuesta.get(grupo)
        dif_plantilla_oficial = (
            salario_plantilla - oficial if salario_plantilla is not None else None
        )
        if salario_plantilla is None:
            observacion = "Sin propuesta en la plantilla del Decreto 87"
        else:
            observacion = ""
        filas.append({
            "grupo": grupo,
            "anterior": anterior,
            "factor": factor,
            "oficial": oficial,
            "diferencia": diferencia,
            "incremento_pct": incremento_pct,
            "plantilla": salario_plantilla,
            "dif_plantilla": dif_plantilla_oficial,
            "observacion": observacion,
        })
    return filas


# ---------------------------------------------------------------------------
# 4. EXCEL
# ---------------------------------------------------------------------------
AZUL = "1F4E79"
AZUL_CLARO = "D9E1F2"
VERDE_CLARO = "E2EFDA"
VERDE = "375623"
GRIS = "F2F2F2"
AMARILLO = "FFF2CC"
BLANCO = "FFFFFF"

borde_fino = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def _celda(ws, coord, valor, *, negrita=False, color=BLANCO, relleno=None,
           formato=None, alineacion="center", tamano=10):
    c = ws[coord]
    c.value = valor
    c.font = Font(name="Calibri", bold=negrita, size=tamano,
                  color=color if relleno else "000000")
    c.alignment = Alignment(horizontal=alineacion, vertical="center", wrap_text=True)
    c.border = borde_fino
    if relleno:
        c.fill = PatternFill("solid", fgColor=relleno)
    if formato:
        c.number_format = formato


def escribir_excel(filas, contadores, registros, anomalias):
    wb = Workbook()

    # ---- Hoja 1: escala principal ----------------------------------------
    ws = wb.active
    ws.title = "Escala Salarial 3210"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:I1")
    _celda(ws, "A1",
           "NUEVA ESCALA SALARIAL — SALARIO MÍNIMO $3,210.00 (Res. 14/2026 y 15/2026, GOC-2026-O60)",
           negrita=True, color=BLANCO, relleno=AZUL, tamano=13,
           alineacion="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:I2")
    hoy = dt.date.today().strftime("%d/%m/%Y")
    _celda(ws, "A2",
           f"Generado el {hoy}. La columna G se alimenta del archivo "
           f"“propuesta de plantilla por la resolucion 87.xls”. Montos en pesos cubanos (CUP).",
           negrita=False, relleno=GRIS, alineacion="left", tamano=9)

    encabezados = [
        "Grupo de complejidad",
        "Salario escala anterior\nRes. 29/2020",
        "Factor de correlación\nRes. 29/2020",
        "Nueva escala oficial\nRes. 15/2026 (44 h)",
        "Diferencia\n(oficial − anterior)",
        "Incremento %",
        "Salario propuesto\nPlantilla Decreto 87 (.xls)",
        "Diferencia\n(plantilla − oficial)",
        "Observaciones",
    ]
    fila_enc = 3
    for j, titulo in enumerate(encabezados, start=1):
        _celda(ws, f"{get_column_letter(j)}{fila_enc}", titulo,
               negrita=True, color=BLANCO,
               relleno=VERDE if j == 7 else AZUL, tamano=10)
    ws.row_dimensions[fila_enc].height = 42

    for i, f in enumerate(filas):
        r = fila_enc + 1 + i
        banda = GRIS if i % 2 else BLANCO
        _celda(ws, f"A{r}", f["grupo"], negrita=True, relleno=AZUL_CLARO)
        _celda(ws, f"B{r}", f["anterior"], relleno=banda, formato=FORMATO_MONEDA)
        if f["factor"] is not None:
            _celda(ws, f"C{r}", f["factor"], relleno=banda, formato=FORMATO_FACTOR)
        else:
            _celda(ws, f"C{r}", "—", relleno=banda)
        _celda(ws, f"D{r}", f["oficial"], negrita=True, relleno=banda,
               formato=FORMATO_MONEDA_ENTERO)
        _celda(ws, f"E{r}", f["diferencia"], relleno=banda, formato=FORMATO_MONEDA)
        _celda(ws, f"F{r}", f["incremento_pct"], relleno=banda,
               formato=FORMATO_PORCIENTO)
        if f["plantilla"] is not None:
            _celda(ws, f"G{r}", f["plantilla"], negrita=True,
                   relleno=VERDE_CLARO, formato=FORMATO_MONEDA_ENTERO)
            _celda(ws, f"H{r}", f["dif_plantilla"], relleno=VERDE_CLARO,
                   formato=FORMATO_MONEDA)
        else:
            _celda(ws, f"G{r}", "—", relleno=VERDE_CLARO)
            _celda(ws, f"H{r}", "—", relleno=VERDE_CLARO)
        _celda(ws, f"I{r}", f["observacion"], relleno=banda, alineacion="left",
               tamano=9)

    # notas de anomalías dentro de la hoja
    fila_nota = fila_enc + 1 + len(filas) + 1
    ws.merge_cells(start_row=fila_nota, start_column=1,
                   end_row=fila_nota, end_column=9)
    _celda(ws, f"A{fila_nota}",
           "NOTAS: " + " / ".join(anomalias) if anomalias else "NOTAS: sin anomalías.",
           relleno=AMARILLO, alineacion="left", tamano=9)
    ws.row_dimensions[fila_nota].height = 30

    anchos = [14, 18, 16, 18, 18, 12, 20, 18, 34]
    for j, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A4"

    # ---- Hoja 2: cálculo por correlación (fórmulas vivas) ----------------
    ws2 = wb.create_sheet("Escala por correlacion")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:E1")
    _celda(ws2, "A1",
           "ESCALA CALCULADA POR FACTORES DE CORRELACIÓN (metodología del archivo "
           "Escala 3210_OK.xlsx) — con fórmulas de Excel",
           negrita=True, color=BLANCO, relleno=AZUL, alineacion="left", tamano=11)
    for j, t in enumerate([
        "Grupo de complejidad", "Salario Res. 29/2020",
        "Nueva escala (mínimo $3,210.00)", "Correlación Res. 29",
        "Diferencia"], start=1):
        _celda(ws2, f"{get_column_letter(j)}2", t, negrita=True,
               color=BLANCO, relleno=AZUL)
    ws2.row_dimensions[2].height = 32
    for i, f in enumerate(filas):
        r = 3 + i
        _celda(ws2, f"A{r}", f["grupo"], negrita=True, relleno=AZUL_CLARO)
        _celda(ws2, f"B{r}", f["anterior"], formato=FORMATO_MONEDA)
        if i == 0:
            _celda(ws2, f"C{r}", 3210, negrita=True, formato=FORMATO_MONEDA_ENTERO)
            _celda(ws2, f"D{r}", "—")
            _celda(ws2, f"E{r}", "—")
        else:
            _celda(ws2, f"D{r}", f"=B{r}/B{r-1}", formato=FORMATO_FACTOR)
            _celda(ws2, f"C{r}", f"=D{r}*C{r-1}", negrita=True,
                   formato=FORMATO_MONEDA)
            _celda(ws2, f"E{r}", f"=C{r}-C{r-1}", formato=FORMATO_MONEDA)
    for j, w in enumerate([20, 18, 22, 18, 16], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = w
    ws2.freeze_panes = "A3"

    # ---- Hoja 3: detalle auditable de la plantilla -----------------------
    ws3 = wb.create_sheet("Detalle Plantilla 87")
    ws3.sheet_view.showGridLines = False
    for j, t in enumerate(["Hoja fuente", "Fila", "Cargo",
                           "Grupo propuesto", "Salario propuesto (CUP)"], start=1):
        _celda(ws3, f"{get_column_letter(j)}1", t, negrita=True,
               color=BLANCO, relleno=AZUL)
    for i, reg in enumerate(registros, start=2):
        banda = GRIS if i % 2 else BLANCO
        _celda(ws3, f"A{i}", reg["hoja"], relleno=banda, alineacion="left")
        _celda(ws3, f"B{i}", reg["fila"], relleno=banda)
        _celda(ws3, f"C{i}", reg["cargo"], relleno=banda, alineacion="left")
        _celda(ws3, f"D{i}", reg["grupo"], negrita=True, relleno=banda)
        _celda(ws3, f"E{i}", reg["salario"], relleno=banda,
               formato=FORMATO_MONEDA_ENTERO)
    for j, w in enumerate([26, 8, 52, 14, 20], start=1):
        ws3.column_dimensions[get_column_letter(j)].width = w
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:E{len(registros)+1}"

    # ---- Hoja 4: metodología y fuentes -----------------------------------
    ws4 = wb.create_sheet("Metodologia y fuentes")
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions["A"].width = 110
    texto = [
        ("METODOLOGÍA Y FUENTES", True),
        ("", False),
        ("1. Salario mínimo: $3,210.00 mensuales, Resolución 14/2026 del MTSS "
         "(GOC-2026-393-O60, Gaceta Oficial No. 60 Ordinaria de 17/07/2026, p. 1839).", False),
        ("2. Escala oficial de 32 grupos (I a XXXII), régimen de 44 horas: "
         "Resolución 15/2026 (GOC-2026-394-O60), pp. 1839-1840. Los valores son los "
         "publicados oficialmente (redondeados).", False),
        ("3. Escala anterior: Resolución 29/2020, leída de TES/Salariales/Escala 3210_OK.xlsx.", False),
        ("4. Columna “Salario propuesto Plantilla Decreto 87”: extraída del libro "
         "“propuesta de plantilla por la resolucion 87.xls” (9 hojas de cargos). "
         "Por cada planta se localiza el encabezado G.E.A/G.E.P (grupo escala propuesto) "
         "y se toma el salario de la columna siguiente; cuando un grupo aparece con varios "
         "salarios, se adopta el valor mayoritario.", False),
        ("5. Grupos sin propuesta en la plantilla: I, II, III y XIX al XXXII "
         "(la plantilla solo propone cargos entre los grupos IV y XVIII).", False),
        ("6. La hoja “Escala por correlacion” replica el cálculo automático "
         "(salario_n = correlación_n × salario_n-1) con fórmulas vivas de Excel; "
         "difiere levemente de la escala oficial porque esta última está redondeada.", False),
        ("", False),
        ("ANOMALÍAS DETECTADAS EN LA PLANTILLA (.xls)", True),
    ]
    for a in anomalias:
        texto.append(("• " + a, False))
    texto += [
        ("", False),
        ("LENGUAJE DE PROGRAMACIÓN UTILIZADO: Python 3 — librerías xlrd (lectura del "
         ".xls heredado), openpyxl (Excel con fórmulas/estilos) y fpdf2 (PDF). "
         "Es la opción más adecuada porque lee formatos legacy, permite auditoría "
         "registro a registro, se re-ejecuta sin intervención manual y genera ambos "
         "formatos de salida con un único código mantenible.", False),
    ]
    for i, (t, negrita) in enumerate(texto, start=1):
        c = ws4.cell(row=i, column=1, value=t)
        c.font = Font(bold=negrita, size=12 if negrita else 10,
                      color=AZUL if negrita else "000000")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws4.row_dimensions[i].height = 30 if len(t) > 120 else 16

    wb.save(SALIDA_XLSX)
    print(f"  Excel generado: {SALIDA_XLSX.name}")


# ---------------------------------------------------------------------------
# 5. PDF (A4 horizontal)
# ---------------------------------------------------------------------------
FUENTE = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FUENTE_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


class EscalaPDF(FPDF):
    def header(self):
        self.set_font("dejavu", "B", 13)
        self.cell(0, 8, "NUEVA ESCALA SALARIAL — SALARIO MÍNIMO $3,210.00",
                  align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_font("dejavu", "", 9)
        self.cell(0, 5,
                  "Resoluciones 14/2026 y 15/2026 (GOC-2026-O60) · "
                  "Propuesta de plantilla Decreto 87 · montos en CUP",
                  align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_font("dejavu", "", 7.5)
        self.cell(0, 5,
                  f"Generado el {dt.date.today().strftime('%d/%m/%Y')} · "
                  f"Página {self.page_no()} de {{nb}} · fuente: propuesta de "
                  "plantilla resolución 87.xls",
                  align="C")


def _moneda(v, decimales=True):
    if v is None:
        return "—"
    plantilla = f"${abs(v):,.2f}" if decimales else f"${abs(v):,.0f}"
    return f"-{plantilla}" if v < 0 else plantilla


def escribir_pdf(filas, anomalias):
    pdf = EscalaPDF(orientation="L", unit="mm", format="A4")
    pdf.add_font("dejavu", "", FUENTE)
    pdf.add_font("dejavu", "B", FUENTE_B)
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    anchos = [18, 34, 24, 38, 30, 22, 36, 34]
    titulos = ["Grupo", "Escala anterior\nRes. 29/2020", "Coef. de\ncorrelación",
               "Nueva escala oficial\nRes. 15/2026 (44 h)", "Diferencia",
               "Incremento %", "Plantilla\nDecreto 87",
               "Diferencia\nplantilla − oficial"]
    alto_fila, alto_linea = 6.5, 4.2

    def _n_lineas(texto, ancho):
        total = 0
        for segmento in texto.split("\n"):
            if not segmento:
                total += 1
            else:
                total += max(1, math.ceil(
                    pdf.get_string_width(segmento) / (ancho - 2)))
        return total

    def dibujar_encabezado():
        pdf.set_font("dejavu", "B", 8)
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        max_lineas = max(_n_lineas(t, w) for t, w in zip(titulos, anchos))
        alto_total = max_lineas * alto_linea
        x0, y0 = pdf.get_x(), pdf.get_y()
        for w, t in zip(anchos, titulos):
            x = pdf.get_x()
            # rectángulo azul completo con borde
            pdf.rect(x, y0, w, alto_total, style="DF")
            # texto centrado verticalmente dentro del rectángulo
            alto_texto = _n_lineas(t, w) * alto_linea
            pdf.set_xy(x, y0 + (alto_total - alto_texto) / 2)
            pdf.multi_cell(w, alto_linea, t, align="C",
                           new_x="RIGHT", new_y="TOP")
        pdf.set_xy(x0, y0 + alto_total)
        pdf.set_text_color(0, 0, 0)

    dibujar_encabezado()
    for i, f in enumerate(filas):
        # saltar de página repitiendo el encabezado
        if pdf.get_y() + alto_fila > pdf.page_break_trigger:
            pdf.add_page()
            dibujar_encabezado()
        relleno = (242, 242, 242) if i % 2 else (255, 255, 255)
        valores = [
            (f["grupo"], True, "C"),
            (_moneda(f["anterior"]), False, "R"),
            (f"{f['factor']:.4f}" if f["factor"] else "—", False, "C"),
            (_moneda(f["oficial"], decimales=False), True, "R"),
            (_moneda(f["diferencia"]), False, "R"),
            (f"{f['incremento_pct']*100:.2f}%", False, "C"),
            (_moneda(f["plantilla"], decimales=False), True, "R"),
            (_moneda(f["dif_plantilla"]), False, "R"),
        ]
        for j, (w, (txt, negrita, alineacion)) in enumerate(zip(anchos, valores)):
            if j == 6:
                pdf.set_fill_color(226, 239, 218)  # columna de la plantilla
            else:
                pdf.set_fill_color(*relleno)
            pdf.set_font("dejavu", "B" if negrita else "", 8.5)
            pdf.cell(w, alto_fila, txt, border=1, align=alineacion, fill=True,
                     new_x="RIGHT", new_y="TOP")
        pdf.ln(alto_fila)

    pdf.ln(3)
    pdf.set_font("dejavu", "B", 9)
    pdf.cell(0, 5, "Notas:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("dejavu", "", 8)
    notas = list(anomalias) + [
        "Grupos I a III y XIX a XXXII: la plantilla del Decreto 87 no incluye cargos "
        "propuestos para esos grupos (columnas en verde con “—”).",
        "La escala oficial (Res. 15/2026) es la publicación redondeada; el cálculo "
        "exacto por correlación se entrega en la hoja “Escala por correlacion” del Excel.",
    ]
    for n in notas:
        pdf.multi_cell(0, 4.5, "• " + n, new_x="LMARGIN", new_y="NEXT")

    pdf.output(SALIDA_PDF)
    print(f"  PDF generado:   {SALIDA_PDF.name}")


# ---------------------------------------------------------------------------
# 6. PROGRAMA PRINCIPAL
# ---------------------------------------------------------------------------
def main():
    print("=" * 70)
    print(" GENERADOR DE ESCALA SALARIAL — MÍNIMO $3,210.00")
    print("=" * 70)

    print("[1/4] Leyendo escala anterior (Res. 29/2020)...")
    salarios_res29 = cargar_salarios_res29()
    print(f"      {len(salarios_res29)} grupos cargados.")

    print("[2/4] Extrayendo salarios propuestos de la plantilla Decreto 87...")
    propuesta, contadores, registros, anomalias = extraer_propuesta_plantilla()
    print(f"      {len(registros)} cargos registrados en {len(contadores)} grupos.")
    for g in sorted(propuesta, key=lambda x: ROMANO_A_ENTERO[x]):
        print(f"        grupo {g:>6}: ${propuesta[g]:,.2f}")
    if anomalias:
        print("      ANOMALÍAS:")
        for a in anomalias:
            print("        - " + a)

    print("[3/4] Construyendo tabla de 32 grupos...")
    filas = construir_filas(salarios_res29, propuesta)

    print("[4/4] Exportando a Excel y PDF...")
    escribir_excel(filas, contadores, registros, anomalias)
    escribir_pdf(filas, anomalias)
    print("=" * 70)
    print(" FIN. Archivos listos en:", CARPETA)
    print("=" * 70)


if __name__ == "__main__":
    main()
