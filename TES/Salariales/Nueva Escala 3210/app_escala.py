# -*- coding: utf-8 -*-
"""Interfaz interactiva para recalcular e imprimir la propuesta salarial."""
from __future__ import annotations

import io
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

CARPETA = Path(__file__).resolve().parent
RAIZ_REPO = CARPETA.parents[2]
if str(CARPETA) not in sys.path:
    sys.path.insert(0, str(CARPETA))
import generar_escala as base  # noqa: E402

# Permite el comando intuitivo `python app_escala.py` sin ejecutar widgets en
# modo bare. La ejecución normal de Streamlit ya tiene ScriptRunContext y sigue
# directamente al renderizado de la aplicación.
if __name__ == "__main__":
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    if get_script_run_ctx(suppress_warning=True) is None:
        from streamlit.web import cli as stcli

        sys.argv = ["streamlit", "run", str(Path(__file__).resolve()), *sys.argv[1:]]
        raise SystemExit(stcli.main())

AZUL = "15324B"
VERDE = "16866B"
DORADO = "E6B85C"
FUENTE = base.FUENTE
FUENTE_B = base.FUENTE_B

st.set_page_config(page_title="Escala salarial | Capital Humano", page_icon="📊", layout="wide")


@st.cache_data(show_spinner="Leyendo la plantilla y verificando los grupos…")
def cargar_datos():
    propuesta, contadores, registros, anomalias = base.extraer_propuesta_plantilla()
    anterior = base.cargar_salarios_res29()
    return propuesta, contadores, registros, anomalias, anterior


def crear_tablas(factor, propuesta, registros, anterior):
    filas = base.construir_filas(anterior, propuesta)
    escala = []
    for f in filas:
        salario_base = f["plantilla"]
        if salario_base is not None:
            techo = salario_base * 6
            salario = min(salario_base * factor, techo)
            fuente = "Plantilla Resolución 87"
        else:
            # Los grupos sin cargos propuestos conservan la escala oficial vigente;
            # el mínimo del grupo I queda anclado en $3,210.
            salario_base = f["oficial"]
            techo = salario_base * 6
            salario = salario_base
            fuente = "Escala oficial Res. 15/2026"
        escala.append({
            "Grupo": f["grupo"],
            "Salario anterior Res. 29/2020 (CUP)": f["anterior"],
            "Escala oficial Res. 15/2026 (CUP)": f["oficial"],
            "Base propuesta / referencia (CUP)": salario_base,
            "Factor aplicado": factor if f["plantilla"] is not None else 1.0,
            "Tope máximo (6× base) (CUP)": techo,
            "Nueva propuesta (CUP)": salario,
            "Fuente del salario": fuente,
        })
    detalle = []
    for reg in registros:
        base_salario = float(reg["salario"])
        tope = base_salario * 6
        ajustado = min(base_salario * factor, tope)
        # Invariante de negocio: nunca superar seis veces el valor de origen.
        assert ajustado <= tope + 1e-9
        detalle.append({
            "Hoja de origen": reg["hoja"],
            "Fila de origen": reg["fila"],
            "Cargo": reg["cargo"],
            "Grupo propuesto": reg["grupo"],
            "Salario original (CUP)": base_salario,
            "Factor aplicado": factor,
            "Tope máximo 6× (CUP)": tope,
            "Salario ajustado (CUP)": ajustado,
        })
    return pd.DataFrame(escala), pd.DataFrame(detalle)


def crear_excel(escala: pd.DataFrame, detalle: pd.DataFrame, factor: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Escala interactiva"
    ws.sheet_view.showGridLines = False
    ws.append(["PROPUESTA SALARIAL · CAPITAL HUMANO"])
    ws.append([f"Factor: {factor:.2f}× · Límite: 6× del salario fuente · Generado: {date.today():%d/%m/%Y}"])
    ws.append(list(escala.columns))
    for row in escala.itertuples(index=False, name=None):
        ws.append(list(row))
    _estilizar_hoja(ws, 3, len(escala) + 3, "TablaEscala", [12, 25, 26, 27, 16, 25, 24, 32])
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:H{len(escala) + 3}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:3"
    ws.print_area = f"A1:H{len(escala)+3}"

    wd = wb.create_sheet("Plantilla Resolución 87")
    wd.sheet_view.showGridLines = False
    wd.append(["PROPUESTA ACTUALIZADA DE PLANTILLA · RESOLUCIÓN 87"])
    wd.append([f"Factor: {factor:.2f}× · Máximo permitido: 6× del salario original · Generado: {date.today():%d/%m/%Y}"])
    wd.append(list(detalle.columns))
    for row in detalle.itertuples(index=False, name=None):
        wd.append(list(row))
    _estilizar_hoja(wd, 3, len(detalle) + 3, "DetallePlantilla", [28, 14, 56, 18, 24, 17, 23, 25])
    wd.freeze_panes = "A4"
    wd.auto_filter.ref = f"A3:H{len(detalle) + 3}"
    wd.sheet_properties.pageSetUpPr.fitToPage = True
    wd.page_setup.orientation = "landscape"
    wd.page_setup.paperSize = wd.PAPERSIZE_A4
    wd.page_setup.fitToWidth = 1
    wd.page_setup.fitToHeight = 0
    wd.print_title_rows = "1:3"
    wd.print_area = f"A1:H{len(detalle)+3}"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _estilizar_hoja(ws, header_row, last_row, table_name, widths):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(widths))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(widths))
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    for cell in ws[2]:
        cell.fill = PatternFill("solid", fgColor="EAF0F4")
        cell.font = Font(name="Calibri", size=10, italic=True, color=AZUL)
    for cell in ws[header_row]:
        cell.fill = PatternFill("solid", fgColor=VERDE)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 36
    edge = Side(style="hair", color="CBD5DC")
    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_row):
        for cell in row:
            cell.border = Border(bottom=edge)
            cell.alignment = Alignment(vertical="center", wrap_text=(cell.column == 3))
            if cell.column in (2, 3, 4, 6, 7, 8) and isinstance(cell.value, (int, float)):
                cell.number_format = '"$"#,##0.00'
    if last_row > header_row:
        tab = Table(displayName=table_name, ref=f"A{header_row}:H{last_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                            showColumnStripes=False)
        ws.add_table(tab)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+i)].width = width


def crear_pdf(escala: pd.DataFrame, detalle: pd.DataFrame, factor: float) -> bytes:
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_font("DejaVu", "", FUENTE)
    pdf.add_font("DejaVu", "B", FUENTE_B)
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.set_margins(10, 10, 10)
    pdf.title = "Propuesta de escala y plantilla salarial"
    pdf.set_author("Capital Humano")

    def cabecera(titulo, subtitulo):
        pdf.set_font("DejaVu", "B", 15)
        pdf.set_text_color(21, 50, 75)
        pdf.cell(0, 8, titulo, new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font("DejaVu", "", 9)
        pdf.set_text_color(70, 85, 95)
        pdf.cell(0, 6, subtitulo, new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.ln(3)

    def tabla(df, titulos, widths, money_cols, font_size=7.3):
        pdf.set_font("DejaVu", "B", font_size)
        pdf.set_fill_color(22, 134, 107)
        pdf.set_text_color(255, 255, 255)
        for title, width in zip(titulos, widths):
            pdf.multi_cell(width, 8, str(title), border=1, align="C", fill=True,
                           new_x="RIGHT", new_y="TOP", max_line_height=4)
        pdf.ln(8)
        pdf.set_text_color(25, 39, 49)
        for ridx, row in enumerate(df.itertuples(index=False, name=None)):
            if pdf.get_y() > 185:
                pdf.add_page()
                pdf.set_font("DejaVu", "B", font_size)
                pdf.set_fill_color(22, 134, 107)
                pdf.set_text_color(255, 255, 255)
                for title, width in zip(titulos, widths):
                    pdf.multi_cell(width, 8, str(title), border=1, align="C", fill=True,
                                   new_x="RIGHT", new_y="TOP", max_line_height=4)
                pdf.ln(8)
                pdf.set_text_color(25, 39, 49)
            pdf.set_fill_color(244, 247, 248) if ridx % 2 else pdf.set_fill_color(255, 255, 255)
            pdf.set_font("DejaVu", "", font_size)
            height = 6
            for idx, (value, width) in enumerate(zip(row, widths)):
                if isinstance(value, (int, float)) and idx in money_cols:
                    text = f"${value:,.2f}"
                elif isinstance(value, float):
                    text = f"{value:.2f}×" if idx == 5 else f"{value:,.2f}"
                else:
                    text = str(value)
                if len(text) > 52 and width < 55:
                    text = text[:49] + "…"
                pdf.cell(width, height, text, border=1, align="R" if idx in money_cols else "L",
                         fill=True, new_x="RIGHT", new_y="TOP")
            pdf.ln(height)

    pdf.add_page()
    cabecera("NUEVA ESCALA SALARIAL", f"Mínimo $3,210.00 · Factor de plantilla {factor:.2f}× · Tope absoluto 6× del salario base · CUP")
    cols = ["Grupo", "Anterior\nRes. 29/2020", "Oficial\nRes. 15/2026", "Base", "Factor", "Tope 6×", "Nueva propuesta", "Fuente"]
    tabla(escala, cols, [14, 30, 30, 30, 18, 30, 31, 51], {1, 2, 3, 5, 6})
    pdf.ln(2)
    pdf.set_font("DejaVu", "", 8)
    pdf.multi_cell(0, 4.5, "El grupo I conserva el salario mínimo oficial de $3,210.00. Los grupos sin propuesta de cargos conservan la escala oficial. La tabla detallada siguiente aplica el factor a cada salario de la plantilla fuente y limita cada resultado a seis veces ese salario.")
    pdf.add_page()
    cabecera("PLANTILLA PROPUESTA · DETALLE DE CARGOS", f"Fuente: propuesta de plantilla por la resolución 87.xls · {len(detalle)} cargos · Factor {factor:.2f}×")
    det_cols = ["Hoja", "Fila", "Cargo", "Grupo", "Original", "Factor", "Tope 6×", "Ajustado"]
    tabla(detalle, det_cols, [33, 12, 75, 15, 27, 17, 27, 28], {4, 6, 7}, font_size=6.8)
    pdf.alias_nb_pages()
    for n in range(1, pdf.page_no() + 1):
        pdf.page = n
        pdf.set_y(-9)
        pdf.set_font("DejaVu", "", 7)
        pdf.set_text_color(100, 110, 118)
        pdf.cell(0, 5, f"Capital Humano · generado {date.today():%d/%m/%Y} · Página {n} de {{nb}}", align="C")
    return bytes(pdf.output())


# ---- Encabezado -----------------------------------------------------------
st.markdown("""
<style>
.block-container{padding-top:1.8rem;max-width:1500px}
.hero{background:linear-gradient(115deg,#15324b,#205b62);padding:1.5rem 1.8rem;border-radius:18px;color:#fff;margin-bottom:1rem}
.hero h1{color:white;margin:0;font-size:2rem}.hero p{color:#d9e9e8;margin:.45rem 0 0}
.small-note{color:#60717d;font-size:.88rem}
</style>
<div class="hero"><h1>📊 Escala salarial · Capital Humano</h1>
<p>Ajusta la propuesta en vivo, verifica el límite y exporta una versión lista para revisar e imprimir.</p></div>
""", unsafe_allow_html=True)

try:
    propuesta, contadores, registros, anomalias, anterior = cargar_datos()
except Exception as exc:
    st.error(f"No se pudo leer la plantilla de origen: {exc}")
    st.stop()

with st.sidebar:
    st.markdown("### Parámetros de la propuesta")
    factor = st.slider("Multiplicador salarial", min_value=1.0, max_value=6.0,
                       value=1.0, step=0.05, format="%.2f×",
                       help="El límite superior es seis veces el salario original de cada cargo en la plantilla fuente.")
    st.metric("Límite máximo", "6.00×")
    st.caption("El ajuste se actualiza inmediatamente. El grupo I permanece anclado al mínimo oficial de $3,210.00.")
    st.divider()
    st.markdown("**Fuentes**\n\n- Plantilla propuesta: archivo .xls de Resolución 87.\n- Escala vigente: Resolución 15/2026, 44 horas.\n- Mínimo: Resolución 14/2026.")

escala_df, detalle_df = crear_tablas(factor, propuesta, registros, anterior)
assert detalle_df["Salario ajustado (CUP)"].le(detalle_df["Tope máximo 6× (CUP)"].round(2)).all()
propuesta_rows = escala_df[escala_df["Fuente del salario"] == "Plantilla Resolución 87"]

k1, k2, k3, k4 = st.columns(4)
k1.metric("Salario mínimo", "$3,210.00", "Grupo I · Res. 14/2026")
k2.metric("Multiplicador actual", f"{factor:.2f}×", f"máximo 6.00×")
k3.metric("Cargos de la plantilla", f"{len(detalle_df):,}", "Detalle auditable")
k4.metric("Grupos con propuesta", f"{len(propuesta_rows)} de 32", "Grupos IV–XVIII")

st.subheader("Escala completa de 32 grupos")
st.caption("La columna verde se recalcula al mover el control. En los grupos sin propuesta de plantilla se conserva el valor oficial de la Resolución 15/2026.")
st.dataframe(escala_df, width="stretch", hide_index=True, height=440,
             column_config={
                 "Nueva propuesta (CUP)": st.column_config.NumberColumn("Nueva propuesta (CUP)", format="$%.2f", help="Resultado recalculado; limitado a 6× la base."),
                 "Tope máximo (6× base) (CUP)": st.column_config.NumberColumn("Tope máximo (6× base) (CUP)", format="$%.2f"),
                 "Factor aplicado": st.column_config.NumberColumn("Factor aplicado", format="%.2f×"),
             })

with st.expander("Ver plantilla completa por cargo (183 registros)", expanded=False):
    st.caption("Cada fila conserva su hoja y fila de origen. El salario ajustado nunca supera seis veces el valor original de esa fila.")
    st.dataframe(detalle_df, width="stretch", hide_index=True, height=520,
                 column_config={
                     "Salario original (CUP)": st.column_config.NumberColumn(format="$%.2f"),
                     "Tope máximo 6× (CUP)": st.column_config.NumberColumn(format="$%.2f"),
                     "Salario ajustado (CUP)": st.column_config.NumberColumn(format="$%.2f"),
                     "Factor aplicado": st.column_config.NumberColumn(format="%.2f×"),
                 })

st.subheader("Descargar o imprimir")
excel_bytes = crear_excel(escala_df, detalle_df, factor)
pdf_bytes = crear_pdf(escala_df, detalle_df, factor)
b1, b2, b3 = st.columns([1, 1, 2])
b1.download_button("⬇️ Descargar Excel (.xlsx)", data=excel_bytes,
                   file_name=f"Propuesta_Escala_Resolucion87_{factor:.2f}x.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   width="stretch")
b2.download_button("⬇️ Descargar PDF imprimible", data=pdf_bytes,
                   file_name=f"Propuesta_Escala_Resolucion87_{factor:.2f}x.pdf",
                   mime="application/pdf", width="stretch")
b3.info("El Excel incluye una hoja de escala y otra con el detalle completo de cargos. El PDF incluye ambas tablas, paginadas para impresión en A4 horizontal.")

if anomalias:
    with st.expander("Revisar observaciones de calidad de los datos"):
        for nota in anomalias:
            st.warning(nota)

st.markdown("---")
st.markdown('<p class="small-note">Herramienta de apoyo para modelar propuestas; validar los importes y las normas aplicables antes de su aprobación oficial. Fuentes detalladas en la documentación del repositorio.</p>', unsafe_allow_html=True)
