# -*- coding: utf-8 -*-
"""Interfaz interactiva para recalcular e imprimir la propuesta salarial."""
from __future__ import annotations

import io
import os
import platform
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

st.set_page_config(page_title="Capital Humano | Panel de Control", page_icon="📊", layout="wide")


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
            techo = salario_base * base.FACTOR_MAXIMO_ABSOLUTO
            salario = base.aplicar_factor_con_tope(salario_base, factor)
            fuente = "Plantilla Resolución 87"
        else:
            # Los grupos sin cargos propuestos conservan la escala oficial vigente;
            # el mínimo del grupo I queda anclado en $3,210.
            salario_base = f["oficial"]
            techo = salario_base * base.FACTOR_MAXIMO_ABSOLUTO
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
        tope = base_salario * base.FACTOR_MAXIMO_ABSOLUTO
        ajustado = base.aplicar_factor_con_tope(base_salario, factor)
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


def crear_comparacion_escenarios(factor_a, factor_b, propuesta, anterior):
    """Devuelve una tabla auditable de diferencias salariales por grupo."""
    return pd.DataFrame(base.comparar_escenarios(factor_a, factor_b, propuesta, anterior))


def crear_excel_comparacion(comparacion: pd.DataFrame, factor_a: float, factor_b: float) -> bytes:
    """Crea un Excel de comparación, listo para filtrar e imprimir en A4."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparación por grupo"
    ws.sheet_view.showGridLines = False
    ws.append(["COMPARACIÓN DE ESCENARIOS · CAPITAL HUMANO"])
    ws.merge_cells("A1:F1")
    ws.append([f"Escenario A: {factor_a:.2f}× · Escenario B: {factor_b:.2f}× · CUP · {date.today():%d/%m/%Y}"])
    ws.merge_cells("A2:F2")
    ws.append(list(comparacion.columns))
    for fila in comparacion.itertuples(index=False, name=None):
        ws.append(list(fila))
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    for cell in ws[2]:
        cell.fill = PatternFill("solid", fgColor="EAF0F4")
        cell.font = Font(name="Calibri", size=10, italic=True, color=AZUL)
    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor=VERDE)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[3].height = 34
    edge = Side(style="hair", color="CBD5DC")
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for cell in row:
            cell.border = Border(bottom=edge)
            cell.alignment = Alignment(vertical="center", wrap_text=cell.column in (1, 6))
        for index in (2, 3, 4):
            row[index - 1].number_format = '"$"#,##0.00'
        row[4].number_format = "0.00%"
    for column, width in enumerate((13, 22, 22, 24, 22, 45), start=1):
        ws.column_dimensions[chr(64 + column)].width = width
    if ws.max_row > 3:
        table = Table(displayName="ComparacionEscenarios", ref=f"A3:F{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{ws.max_row}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:3"
    ws.print_area = f"A1:F{ws.max_row}"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


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


# ---- Interfaz profesional / panel de control -------------------------------
NAV = [
    "Panel de control",
    "Escala salarial",
    "Comparar escenarios",
    "Plantilla por cargo",
    "Exportaciones",
    "Sistema y HTTPS",
    "Mejoras futuras",
]

st.markdown("""
<style>
:root{--ch-navy:#15324b;--ch-teal:#16866b;--ch-gold:#e6b85c;--ch-ink:#233746;--ch-muted:#667b89}
.block-container{padding-top:1.3rem;padding-bottom:3rem;max-width:1520px}
[data-testid="stSidebar"]{border-right:1px solid #e5edf0}
[data-testid="stSidebar"] [data-testid="stMetric"]{background:#f3f8f7;border:1px solid #e1eeeb;padding:.55rem .7rem;border-radius:12px}
.ch-hero{background:linear-gradient(118deg,#10283c,#174b59 62%,#16866b);padding:1.45rem 1.7rem;border-radius:18px;color:white;margin-bottom:1.2rem;box-shadow:0 10px 28px rgba(21,50,75,.12)}
.ch-hero h1{color:white;font-size:2.05rem;line-height:1.2;margin:0}.ch-hero p{color:#d9e9e8;margin:.45rem 0 0;font-size:1rem}
.ch-kicker{font-size:.72rem;letter-spacing:.12em;text-transform:uppercase;color:#b8ddd4;font-weight:700}
.ch-card{border:1px solid #e3ebef;background:#fff;border-radius:15px;padding:1rem 1.1rem;height:100%;box-shadow:0 4px 15px rgba(21,50,75,.035)}
.ch-card h3{font-size:1.03rem;margin:.1rem 0 .45rem;color:var(--ch-navy)}
.ch-card p{margin:0;color:var(--ch-muted);font-size:.91rem}
.ch-muted{color:var(--ch-muted);font-size:.88rem}
.ch-status{display:inline-block;padding:.22rem .58rem;border-radius:20px;background:#e8f5ee;color:#236645;font-size:.78rem;font-weight:700}
hr{border-color:#e8eef1}
</style>
""", unsafe_allow_html=True)

try:
    propuesta, contadores, registros, anomalias, anterior = cargar_datos()
except Exception as exc:
    st.error(f"No se pudo cargar la plantilla de origen: {exc}")
    st.info("Comprueba que el archivo .xls de la plantilla y la escala histórica estén dentro del repositorio. Si las rutas son correctas, abre la guía de diagnóstico del README.")
    st.stop()

with st.sidebar:
    st.markdown("<div style='padding:.25rem 0 1rem'><div style='font-size:1.65rem'>▦ <b>Capital Humano</b></div><div style='color:#667b89;font-size:.83rem'>Panel de Escala Salarial</div></div>", unsafe_allow_html=True)
    st.markdown("#### Navegación")
    vista = st.radio("Secciones", NAV, label_visibility="collapsed", key="panel_seccion")
    st.divider()
    st.markdown("#### Escenario salarial")
    factor = st.slider(
        "Multiplicador común", min_value=1.0, max_value=6.0,
        value=1.0, step=0.05, format="%.2f×", key="factor_global",
        help="Se aplica por cargo. El límite técnico nunca excede seis veces el salario de origen.",
    )
    st.metric("Tope permitido", "6.00×")
    st.caption("El grupo I mantiene el mínimo oficial de $3,210.00 CUP.")
    st.divider()
    st.caption("**Fuentes activas**\n\nRes. 14/2026 · Res. 15/2026 (44 h) · plantilla propuesta .xls")

escala_df, detalle_df = crear_tablas(factor, propuesta, registros, anterior)
assert detalle_df["Salario ajustado (CUP)"].le(detalle_df["Tope máximo 6× (CUP)"].round(2)).all()
propuesta_rows = escala_df[escala_df["Fuente del salario"] == "Plantilla Resolución 87"]
variacion = detalle_df["Salario ajustado (CUP)"].sum() - detalle_df["Salario original (CUP)"].sum()


def encabezado(titulo, bajada):
    st.markdown(
        f"<div class='ch-hero'><div class='ch-kicker'>Capital Humano · Gestión salarial</div><h1>{titulo}</h1><p>{bajada}</p></div>",
        unsafe_allow_html=True,
    )


def mostrar_tabla_escala(tabla):
    st.dataframe(
        tabla, width="stretch", hide_index=True, height=490,
        column_config={
            "Salario anterior Res. 29/2020 (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Escala oficial Res. 15/2026 (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Base propuesta / referencia (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Tope máximo (6× base) (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Nueva propuesta (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Factor aplicado": st.column_config.NumberColumn(format="%.2f×"),
        },
    )


def mostrar_tabla_detalle(tabla):
    st.dataframe(
        tabla, width="stretch", hide_index=True, height=560,
        column_config={
            "Salario original (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Tope máximo 6× (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Salario ajustado (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Factor aplicado": st.column_config.NumberColumn(format="%.2f×"),
        },
    )


if vista == "Panel de control":
    encabezado("Panel de Control", "Resumen ejecutivo, validaciones y accesos directos para preparar la propuesta.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Mínimo legal de referencia", "$3,210.00", "Grupo I · CUP")
    c2.metric("Multiplicador seleccionado", f"{factor:.2f}×", "Rango permitido: 1–6×")
    c3.metric("Cargos encontrados", f"{len(detalle_df):,}", f"{detalle_df['Grupo propuesto'].nunique()} grupos con propuesta")
    c4.metric("Cambio total en los cargos", f"${variacion:,.2f}", "vs. salarios fuente")

    left, right = st.columns([1.7, 1])
    with left:
        st.subheader("Vista comparativa de la escala")
        grafico = propuesta_rows[[
            "Grupo", "Escala oficial Res. 15/2026 (CUP)", "Nueva propuesta (CUP)"
        ]].set_index("Grupo").rename(columns={
            "Escala oficial Res. 15/2026 (CUP)": "Oficial Res. 15/2026",
            "Nueva propuesta (CUP)": f"Propuesta {factor:.2f}×",
        })
        st.bar_chart(grafico, width="stretch", height=360)
        st.caption("Comparación de los grupos presentes en la plantilla fuente (IV–XVIII).")
    with right:
        st.subheader("Estado de revisión")
        fuentes_ok = base.RUTA_PLANTILLA.is_file() and base.RUTA_ESCALA_BASE.is_file()
        st.markdown(f"<span class='ch-status'>{'● FUENTES DISPONIBLES' if fuentes_ok else '● REVISAR FUENTES'}</span>", unsafe_allow_html=True)
        st.write(f"**Registros en plantilla:** {len(registros)}")
        st.write(f"**Grupos cubiertos:** {', '.join(sorted(propuesta, key=lambda g: base.ROMANO_A_ENTERO[g]))}")
        st.write(f"**Grupos sin propuesta:** {32 - len(propuesta)} de 32; conservan valores oficiales.")
        st.write(f"**Control de tope:** {'Correcto' if detalle_df['Salario ajustado (CUP)'].le(detalle_df['Tope máximo 6× (CUP)'].round(2)).all() else 'ERROR'}")
        st.write(f"**Discrepancias de origen:** {len(anomalias)} grupos con valores distintos.")
    st.divider()
    st.subheader("Acciones rápidas")
    a1, a2, a3 = st.columns(3)
    with a1:
        st.markdown("<div class='ch-card'><h3>📊 Revisar escala</h3><p>Consulta los 32 grupos, salarios oficiales y propuesta ajustada.</p></div>", unsafe_allow_html=True)
    with a2:
        st.markdown("<div class='ch-card'><h3>👥 Auditar plantilla</h3><p>Busca cargos, filtra por grupo y conserva hoja/fila de procedencia.</p></div>", unsafe_allow_html=True)
    with a3:
        st.markdown("<div class='ch-card'><h3>📄 Preparar entregables</h3><p>Exporta el escenario activo a un libro Excel o PDF imprimible.</p></div>", unsafe_allow_html=True)
    if anomalias:
        with st.expander("Observaciones detectadas en la fuente", expanded=False):
            for nota in anomalias:
                st.warning(nota)

elif vista == "Escala salarial":
    encabezado("Escala Salarial", f"Escenario en vivo · multiplicador {factor:.2f}× · límite absoluto 6× de la base de origen.")
    st.info("Los grupos con propuesta se recalculan desde los salarios del archivo de plantilla. Los grupos sin propuesta conservan la escala oficial de la Resolución 15/2026; el grupo I permanece en $3,210.00.")
    mostrar_tabla_escala(escala_df)
    with st.expander("Metodología y supuestos"):
        st.markdown("""
        - El multiplicador global se aplica al salario por cargo de la plantilla fuente.
        - El tope de seis veces se comprueba individualmente en cada registro.
        - Para la tabla consolidada de grupo se usa el valor mayoritario cuando el .xls contiene discrepancias.
        - La tabla detallada conserva cada importe individual para auditoría.
        - Los grupos ausentes en la plantilla no reciben importes inventados.
        """)

elif vista == "Comparar escenarios":
    encabezado("Comparar escenarios", "Contrasta dos factores en paralelo y revisa la diferencia salarial por grupo antes de decidir.")
    ca, cb = st.columns(2)
    factor_a = ca.slider("Escenario A · factor", 1.0, 6.0, 1.0, 0.05, format="%.2f×", key="comparador_a")
    factor_b = cb.slider("Escenario B · factor", 1.0, 6.0, 1.5, 0.05, format="%.2f×", key="comparador_b")
    comparacion_df = crear_comparacion_escenarios(factor_a, factor_b, propuesta, anterior)
    grupos_modificados = comparacion_df["Variación B − A (CUP)"].abs().gt(0.005).sum()
    maxima_variacion = comparacion_df["Variación B − A (CUP)"].abs().max()
    k1, k2, k3 = st.columns(3)
    k1.metric("Grupos con cambio", f"{grupos_modificados} de {len(comparacion_df)}")
    k2.metric("Mayor cambio por grupo", f"${maxima_variacion:,.2f}")
    k3.metric("Factores comparados", f"{factor_a:.2f}× vs. {factor_b:.2f}×")
    grafico_comparacion = comparacion_df.set_index("Grupo")[["Variación B − A (CUP)"]]
    st.bar_chart(grafico_comparacion, width="stretch", height=350)
    st.dataframe(
        comparacion_df, width="stretch", hide_index=True,
        column_config={
            "Escenario A (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Escenario B (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Variación B − A (CUP)": st.column_config.NumberColumn(format="$%.2f"),
            "Variación relativa": st.column_config.NumberColumn(format="percent"),
        },
    )
    st.warning("La suma de importes por grupo no equivale al impacto total de nómina: no se dispone aquí de la cantidad de plazas ocupadas/vacantes por grupo. No uses esta comparación como presupuesto de nómina.")
    out_csv, out_xlsx = st.columns(2)
    csv_comparacion = comparacion_df.to_csv(index=False).encode("utf-8-sig")
    with out_csv:
        st.download_button("Descargar comparación CSV", csv_comparacion, "Comparacion_escenarios_Capital_Humano.csv", "text/csv", width="stretch")
    with out_xlsx:
        xlsx_comparacion = crear_excel_comparacion(comparacion_df, factor_a, factor_b)
        st.download_button(
            "Descargar comparación Excel", xlsx_comparacion,
            f"Comparacion_escenarios_{factor_a:.2f}x_{factor_b:.2f}x.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

elif vista == "Plantilla por cargo":
    encabezado("Plantilla por cargo", "Detalle rastreable al archivo de origen con filtros y búsqueda rápida.")
    f1, f2, f3 = st.columns([1, 1, 2])
    grupos = sorted(detalle_df["Grupo propuesto"].unique(), key=lambda g: base.ROMANO_A_ENTERO[g])
    hojas = sorted(detalle_df["Hoja de origen"].unique())
    filtro_grupo = f1.multiselect("Grupo", grupos, placeholder="Todos los grupos")
    filtro_hoja = f2.multiselect("Hoja fuente", hojas, placeholder="Todas las hojas")
    busqueda = f3.text_input("Buscar cargo", placeholder="Ej.: director, técnico, operador…")
    filtrada = detalle_df.copy()
    if filtro_grupo:
        filtrada = filtrada[filtrada["Grupo propuesto"].isin(filtro_grupo)]
    if filtro_hoja:
        filtrada = filtrada[filtrada["Hoja de origen"].isin(filtro_hoja)]
    if busqueda.strip():
        filtrada = filtrada[filtrada["Cargo"].str.contains(busqueda.strip(), case=False, regex=False, na=False)]
    m1, m2, m3 = st.columns(3)
    m1.metric("Registros visibles", f"{len(filtrada):,}")
    m2.metric("Salario fuente total", f"${filtrada['Salario original (CUP)'].sum():,.2f}")
    m3.metric("Propuesta calculada", f"${filtrada['Salario ajustado (CUP)'].sum():,.2f}")
    mostrar_tabla_detalle(filtrada)
    csv = filtrada.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Descargar este filtro en CSV", csv, "Plantilla_Resolucion87_filtrada.csv", "text/csv")

elif vista == "Exportaciones":
    encabezado("Exportar y presentar", "Genera documentos consistentes con el escenario mostrado en el panel.")
    st.write(f"**Escenario activo:** {factor:.2f}× · **Cargos:** {len(detalle_df)} · **Fecha:** {date.today():%d/%m/%Y}")
    st.info("El Excel incluye hojas de escala y detalle auditable. El PDF presenta ambas tablas en páginas A4 horizontales.")
    with st.spinner("Preparando los archivos para descarga…"):
        excel_bytes = crear_excel(escala_df, detalle_df, factor)
        pdf_bytes = crear_pdf(escala_df, detalle_df, factor)
    e1, e2 = st.columns(2)
    with e1:
        st.markdown("<div class='ch-card'><h3>📗 Libro Excel</h3><p>32 grupos, fórmulas/formatos y detalle completo de la plantilla fuente.</p></div>", unsafe_allow_html=True)
        st.download_button(
            "Descargar Excel (.xlsx)", excel_bytes,
            f"Propuesta_Escala_Resolucion87_{factor:.2f}x.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    with e2:
        st.markdown("<div class='ch-card'><h3>📕 Informe PDF</h3><p>Escala resumida y listado paginado de cargos, listo para imprimir.</p></div>", unsafe_allow_html=True)
        st.download_button(
            "Descargar PDF imprimible", pdf_bytes,
            f"Propuesta_Escala_Resolucion87_{factor:.2f}x.pdf",
            "application/pdf", width="stretch",
        )
    st.caption("Consejo: antes de distribuir, revisa discrepancias en el origen y confirma el factor aprobado.")

elif vista == "Sistema y HTTPS":
    encabezado("Sistema y seguridad", "Comprobaciones de las fuentes, del entorno de ejecución y del canal seguro.")
    ssl_cert = str(st.get_option("server.sslCertFile") or "").strip()
    try:
        request_headers = {str(k).lower(): str(v) for k, v in st.context.headers.items()}
    except Exception:
        request_headers = {}
    proxy_tls = request_headers.get("x-forwarded-proto", "").lower() == "https"
    cert_file_ok = bool(ssl_cert and Path(ssl_cert).is_file())
    tls_state = "HTTPS confirmado por proxy" if proxy_tls else ("Certificado TLS nativo configurado" if cert_file_ok else "No confirmado por esta sesión")
    q1, q2, q3 = st.columns(3)
    q1.metric("Canal", tls_state)
    q2.metric("Plantilla fuente", "Disponible" if base.RUTA_PLANTILLA.is_file() else "No encontrada")
    q3.metric("Escala histórica", "Disponible" if base.RUTA_ESCALA_BASE.is_file() else "No encontrada")
    st.markdown("#### Información de ejecución")
    sys_df = pd.DataFrame([
        {"Comprobación": "Sistema operativo", "Resultado": platform.platform()},
        {"Comprobación": "Python", "Resultado": sys.version.split()[0]},
        {"Comprobación": "Directorio de la aplicación", "Resultado": str(CARPETA)},
        {"Comprobación": "Configuración de certificado nativo", "Resultado": "Presente" if cert_file_ok else "No detectada"},
        {"Comprobación": "Proxy HTTPS", "Resultado": "Detectado" if proxy_tls else "No detectado"},
        {"Comprobación": "Validación de tope 6×", "Resultado": "Correcta" if detalle_df["Salario ajustado (CUP)"].le(detalle_df["Tope máximo 6× (CUP)"].round(2)).all() else "Revisar"},
    ])
    st.dataframe(sys_df, width="stretch", hide_index=True)
    st.warning("HTTPS cifra y autentica el servidor; no sustituye controles de acceso a nivel de aplicación. Para uso público, usa Caddy + Let's Encrypt y autenticación Basic configurada en el despliegue. No expongas directamente el puerto 8501.")
    if os.name == "nt":
        control_script = CARPETA / "control_app.ps1"
        st.markdown("#### Inicio automático HTTPS en Windows")
        if control_script.is_file():
            st.success("Está disponible el controlador de inicio/estado de la app.")
            st.code(".\\control_app.ps1 -Accion instalar-autoinicio\n.\\control_app.ps1 -Accion estado\n.\\control_app.ps1 -Accion detener\n.\\control_app.ps1 -Accion quitar-autoinicio", language="powershell")
            st.caption("El inicio automático se programa al iniciar sesión del mismo usuario; la CA mkcert debe instalarse una vez en ese perfil.")
        else:
            st.info("El control de inicio automático se habilita con el script de gestión de Windows incluido en esta versión.")
    else:
        st.markdown("#### Autoarranque del servidor")
        st.info("En Windows usa el programador de tareas del script de control. En servidores Linux, se recomienda el despliegue Docker Compose con Caddy y políticas de reinicio; certificados persistentes en el volumen caddy_data.")
    with st.expander("Recomendaciones operativas"):
        st.markdown("""
        - Actualiza Python y las dependencias desde un entorno virtual x64 (Windows).
        - No compartas `.certs`, claves privadas, `deploy/.env` ni credenciales.
        - En producción publica solo 80/443 mediante Caddy; no abras 8501 a Internet.
        - Comprueba HTTPS desde el navegador y que la fecha/hostname coincida con el certificado.
        - Conserva copia de seguridad y restringe el acceso al repositorio si contiene datos sensibles.
        """)

elif vista == "Mejoras futuras":
    encabezado("Próximas mejoras", "Selecciona prioridades para orientar la siguiente versión; esta selección se procesa localmente y no se envía a ningún servidor.")
    opciones = [
        "Multiplicadores independientes por grupo",
        "Guardar escenarios con nombre y recuperarlos de forma segura",
        "Cálculo del impacto presupuestario por centro/campamento",
        "Importar y validar nuevas versiones de la plantilla .xls/.xlsx",
        "Exportación PDF con logo, firmas y campos de aprobación",
        "Control de usuarios, roles y bitácora de cambios",
        "Panel Windows con botones de iniciar/detener/abrir HTTPS",
        "Programar copias de seguridad y actualización de la escala",
        "Indicadores de nómina a partir de plazas autorizadas y verificadas",
        "Mejoras de accesibilidad y experiencia móvil (WCAG)",
        "Interfaz en inglés con referencias legales cubanas intactas",
        "Huella verificable de fuentes en los informes descargados",
    ]
    seleccionadas = st.multiselect("¿Qué mejoras priorizarías?", opciones, placeholder="Selecciona una o varias")
    prioridad = st.select_slider("Prioridad global", options=["Baja", "Media", "Alta", "Urgente"], value="Media")
    comentarios = st.text_area("Otras ideas o requisitos", placeholder="Describe los cambios que necesitas en futuras versiones…")
    resumen = ["SUGERENCIAS DE MEJORA · CAPITAL HUMANO", "", f"Prioridad: {prioridad}", "", "Mejoras seleccionadas:"]
    resumen += [f"- {x}" for x in seleccionadas] or ["- No se seleccionaron opciones."]
    resumen += ["", "Comentarios:", comentarios.strip() or "(sin comentarios)", ""]
    st.download_button("Descargar lista de mejoras (.txt)", "\n".join(resumen), "Mejoras_Capital_Humano.txt", "text/plain")
    st.markdown("#### Ideas recomendadas por impacto")
    st.markdown("1. **Guardar y recuperar escenarios** para comparar decisiones aprobadas.\n2. **Presupuesto real** solo cuando exista un registro autorizado de plazas ocupadas/vacantes.\n3. **Roles individuales y bitácora** antes de permitir colaboración con datos salariales.\n4. **Trazabilidad de fuentes y aprobación** en los informes imprimibles.")
    st.caption("Consulta criterios de aceptación y prioridades completas en la [hoja de ruta de Capital Humano](https://github.com/yhquintero/Capital_Humanos/blob/main/docs/ROADMAP.md). No adjuntes datos personales ni salarios reales a propuestas públicas de GitHub.")

st.markdown("---")
st.markdown("<div class='ch-muted'>Capital Humano · herramienta de modelación y apoyo. Validar normativa e importes antes de aplicar una propuesta salarial.</div>", unsafe_allow_html=True)
