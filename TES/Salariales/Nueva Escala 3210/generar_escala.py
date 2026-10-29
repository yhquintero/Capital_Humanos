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
   - Escala_Salarial_3210_Nueva.xlsx   (5 hojas, con trazabilidad SHA-256)
   - Escala_Salarial_3210_Nueva.pdf    (informe A4 horizontal y trazabilidad)

 Requisitos:  pip install xlrd openpyxl fpdf2
 (entorno virtual recomendado)
=============================================================================
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import math
import os
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path

import xlrd
from fpdf import FPDF
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
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
RUTA_MANIFIESTO_FUENTES = RAIZ_REPO / "docs" / "source_manifest.json"
VERSION_APP = "0.1.0-dev"
FUENTES_CONTROLADAS = (
    (
        "plantilla_res87",
        "Plantilla de cargos · propuesta Resolución 87",
        "Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls",
    ),
    ("escala_res29", "Escala histórica · Resolución 29/2020", "TES/Salariales/Escala 3210_OK.xlsx"),
    ("gaceta_2026", "Gaceta Oficial · Resoluciones 14/2026 y 15/2026", "TES/Salariales/goc-2026-o60.pdf"),
    ("motor_calculo", "Código de extracción y cálculo", "TES/Salariales/Nueva Escala 3210/generar_escala.py"),
    ("interfaz_app", "Código de la aplicación y exportadores", "TES/Salariales/Nueva Escala 3210/app_escala.py"),
    ("actualizador_manifest", "Herramienta de actualización del manifiesto", "tools/actualizar_manifiesto_fuentes.py"),
    ("verificador_paquete", "Verificador offline de paquetes de exportación", "tools/verificar_paquete_exportacion.py"),
)

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
FACTOR_MAXIMO_ABSOLUTO = 6.0


def aplicar_factor_con_tope(salario_base: float, factor: float) -> float:
    """Calcula una propuesta sin permitir que el factor exceda el máximo 6×."""
    salario = float(salario_base)
    multiplicador = float(factor)
    if not math.isfinite(salario) or salario < 0:
        raise ValueError("El salario base debe ser un número finito no negativo.")
    if not math.isfinite(multiplicador) or not 0 <= multiplicador <= FACTOR_MAXIMO_ABSOLUTO:
        raise ValueError(f"El factor debe estar entre 0× y {FACTOR_MAXIMO_ABSOLUTO:g}×.")
    return min(salario * multiplicador, salario * FACTOR_MAXIMO_ABSOLUTO)


def _sha256(path: Path) -> str:
    """Calcula SHA-256 en bloques para no cargar archivos grandes en memoria."""
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def crear_manifiesto_fuentes(root: Path = RAIZ_REPO) -> dict:
    """Crea el manifiesto de las entradas y del código que gobierna el cálculo."""
    root = Path(root).resolve()
    entries = []
    for source_id, label, relative_path in FUENTES_CONTROLADAS:
        file_path = root / relative_path
        if not file_path.is_file():
            raise FileNotFoundError(f"No se encontró la fuente que se va a registrar: {relative_path}")
        entries.append({
            "id": source_id,
            "label": label,
            "path": relative_path,
            "size_bytes": file_path.stat().st_size,
            "sha256": _sha256(file_path),
        })
    return {
        "manifest_schema": 1,
        "application_version": VERSION_APP,
        "algorithm": "SHA-256",
        "sources": entries,
    }


def verificar_manifiesto_fuentes(manifest_path: Path = RUTA_MANIFIESTO_FUENTES, root: Path = RAIZ_REPO) -> dict:
    """Compara los bytes actuales con las huellas aprobadas y reporta cualquier cambio."""
    manifest_path = Path(manifest_path)
    root = Path(root).resolve()
    approved = json.loads(manifest_path.read_text(encoding="utf-8"))
    approved_sources = {item["id"]: item for item in approved.get("sources", [])}
    actual_sources = []
    for source_id, label, relative_path in FUENTES_CONTROLADAS:
        expected = approved_sources.get(source_id, {})
        file_path = root / relative_path
        if not file_path.is_file():
            actual_sources.append({
                "id": source_id, "label": label, "path": relative_path,
                "expected_sha256": expected.get("sha256", ""),
                "actual_sha256": "", "size_bytes": None, "status": "FALTANTE",
            })
            continue
        actual_digest = _sha256(file_path)
        size = file_path.stat().st_size
        matches = (
            expected.get("path") == relative_path
            and expected.get("sha256") == actual_digest
            and expected.get("size_bytes") == size
        )
        actual_sources.append({
            "id": source_id, "label": label, "path": relative_path,
            "expected_sha256": expected.get("sha256", ""),
            "actual_sha256": actual_digest, "size_bytes": size,
            "status": "VERIFICADA" if matches else "MODIFICADA",
        })
    current_ids = {item["id"] for item in actual_sources}
    approved_ids = set(approved_sources)
    consistent_version = approved.get("application_version") == VERSION_APP
    integrity_ok = (
        approved.get("manifest_schema") == 1
        and approved.get("algorithm") == "SHA-256"
        and approved_ids == current_ids
        and len(approved.get("sources", [])) == len(approved_sources)
        and consistent_version
        and all(item["status"] == "VERIFICADA" for item in actual_sources)
    )
    checked_at = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()
    return {
        **approved,
        "checked_at_utc": checked_at,
        "integrity_ok": integrity_ok,
        "version_ok": consistent_version,
        "sources": actual_sources,
    }


def manifiesto_para_salida() -> dict:
    """Devuelve un estado utilizable incluso si el manifiesto no está disponible."""
    try:
        return verificar_manifiesto_fuentes()
    except Exception as exc:
        return {
            "application_version": VERSION_APP,
            "checked_at_utc": "No disponible",
            "integrity_ok": False,
            "version_ok": False,
            "sources": [],
            "error": f"{type(exc).__name__}: {exc}",
        }


def agregar_hoja_trazabilidad(wb: Workbook, manifest: dict) -> None:
    """Añade una hoja imprimible con hashes esperados y actuales."""
    ws = wb.create_sheet("Trazabilidad SHA-256")
    ws.sheet_view.showGridLines = False
    ws.append(["TRAZABILIDAD DE FUENTES · CAPITAL HUMANO"])
    ws.merge_cells("A1:G1")
    ws.append(["Versión de la aplicación", manifest.get("application_version", "desconocida")])
    ws.append(["Integridad del manifiesto", "VERIFICADA" if manifest.get("integrity_ok") else "REVISIÓN REQUERIDA"])
    ws.append(["Verificación UTC", manifest.get("checked_at_utc", "No disponible")])
    ws.append(["SHA-256 compara los bytes con el manifiesto versionado; no sustituye revisión normativa ni firma digital."])
    ws.merge_cells("A5:G5")
    ws.append(["ID", "Fuente", "Ruta del repositorio", "Estado", "SHA-256 aprobado", "SHA-256 actual", "Tamaño (bytes)"])
    for source in manifest.get("sources", []):
        ws.append([
            source.get("id", ""), source.get("label", ""), source.get("path", ""),
            source.get("status", "NO VERIFICADA"), source.get("expected_sha256", ""),
            source.get("actual_sha256", ""), source.get("size_bytes"),
        ])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.font = Font(name="Calibri", size=15, bold=True, color=BLANCO)
    for row_number in (2, 3, 4):
        ws.cell(row_number, 1).font = Font(bold=True, color=AZUL)
        ws.cell(row_number, 1).fill = PatternFill("solid", fgColor="EAF0F4")
    ws[5][0].font = Font(name="Calibri", size=9, italic=True, color="536875")
    ws[5][0].alignment = Alignment(wrap_text=True)
    for cell in ws[6]:
        cell.fill = PatternFill("solid", fgColor=VERDE)
        cell.font = Font(name="Calibri", size=9, bold=True, color=BLANCO)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[5].height = 30
    ws.row_dimensions[6].height = 36
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in (2, 3))
            cell.border = Border(bottom=Side(style="hair", color="CBD5DC"))
            if cell.column in (5, 6):
                cell.font = Font(name="Consolas", size=8)
        row[3].fill = PatternFill("solid", fgColor="E8F5EE" if row[3].value == "VERIFICADA" else "FFF2CC")
    for letter, width in zip("ABCDEFG", (20, 48, 76, 22, 68, 68, 18)):
        ws.column_dimensions[letter].width = width
    if ws.max_row > 6:
        table = Table(displayName="ManifiestoSHA256", ref=f"A6:G{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
        ws.auto_filter.ref = f"A6:G{ws.max_row}"
    ws.freeze_panes = "A7"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 2
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:6"
    ws.print_area = f"A1:G{ws.max_row}"


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


def _crear_hallazgo_calidad(tipo, hoja, fila, cargo, grupo, valor, detalle):
    return {
        "Tipo de incidencia": tipo,
        "Hoja de origen": hoja,
        "Fila de origen": fila,
        "Cargo": cargo or "(sin descripción)",
        "Grupo": grupo,
        "Valor detectado": valor,
        "Detalle": detalle,
    }


MAX_PLANTILLA_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_PLANTILLA_XLSX_EXPANDED_BYTES = 64 * 1024 * 1024
MAX_PLANTILLA_XLSX_ENTRIES = 200
MAX_PLANTILLA_SHEETS = 80
MAX_PLANTILLA_CELLS = 2_000_000


class _HojaXlsxAdaptada:
    """Expone celdas XLSX con la interfaz de lectura utilizada para XLS."""

    def __init__(self, worksheet):
        self._worksheet = worksheet
        self.name = worksheet.title
        self.nrows = worksheet.max_row or 0
        self.ncols = worksheet.max_column or 0

    def cell_value(self, row: int, column: int):
        return self._worksheet.cell(row=row + 1, column=column + 1).value


def _abrir_plantilla_subida(file_name: str, payload: bytes):
    """Abre XLS/XLSX en memoria tras comprobar tamaño y estructura básica."""
    extension = Path(file_name).suffix.casefold()
    if extension not in {".xls", ".xlsx"}:
        raise ValueError("Formato no admitido: usa un archivo .xls o .xlsx sin macros.")
    if not payload or len(payload) > MAX_PLANTILLA_UPLOAD_BYTES:
        raise ValueError("El archivo está vacío o supera el límite de 15 MiB.")

    if extension == ".xls":
        if not payload.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            raise ValueError("La extensión .xls no coincide con la firma del archivo.")
        workbook = xlrd.open_workbook(file_contents=payload, on_demand=True)
        if workbook.nsheets > MAX_PLANTILLA_SHEETS:
            workbook.release_resources()
            raise ValueError(f"El libro supera el máximo de {MAX_PLANTILLA_SHEETS} hojas.")
        sheets = workbook.sheets()
        total_cells = sum(sheet.nrows * sheet.ncols for sheet in sheets)
        if total_cells > MAX_PLANTILLA_CELLS:
            workbook.release_resources()
            raise ValueError("El libro supera el límite seguro de celdas para revisión.")
        return workbook, sheets, "xls"

    try:
        with zipfile.ZipFile(BytesIO(payload)) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_PLANTILLA_XLSX_ENTRIES:
                raise ValueError("El XLSX contiene demasiados componentes internos.")
            if "[Content_Types].xml" not in archive.namelist():
                raise ValueError("El archivo ZIP no tiene la estructura de un libro XLSX.")
            expanded_size = sum(info.file_size for info in infos)
            if expanded_size > MAX_PLANTILLA_XLSX_EXPANDED_BYTES:
                raise ValueError("El XLSX supera el límite seguro de tamaño descomprimido.")
            if any(info.file_size > 1024 * 1024 and info.compress_size == 0 for info in infos):
                raise ValueError("El XLSX contiene una entrada comprimida inválida.")
            if any(
                info.compress_size and info.file_size / info.compress_size > 200
                for info in infos
            ):
                raise ValueError("El XLSX supera el límite de compresión permitido.")
    except zipfile.BadZipFile as exc:
        raise ValueError("El archivo .xlsx no es un ZIP válido.") from exc

    workbook = load_workbook(BytesIO(payload), data_only=True, read_only=True)
    worksheets = workbook.worksheets
    if not worksheets or len(worksheets) > MAX_PLANTILLA_SHEETS:
        workbook.close()
        raise ValueError(f"El libro debe tener entre 1 y {MAX_PLANTILLA_SHEETS} hojas.")
    total_cells = sum((sheet.max_row or 0) * (sheet.max_column or 0) for sheet in worksheets)
    if total_cells > MAX_PLANTILLA_CELLS:
        workbook.close()
        raise ValueError("El libro supera el límite seguro de celdas para revisión.")
    return workbook, [_HojaXlsxAdaptada(sheet) for sheet in worksheets], "xlsx"


def extraer_propuesta_plantilla_con_calidad(file_bytes: bytes | None = None, file_name: str | None = None):
    """Extrae cargos y devuelve incidencias revisables sin modificar la fuente.

    Sin argumentos lee la plantilla oficial del repositorio. Con ``file_bytes``
    y ``file_name`` inspecciona en memoria un XLS/XLSX subido, sin guardarlo.
    Los posibles duplicados se conservan; ninguna incidencia se corrige aquí.
    """
    if file_bytes is None:
        libro = xlrd.open_workbook(RUTA_PLANTILLA)
        hojas = libro.sheets()
    else:
        if not file_name:
            raise ValueError("Se necesita el nombre de archivo para validar el formato.")
        libro, hojas, formato = _abrir_plantilla_subida(file_name, file_bytes)
    contadores: dict[str, Counter] = defaultdict(Counter)
    registros = []
    hallazgos = []
    vistos = {}

    hojas_con_encabezado = 0
    for hoja in hojas:
        if hoja.name == "Hoja1":
            continue  # resumen por centro de costos, no es plantilla de cargos
        fila_enc, col_grupo = _localizar_columna_ge(hoja)
        if col_grupo is None:
            continue
        hojas_con_encabezado += 1
        col_salario = col_grupo + 1
        for r in range(fila_enc + 1, hoja.nrows):
            grupo_raw = hoja.cell_value(r, col_grupo)
            if not (isinstance(grupo_raw, str) and grupo_raw.strip() in ROMANO_A_ENTERO):
                continue
            grupo = grupo_raw.strip()
            fila_origen = r + 1
            cargo = _mejor_cargo_de_la_fila(hoja, r, col_grupo)
            if not cargo:
                hallazgos.append(_crear_hallazgo_calidad(
                    "Cargo sin descripción", hoja.name, fila_origen, cargo, grupo, "",
                    "La fila tiene un grupo válido, pero no se encontró texto descriptivo de cargo.",
                ))

            raw_salario = hoja.cell_value(r, col_salario) if col_salario < hoja.ncols else None
            if raw_salario is None or (isinstance(raw_salario, str) and not raw_salario.strip()):
                hallazgos.append(_crear_hallazgo_calidad(
                    "Salario ausente", hoja.name, fila_origen, cargo, grupo, "",
                    "La fila tiene un grupo válido, pero la celda de salario está vacía.",
                ))
                continue
            try:
                salario = float(raw_salario)
            except (TypeError, ValueError, OverflowError):
                hallazgos.append(_crear_hallazgo_calidad(
                    "Salario inválido", hoja.name, fila_origen, cargo, grupo, str(raw_salario),
                    "El salario no se puede interpretar como número; no se incluye en el cálculo.",
                ))
                continue
            if not math.isfinite(salario) or salario <= 0:
                hallazgos.append(_crear_hallazgo_calidad(
                    "Salario inválido", hoja.name, fila_origen, cargo, grupo, str(raw_salario),
                    "Se esperaba un salario numérico, finito y mayor que cero; no se incluye en el cálculo.",
                ))
                continue

            registro = {
                "hoja": hoja.name, "fila": fila_origen, "cargo": cargo,
                "grupo": grupo, "salario": salario,
            }
            contadores[grupo][salario] += 1
            registros.append(registro)

            texto_cargo = unicodedata.normalize("NFKC", cargo).casefold()
            texto_cargo = " ".join(texto_cargo.split())
            clave_duplicado = (hoja.name.casefold(), grupo, texto_cargo, salario)
            if clave_duplicado in vistos:
                hoja_previa, fila_previa = vistos[clave_duplicado]
                hallazgos.append(_crear_hallazgo_calidad(
                    "Posible duplicado", hoja.name, fila_origen, cargo, grupo, f"${salario:,.2f}",
                    f"Coincide con {hoja_previa}, fila {fila_previa}. Se conserva para revisión; no se elimina.",
                ))
            else:
                vistos[clave_duplicado] = (hoja.name, fila_origen)

    mapa_canonico = {g: cont.most_common(1)[0][0] for g, cont in contadores.items()}
    anomalias = []
    for grupo in sorted(contadores, key=lambda item: ROMANO_A_ENTERO[item]):
        contador = contadores[grupo]
        if len(contador) <= 1:
            continue
        valores = ", ".join(
            f"${salario:,.2f} ({cantidad} ocurrencias)"
            for salario, cantidad in contador.most_common()
        )
        anomalias.append(
            f"Grupo {grupo}: valores encontrados -> {valores}. "
            f"Se adopta ${mapa_canonico[grupo]:,.2f} por ser el mayoritario."
        )
        frecuencia_maxima = max(contador.values())
        valores_mayoritarios = {
            salario for salario, cantidad in contador.items() if cantidad == frecuencia_maxima
        }
        detalle = (
            f"El grupo contiene varios salarios: {valores}. "
            f"La tabla consolidada actual usa ${mapa_canonico[grupo]:,.2f}; "
            "la fuente permanece intacta y debe revisarse antes del uso oficial."
        )
        for registro in registros:
            if registro["grupo"] != grupo:
                continue
            if len(valores_mayoritarios) == 1 and registro["salario"] in valores_mayoritarios:
                continue
            hallazgos.append(_crear_hallazgo_calidad(
                "Discrepancia salarial de grupo", registro["hoja"], registro["fila"],
                registro["cargo"], grupo, f"${registro['salario']:,.2f}", detalle,
            ))

    hallazgos.sort(key=lambda item: (
        str(item["Hoja de origen"]).casefold(),
        item["Fila de origen"],
        item["Tipo de incidencia"],
    ))
    resultado = (mapa_canonico, contadores, registros, anomalias, hallazgos)
    if file_bytes is None:
        return resultado
    close = getattr(libro, "close", None)
    if callable(close):
        close()
    release = getattr(libro, "release_resources", None)
    if callable(release):
        release()
    return (*resultado, {"sheets": len(hojas), "template_sheets": hojas_con_encabezado})


def validar_plantilla_subida(file_name: str, payload: bytes) -> dict:
    """Valida y previsualiza un archivo subido sin escribirlo ni alterar filas."""
    safe_name = Path(str(file_name).replace("\\", "/")).name
    safe_name = re.sub(r"[^A-Za-z0-9._ -]", "_", safe_name)[:100] or "plantilla"
    extension = Path(safe_name).suffix.casefold()
    result = {
        "filename": safe_name,
        "sha256": hashlib.sha256(payload).hexdigest() if isinstance(payload, bytes) else "",
        "size_bytes": len(payload) if isinstance(payload, bytes) else None,
        "ready": False,
        "errors": [],
        "warnings": [],
        "validation": [],
        "preview": [],
        "records": [],
        "proposal": {},
        "counters": {},
        "anomalies": [],
        "findings": [],
        "record_count": 0,
        "template_sheets": 0,
    }
    if not isinstance(payload, bytes):
        result["errors"].append("El contenido recibido no es binario.")
        return result
    if extension not in {".xls", ".xlsx"}:
        result["errors"].append("Formato no admitido: usa .xls o .xlsx sin macros.")
        return result
    if not payload or len(payload) > MAX_PLANTILLA_UPLOAD_BYTES:
        result["errors"].append("El archivo está vacío o supera el límite de 15 MiB.")
        return result

    try:
        extracted = extraer_propuesta_plantilla_con_calidad(payload, safe_name)
        proposal, counters, records, anomalies, findings, metadata = extracted
    except (ValueError, OSError, xlrd.XLRDError, zipfile.BadZipFile) as exc:
        result["errors"].append(str(exc)[:300])
        return result
    except Exception as exc:
        result["errors"].append(f"No se pudo leer el libro ({type(exc).__name__}).")
        return result

    template_sheets = metadata["template_sheets"]
    missing_or_invalid = sum(
        finding["Tipo de incidencia"] in {"Salario ausente", "Salario inválido"}
        for finding in findings
    )
    duplicate_count = sum(finding["Tipo de incidencia"] == "Posible duplicado" for finding in findings)
    result.update({
        "proposal": dict(proposal),
        "counters": counters,
        "records": records,
        "anomalies": anomalies,
        "findings": findings,
        "record_count": len(records),
        "template_sheets": template_sheets,
        "preview": records[:25],
        "ready": bool(records) and template_sheets > 0 and missing_or_invalid == 0,
    })
    result["validation"] = [
        {
            "Control": "Hojas con encabezado G.E.",
            "Resultado": "Correcto" if template_sheets else "ERROR",
            "Detalle": f"Hojas de plantilla detectadas: {template_sheets}.",
        },
        {
            "Control": "Cargos y salarios válidos",
            "Resultado": "Correcto" if records and missing_or_invalid == 0 else "ERROR",
            "Detalle": f"Filas utilizables: {len(records)}; salarios ausentes o inválidos: {missing_or_invalid}.",
        },
        {
            "Control": "Posibles duplicados conservados",
            "Resultado": "Revisar" if duplicate_count else "Correcto",
            "Detalle": f"Posibles duplicados detectados: {duplicate_count}; no se elimina ninguna fila.",
        },
    ]
    if duplicate_count:
        result["warnings"].append(
            f"Se detectaron {duplicate_count} posibles duplicados. Se conservan; revisa las filas de origen manualmente."
        )
    if anomalies:
        result["warnings"].append(
            f"Se detectaron discrepancias salariales en {len(anomalies)} grupos; la fuente no se corrige."
        )
    if not template_sheets:
        result["errors"].append("No se encontró una fila de encabezados G.E. en las primeras 25 filas de ninguna hoja.")
    if not records:
        result["errors"].append("No se encontró ninguna fila utilizable con grupo romano y salario positivo.")
    if missing_or_invalid:
        result["errors"].append(
            f"Hay {missing_or_invalid} filas con salario ausente o inválido; no se permite activar esta plantilla."
        )
    return result


def extraer_propuesta_plantilla():
    """Compatibilidad: devuelve los cuatro resultados históricos de extracción."""
    propuesta, contadores, registros, anomalias, _ = extraer_propuesta_plantilla_con_calidad()
    return propuesta, contadores, registros, anomalias


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


def comparar_escenarios(factor_a: float, factor_b: float, propuesta, salarios_res29):
    """Compara dos multiplicadores por grupo; no estima el costo de la nómina."""
    filas = []
    for fila in construir_filas(salarios_res29, propuesta):
        salario_base = fila["plantilla"]
        if salario_base is None:
            salario_a = salario_b = float(fila["oficial"])
            fuente = "Escala oficial Res. 15/2026; grupo sin propuesta"
        else:
            salario_a = aplicar_factor_con_tope(salario_base, factor_a)
            salario_b = aplicar_factor_con_tope(salario_base, factor_b)
            fuente = "Plantilla Resolución 87"
        variacion = salario_b - salario_a
        filas.append({
            "Grupo": fila["grupo"],
            "Escenario A (CUP)": salario_a,
            "Escenario B (CUP)": salario_b,
            "Variación B − A (CUP)": variacion,
            "Variación relativa": (variacion / salario_a) if salario_a else 0.0,
            "Fuente": fuente,
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
    if valor is None or (isinstance(valor, str) and valor.strip().casefold() in {"none", "nan", "<na>"}):
        valor = None
    elif isinstance(valor, float) and math.isnan(valor):
        valor = None
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
    anomalias_visibles = [
        _texto_sin_nulos(nota).strip()
        for nota in (anomalias or [])
        if _texto_sin_nulos(nota).strip()
    ]
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
           "NOTAS: " + " / ".join(anomalias_visibles) if anomalias_visibles else "NOTAS: sin anomalías.",
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
    for j, t in enumerate(["Id.", "Hoja fuente", "Fila", "Cargo",
                           "Grupo propuesto", "Salario propuesto (CUP)"], start=1):
        _celda(ws3, f"{get_column_letter(j)}1", t, negrita=True,
               color=BLANCO, relleno=AZUL)
    for i, reg in enumerate(registros, start=2):
        banda = GRIS if i % 2 else BLANCO
        _celda(ws3, f"A{i}", i - 1, negrita=True, relleno=AZUL_CLARO)
        _celda(ws3, f"B{i}", reg["hoja"], relleno=banda, alineacion="left")
        _celda(ws3, f"C{i}", reg["fila"], relleno=banda)
        _celda(ws3, f"D{i}", reg["cargo"], relleno=banda, alineacion="left")
        _celda(ws3, f"E{i}", reg["grupo"], negrita=True, relleno=banda)
        _celda(ws3, f"F{i}", reg["salario"], relleno=banda,
               formato=FORMATO_MONEDA_ENTERO)
    for j, w in enumerate([8, 26, 8, 52, 14, 20], start=1):
        ws3.column_dimensions[get_column_letter(j)].width = w
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:F{len(registros)+1}"

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

    agregar_hoja_trazabilidad(wb, manifiesto_para_salida())
    wb.save(SALIDA_XLSX)
    print(f"  Excel generado: {SALIDA_XLSX.name}")


# ---------------------------------------------------------------------------
# 5. PDF (A4 horizontal)
# ---------------------------------------------------------------------------
def _localizar_fuentes_pdf():
    """Busca fuentes Unicode instaladas en Linux, Windows o macOS."""
    windows = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    parejas = [
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
         Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
         Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf")),
        (windows / "arial.ttf", windows / "arialbd.ttf"),
        (windows / "calibri.ttf", windows / "calibrib.ttf"),
        (Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
         Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf")),
    ]
    for normal, negrita in parejas:
        if normal.is_file() and negrita.is_file():
            return str(normal), str(negrita)
    raise FileNotFoundError(
        "No se encontró una fuente Unicode compatible con PDF. "
        "Instala Arial (Windows) o DejaVu Sans (Linux/macOS)."
    )


FUENTE, FUENTE_B = _localizar_fuentes_pdf()


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
                  f"Versión {VERSION_APP} · Página {self.page_no()} de {{nb}} · fuente: propuesta de "
                  "plantilla resolución 87.xls",
                  align="C")


def _moneda(v, decimales=True):
    if v is None:
        return "—"
    plantilla = f"${abs(v):,.2f}" if decimales else f"${abs(v):,.0f}"
    return f"-{plantilla}" if v < 0 else plantilla


def _texto_sin_nulos(valor, por_defecto=""):
    """Evita imprimir nulos y sus marcadores como si fueran datos."""
    if valor is None:
        return por_defecto
    if isinstance(valor, str) and valor.strip().casefold() in {"none", "nan", "<na>"}:
        return por_defecto
    if isinstance(valor, float) and math.isnan(valor):
        return por_defecto
    return str(valor)


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
    anomalias_visibles = [
        _texto_sin_nulos(nota).strip()
        for nota in (anomalias or [])
        if _texto_sin_nulos(nota).strip()
    ]
    notas = anomalias_visibles + [
        "Grupos I a III y XIX a XXXII: la plantilla del Decreto 87 no incluye cargos "
        "propuestos para esos grupos (columnas en verde con “—”).",
        "La escala oficial (Res. 15/2026) es la publicación redondeada; el cálculo "
        "exacto por correlación se entrega en la hoja “Escala por correlacion” del Excel.",
    ]
    for n in notas:
        pdf.multi_cell(0, 4.5, "• " + n, new_x="LMARGIN", new_y="NEXT")

    manifest = manifiesto_para_salida()
    pdf.add_page()
    pdf.set_font("dejavu", "B", 12)
    pdf.cell(0, 8, "TRAZABILIDAD DE FUENTES · SHA-256", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("dejavu", "", 8)
    integridad = "VERIFICADA" if manifest.get("integrity_ok") else "REVISIÓN REQUERIDA"
    pdf.set_x(pdf.l_margin)
    verificado_utc = _texto_sin_nulos(manifest.get("checked_at_utc"), "No disponible")
    pdf.multi_cell(pdf.epw, 4.5, f"Versión {VERSION_APP} · Integridad: {integridad} · Verificado UTC: {verificado_utc}", new_x="LMARGIN", new_y="NEXT")
    pdf.multi_cell(pdf.epw, 4.5, "Las huellas comparan los bytes con docs/source_manifest.json; no certifican autoría, firma digital ni vigencia jurídica de una fuente.", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    for source in manifest.get("sources", []):
        label = _texto_sin_nulos(source.get("label"), "Fuente sin etiqueta")
        status = _texto_sin_nulos(source.get("status"), "NO VERIFICADA")
        path = _texto_sin_nulos(source.get("path"), "no disponible")
        approved_hash = _texto_sin_nulos(source.get("expected_sha256"), "no disponible")
        current_hash = _texto_sin_nulos(source.get("actual_sha256"), "no disponible")
        size = _texto_sin_nulos(source.get("size_bytes"), "no disponible")
        pdf.set_font("dejavu", "B", 8)
        pdf.multi_cell(pdf.epw, 4.2, f"{label} · {status}", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("dejavu", "", 7)
        pdf.multi_cell(pdf.epw, 4, f"Ruta: {path}\nSHA-256 aprobado: {approved_hash}\nSHA-256 actual: {current_hash} · Tamaño: {size} bytes", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)
    if manifest.get("error"):
        pdf.multi_cell(0, 4, f"Error de manifiesto: {manifest['error']}")

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
