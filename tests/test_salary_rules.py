from __future__ import annotations

import contextlib
import io
import math
import runpy
import sys
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "TES" / "Salariales" / "Nueva Escala 3210"
sys.path.insert(0, str(APP_DIR))

import generar_escala as payroll  # noqa: E402


@pytest.mark.parametrize(
    ("salary", "factor", "expected"),
    [(3210, 1, 3210), (3210, 2.5, 8025), (500, 6, 3000), (500, 7, None)],
)
def test_salary_multiplier_never_exceeds_six_times(salary, factor, expected):
    if expected is None:
        with pytest.raises(ValueError):
            payroll.aplicar_factor_con_tope(salary, factor)
        return
    result = payroll.aplicar_factor_con_tope(salary, factor)
    assert result == pytest.approx(expected)
    assert result <= salary * 6


@pytest.mark.parametrize(
    ("salary", "factor"),
    [(-1, 1), (100, -0.1), (100, math.inf), (math.nan, 2), (math.inf, 1)],
)
def test_salary_multiplier_rejects_invalid_inputs(salary, factor):
    with pytest.raises(ValueError):
        payroll.aplicar_factor_con_tope(salary, factor)


def test_official_scale_has_32_groups_and_preserves_minimum():
    assert len(payroll.ESCALA_OFICIAL_RES15) == 32
    assert set(payroll.ESCALA_OFICIAL_RES15) == set(payroll.GRUPOS_ROMANOS)
    assert payroll.ESCALA_OFICIAL_RES15["I"] == 3210


def test_source_files_are_present_and_template_is_auditable():
    assert payroll.RUTA_PLANTILLA.is_file()
    assert payroll.RUTA_ESCALA_BASE.is_file()
    proposal, counters, records, anomalies = payroll.extraer_propuesta_plantilla()
    assert records
    assert len(records) >= 150
    assert set(proposal).issubset(payroll.ROMANO_A_ENTERO)
    assert all(record["hoja"] and record["fila"] > 0 and record["cargo"] for record in records)
    assert set(counters) == set(proposal)
    assert isinstance(anomalies, list)


def test_comparison_keeps_missing_groups_official_and_caps_proposed_groups():
    proposal, _, _, _ = payroll.extraer_propuesta_plantilla()
    previous = payroll.cargar_salarios_res29()
    comparison = payroll.comparar_escenarios(1, 6, proposal, previous)
    by_group = {row["Grupo"]: row for row in comparison}

    assert len(comparison) == 32
    assert by_group["I"]["Escenario A (CUP)"] == 3210
    assert by_group["I"]["Escenario B (CUP)"] == 3210
    assert by_group["XIX"]["Escenario A (CUP)"] == payroll.ESCALA_OFICIAL_RES15["XIX"]
    assert by_group["XIX"]["Escenario B (CUP)"] == payroll.ESCALA_OFICIAL_RES15["XIX"]

    proposed_group = next(iter(proposal))
    item = by_group[proposed_group]
    assert item["Escenario B (CUP)"] <= proposal[proposed_group] * 6
    assert item["Escenario B (CUP)"] == pytest.approx(proposal[proposed_group] * 6)
    assert item["Variación relativa"] == pytest.approx(
        item["Escenario B (CUP)"] / item["Escenario A (CUP)"] - 1
    )


def test_comparison_workbook_is_readable_and_print_ready():
    proposal, _, _, _ = payroll.extraer_propuesta_plantilla()
    previous = payroll.cargar_salarios_res29()
    comparison_rows = payroll.comparar_escenarios(1, 1.5, proposal, previous)
    app_file = APP_DIR / "app_escala.py"
    # runpy exposes the report function without opening a server; Streamlit's
    # harmless bare-context warning is redirected for this import-only check.
    with contextlib.redirect_stderr(io.StringIO()):
        namespace = runpy.run_path(str(app_file), run_name="capital_humano_report_test")
    frame = pd.DataFrame(comparison_rows)
    data = namespace["crear_excel_comparacion"](frame, 1, 1.5)
    workbook = load_workbook(BytesIO(data), data_only=True)
    sheet = workbook["Comparación por grupo"]

    assert sheet.max_row == 35
    assert sheet.max_column == 6
    assert sheet["A1"].value == "COMPARACIÓN DE ESCENARIOS · CAPITAL HUMANO"
    assert sheet["E4"].number_format == "0.00%"
    assert sheet.freeze_panes == "A4"
    assert sheet.page_setup.orientation == "landscape"
