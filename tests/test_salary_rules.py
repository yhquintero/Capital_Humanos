from __future__ import annotations

import contextlib
import io
import json
import math
import runpy
import sys
import zipfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "TES" / "Salariales" / "Nueva Escala 3210"
sys.path.insert(0, str(APP_DIR))

import generar_escala as payroll  # noqa: E402


@pytest.fixture(scope="module")
def app_namespace():
    # Load export helpers without starting a Streamlit server.
    with contextlib.redirect_stderr(io.StringIO()):
        return runpy.run_path(str(APP_DIR / "app_escala.py"), run_name="capital_humano_export_test")


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


def test_comparison_workbook_is_readable_and_contains_source_hashes(app_namespace):
    proposal, _, _, _ = payroll.extraer_propuesta_plantilla()
    previous = payroll.cargar_salarios_res29()
    comparison_rows = payroll.comparar_escenarios(1, 1.5, proposal, previous)
    frame = pd.DataFrame(comparison_rows)
    manifest = payroll.verificar_manifiesto_fuentes()
    data = app_namespace["crear_excel_comparacion"](frame, 1, 1.5, manifest)
    workbook = load_workbook(BytesIO(data), data_only=True)
    sheet = workbook["Comparación por grupo"]
    provenance = workbook["Trazabilidad SHA-256"]

    assert sheet.max_row == 35
    assert sheet.max_column == 6
    assert sheet["A1"].value == "COMPARACIÓN DE ESCENARIOS · CAPITAL HUMANO"
    assert sheet["E4"].number_format == "0.00%"
    assert sheet.freeze_panes == "A4"
    assert sheet.page_setup.orientation == "landscape"
    assert provenance["B2"].value == payroll.VERSION_APP
    assert provenance["B3"].value == "VERIFICADA"
    assert provenance["D7"].value == "VERIFICADA"
    assert len(provenance["E7"].value) == 64


def test_filtered_csv_package_contains_json_manifest(app_namespace):
    manifest = payroll.verificar_manifiesto_fuentes()
    df = pd.DataFrame([{"Cargo": "Técnico", "Salario": 3210}])
    data = app_namespace["crear_zip_csv_con_manifiesto"](df, "filtro.csv", manifest)
    with zipfile.ZipFile(BytesIO(data)) as archive:
        assert set(archive.namelist()) == {"filtro.csv", "manifiesto_fuentes.json", "LEEME.txt"}
        payload = json.loads(archive.read("manifiesto_fuentes.json"))
        assert payload["integrity_ok"] is True
        assert len(payload["sources"]) == len(payroll.FUENTES_CONTROLADAS)
        assert b"Cargo,Salario" in archive.read("filtro.csv")


def test_approved_source_manifest_matches_all_tracked_files():
    manifest = payroll.verificar_manifiesto_fuentes()
    assert manifest["integrity_ok"] is True
    assert manifest["application_version"] == payroll.VERSION_APP
    assert len(manifest["sources"]) == len(payroll.FUENTES_CONTROLADAS)
    assert all(len(source["actual_sha256"]) == 64 for source in manifest["sources"])


def test_checked_in_reference_exports_include_verified_provenance():
    manifest = payroll.verificar_manifiesto_fuentes()
    example_xlsx = APP_DIR / "Escala_Salarial_3210_Nueva.xlsx"
    workbook = load_workbook(example_xlsx, read_only=True, data_only=True)
    assert "Trazabilidad SHA-256" in workbook.sheetnames
    assert workbook["Trazabilidad SHA-256"]["B3"].value == "VERIFICADA"

    example_pdf = APP_DIR / "Escala_Salarial_3210_Nueva.pdf"
    pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(str(example_pdf)).pages)
    assert "TRAZABILIDAD DE FUENTES" in pdf_text
    assert manifest["sources"][0]["actual_sha256"] in pdf_text


def test_modified_expected_hash_is_detected(tmp_path):
    approved = payroll.crear_manifiesto_fuentes()
    approved["sources"][0]["sha256"] = "0" * 64
    test_manifest = tmp_path / "source_manifest.json"
    test_manifest.write_text(json.dumps(approved), encoding="utf-8")

    result = payroll.verificar_manifiesto_fuentes(test_manifest)
    assert result["integrity_ok"] is False
    assert result["sources"][0]["status"] == "MODIFICADA"


def test_active_scenario_excel_and_pdf_include_traceability(app_namespace):
    proposal, _, records, _ = payroll.extraer_propuesta_plantilla()
    previous = payroll.cargar_salarios_res29()
    scale, details = app_namespace["crear_tablas"](1.25, proposal, records, previous)
    manifest = payroll.verificar_manifiesto_fuentes()
    workbook_bytes = app_namespace["crear_excel"](scale, details, 1.25, manifest)
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True)
    assert "Trazabilidad SHA-256" in workbook.sheetnames
    assert workbook["Trazabilidad SHA-256"]["B3"].value == "VERIFICADA"

    pdf_bytes = app_namespace["crear_pdf"](scale, details, 1.25, manifest)
    assert pdf_bytes.startswith(b"%PDF-")
    pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(BytesIO(pdf_bytes)).pages)
    assert "TRAZABILIDAD DE FUENTES" in pdf_text
    assert manifest["sources"][0]["actual_sha256"] in pdf_text
