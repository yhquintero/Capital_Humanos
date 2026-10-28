from __future__ import annotations

from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "TES" / "Salariales" / "Nueva Escala 3210" / "app_escala.py"
PAGES = [
    "Panel de control",
    "Escala salarial",
    "Comparar escenarios",
    "Plantilla por cargo",
    "Exportaciones",
    "Sistema y HTTPS",
    "Mejoras futuras",
]


def _assert_page_runs(app_test: AppTest, page: str) -> None:
    app_test.radio[0].set_value(page).run(timeout=60)
    assert not app_test.exception, f"{page}: {[item.message for item in app_test.exception]}"


def test_every_navigation_section_renders_without_exception():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    assert not app_test.exception
    assert app_test.radio[0].options == PAGES
    for page in PAGES:
        _assert_page_runs(app_test, page)


def test_six_x_ceiling_holds_in_live_cargo_table():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Plantilla por cargo")
    app_test.slider[-1].set_value(6.0).run(timeout=60)
    assert not app_test.exception
    details = next(
        frame.value for frame in app_test.dataframe if "Id." in frame.value.columns
    )
    assert list(details["Id."]) == list(range(1, len(details) + 1))
    assert (details["Salario ajustado (CUP)"] <= details["Tope máximo 6× (CUP)"].round(2)).all()


def test_comparison_page_exposes_bounded_factors_and_exports():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Comparar escenarios")
    assert len(app_test.slider) >= 3  # factor global y escenarios A/B
    comparison_sliders = {slider.label: slider for slider in app_test.slider}
    for label in ("Escenario A · factor", "Escenario B · factor"):
        assert comparison_sliders[label].min == 1.0
        assert comparison_sliders[label].max == 6.0
    labels = [button.label.lower() for button in app_test.get("download_button")]
    assert any("csv + manifiesto" in label for label in labels)
    assert any("comparación excel" in label for label in labels)
    assert any("csv por cargo" in label for label in labels)
    assert any("por cargo excel" in label for label in labels)


def test_cargo_scenario_comparison_preserves_detail_ids_and_origin():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Comparar escenarios")
    comparison = next(
        frame.value for frame in app_test.dataframe
        if "Factor A" in frame.value.columns and "Id." in frame.value.columns
    )
    assert len(comparison) == 183
    assert list(comparison["Id."]) == list(range(1, 184))
    assert {"Hoja de origen", "Fila de origen", "Variación B − A (CUP)"}.issubset(comparison.columns)
    assert (comparison["Salario A (CUP)"] <= comparison["Salario fuente (CUP)"] * 6).all()
    assert (comparison["Salario B (CUP)"] <= comparison["Salario fuente (CUP)"] * 6).all()


def test_export_page_offers_separate_approval_report():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Exportaciones")
    labels = [button.label for button in app_test.get("download_button")]
    assert "Descargar borrador de aprobación" in labels
    assert "Descargar paquete completo (.zip)" in labels
    preflight = next(
        frame.value for frame in app_test.dataframe
        if {"Control", "Resultado", "Detalle"}.issubset(frame.value.columns)
    )
    assert set(preflight["Resultado"]) == {"Correcto"}


def test_system_page_reports_approved_source_integrity():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Sistema y HTTPS")
    assert any(metric.label == "Integridad SHA-256" and metric.value == "Verificada" for metric in app_test.metric)
    source_status = app_test.dataframe[-1].value
    assert set(source_status["status"]) == {"VERIFICADA"}


def test_quality_report_shows_traceable_source_findings():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Plantilla por cargo")
    quality_table = next(
        frame.value for frame in app_test.dataframe
        if "Tipo de incidencia" in frame.value.columns
    )
    assert set(quality_table["Tipo de incidencia"]).issubset({
        "Posible duplicado", "Discrepancia salarial de grupo",
        "Salario ausente", "Salario inválido", "Cargo sin descripción",
    })
    assert {"Hoja de origen", "Fila de origen", "Detalle"}.issubset(quality_table.columns)
    labels = [button.label.lower() for button in app_test.get("download_button")]
    assert any("incidencias + manifiesto" in label for label in labels)
    review = app_test.session_state["revision_calidad_actual"]
    assert len(review) == len(quality_table)
    assert set(review["Estado de revisión"]) == {"Pendiente"}
    assert {"ID incidencia", "Nota de revisión"}.issubset(review.columns)

    app_test.session_state["editor_revision_calidad"] = {
        "edited_rows": {0: {"Estado de revisión": "Revisada", "Nota de revisión": "Confirmado en el expediente."}},
        "added_rows": [],
        "deleted_rows": [],
    }
    app_test.run(timeout=60)
    retained = app_test.session_state["revision_calidad_actual"]
    assert retained.loc[0, "Estado de revisión"] == "Revisada"
    assert retained.loc[0, "Nota de revisión"] == "Confirmado en el expediente."


def test_filtered_csv_is_packaged_with_source_manifest():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Plantilla por cargo")
    labels = [button.label.lower() for button in app_test.get("download_button")]
    assert any("filtro csv + manifiesto" in label for label in labels)
