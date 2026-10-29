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


def test_english_interface_renders_every_section_and_translates_source_status():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    app_test.session_state["idioma_interfaz"] = "English"
    app_test.run(timeout=60)
    english_pages = [
        "Dashboard", "Salary scale", "Compare Scenarios", "Staff Template",
        "Exports", "System and HTTPS", "Future improvements",
    ]
    assert not app_test.exception
    assert app_test.radio[0].options == english_pages
    for page in english_pages:
        app_test.radio[0].set_value(page).run(timeout=60)
        assert not app_test.exception, f"{page}: {[item.message for item in app_test.exception]}"
    app_test.radio[0].set_value("Salary scale").run(timeout=60)
    scale_columns = app_test.dataframe[0].value.columns.tolist()
    assert "Maximum cap (6× base) (CUP)" in scale_columns
    assert "Salary source" in scale_columns
    app_test.radio[0].set_value("Staff Template").run(timeout=60)
    detail_columns = app_test.dataframe[0].value.columns.tolist()
    assert detail_columns[:4] == ["ID", "Source sheet", "Source row", "Position"]
    app_test.radio[0].set_value("System and HTTPS").run(timeout=60)
    assert any(metric.label == "SHA-256 integrity" and metric.value == "Verified" for metric in app_test.metric)
    assert set(app_test.dataframe[-1].value["Status"]) == {"VERIFIED"}


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


def test_export_page_never_renders_pdf_none_return_values():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Exportaciones")
    rendered_none = [
        markdown.value for markdown in app_test.markdown
        if markdown.value.strip().casefold() in {"none", "`none`"}
    ]
    assert rendered_none == []
    quality_warnings = [
        warning.value for warning in app_test.warning
        if "incidencias de calidad" in warning.value.casefold()
    ]
    assert quality_warnings == []


def test_system_page_reports_approved_source_integrity():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Sistema y HTTPS")
    assert any(metric.label == "Integridad SHA-256" and metric.value == "Verificada" for metric in app_test.metric)
    source_status = app_test.dataframe[-1].value
    assert set(source_status["status"]) == {"VERIFICADA"}


def test_quality_page_shows_clean_state_after_authorized_source_review():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Plantilla por cargo")
    assert any(
        "No se detectaron incidencias" in success.value
        for success in app_test.success
    )
    assert not any(
        "Tipo de incidencia" in frame.value.columns
        for frame in app_test.dataframe
    )
    review = app_test.session_state["revision_calidad_actual"]
    assert review.empty
    assert {"ID incidencia", "Estado de revisión", "Nota de revisión"}.issubset(review.columns)
    labels = [button.label.lower() for button in app_test.get("download_button")]
    assert not any("incidencias + manifiesto" in label for label in labels)


def test_temporary_import_is_marked_unapproved_and_disables_salary_downloads():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    app_test.session_state["plantilla_importada"] = {
        "filename": "temporal.xlsx",
        "sha256": "a" * 64,
        "size_bytes": 128,
        "proposal": {"I": 3210.0},
        "counters": {},
        "records": [{
            "hoja": "Hoja de prueba", "fila": 2, "cargo": "Cargo de prueba",
            "grupo": "I", "salario": 3210.0,
        }],
        "anomalies": [],
        "findings": [],
    }
    app_test.radio[0].set_value("Comparar escenarios").run(timeout=60)
    assert not app_test.exception
    assert app_test.download_button
    assert all(button.disabled for button in app_test.download_button)
    assert any("Plantilla temporal activa" in warning.value for warning in app_test.warning)

    app_test.radio[0].set_value("Exportaciones").run(timeout=60)
    assert not app_test.exception
    assert app_test.download_button == []
    assert any("descargas quedan bloqueadas" in error.value for error in app_test.error)

    app_test.radio[0].set_value("Sistema y HTTPS").run(timeout=60)
    assert not app_test.exception
    sources = app_test.dataframe[-1].value
    assert "NO APROBADA · SOLO SESIÓN" in set(sources["status"])

    app_test.radio[0].set_value("Plantilla por cargo").run(timeout=60)
    assert any(button.label == "Volver a la plantilla oficial" for button in app_test.button)
    assert any(button.label == "Usar temporalmente en esta sesión" for button in app_test.button) is False
    assert all(button.disabled for button in app_test.download_button)


def test_filtered_csv_is_packaged_with_source_manifest():
    app_test = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    _assert_page_runs(app_test, "Plantilla por cargo")
    labels = [button.label.lower() for button in app_test.get("download_button")]
    assert any("filtro csv + manifiesto" in label for label in labels)
