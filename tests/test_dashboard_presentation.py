from __future__ import annotations

from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "TES" / "Salariales" / "Nueva Escala 3210" / "app_escala.py"


def _render(language: str = "Español") -> AppTest:
    app = AppTest.from_file(str(APP), default_timeout=60).run(timeout=60)
    app.session_state["idioma_interfaz"] = language
    app.run(timeout=60)
    assert not app.exception, [item.message for item in app.exception]
    return app


def test_dashboard_has_executive_kpis_coverage_and_non_budget_caveat():
    app = _render()
    content = "\n".join(
        [item.value for item in app.markdown]
        + [item.value for item in app.subheader]
        + [item.value for item in app.caption]
    )
    assert "ch-kpi" in content
    assert "ch-donut" in content
    assert "Cobertura de la escala" in content
    assert "Atención prioritaria" in content
    assert "No se estima el gasto total de nómina" in content
    assert "Impacto total de nómina" in content
    assert "No estimado" in content
    assert {button.key for button in app.button}.issuperset(
        {"quick_open_scale", "quick_open_template", "quick_open_exports"}
    )


def test_dashboard_design_and_caveat_are_localized_in_english():
    app = _render("English")
    content = "\n".join(
        [item.value for item in app.markdown]
        + [item.value for item in app.subheader]
        + [item.value for item in app.caption]
    )
    assert "Scale coverage" in content
    assert "Priority attention" in content
    assert "Total payroll expenditure is not estimated" in content
    assert "Not estimated" in content
    assert any(box.label == "Chart type" for box in app.selectbox)


def test_dashboard_and_comparison_chart_breakdowns_and_types_are_selectable():
    app = _render()
    for key, options in (
        ("dashboard_chart_breakdown", [
            "Comparación oficial / propuesta", "Variación absoluta (CUP)", "Variación porcentual",
        ]),
        ("dashboard_chart_type", ["Barras horizontales", "Columnas", "Líneas", "Área"]),
    ):
        for option in options:
            next(box for box in app.selectbox if box.key == key).set_value(option).run(timeout=60)
            assert not app.exception, f"{key}={option}: {[e.message for e in app.exception]}"

    app.radio[0].set_value("Comparar escenarios").run(timeout=60)
    assert not app.exception
    for key, options in (
        ("comparison_chart_breakdown", [
            "Variación absoluta (CUP)", "Variación porcentual", "Comparación escenarios A/B",
        ]),
        ("comparison_chart_type", ["Barras horizontales", "Columnas", "Líneas", "Área"]),
    ):
        for option in options:
            next(box for box in app.selectbox if box.key == key).set_value(option).run(timeout=60)
            assert not app.exception, f"{key}={option}: {[e.message for e in app.exception]}"


def test_browser_scroll_positions_are_saved_only_in_current_session():
    source = APP.read_text(encoding="utf-8")
    assert "st.iframe(" in source
    assert "capital-humano-scroll-v1" in source
    assert "host.sessionStorage" in source
    assert "main?.scrollTop" in source
    assert "sidebar?.scrollTop" in source
    assert "overflow-y:auto;overscroll-behavior:contain" in source


def test_dashboard_chart_currency_labels_keep_requested_spacing_and_precision():
    source = APP.read_text(encoding="utf-8")
    assert "'$ ' + format(datum.value, ',.2f')" in source
    assert "ui.format_currency(value, UI_LANGUAGE)" in source
