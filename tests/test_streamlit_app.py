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
    details = app_test.dataframe[0].value
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
    assert any("comparación csv" in label for label in labels)
    assert any("comparación excel" in label for label in labels)
