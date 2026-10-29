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
