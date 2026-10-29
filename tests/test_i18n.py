from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd

APP_DIR = Path(__file__).resolve().parents[1] / "TES" / "Salariales" / "Nueva Escala 3210"
sys.path.insert(0, str(APP_DIR))
import i18n


def test_english_currency_date_and_interpolated_messages():
    assert i18n.format_currency(1234.5, "en") == "$ 1,234.50"
    assert i18n.format_currency(3210, "es") == "$ 3,210.00"
    assert i18n.format_currency(-3210, "en") == "$ -3,210.00"
    assert i18n.format_currency(float("inf"), "es") == "—"
    assert i18n.format_date(date(2026, 10, 28), "en") == "2026-10-28"
    assert i18n.translate("Hay {count} incidencias de calidad pendientes de revisión. Puedes descargar los informes para revisión, pero no interpretes su generación como resolución de esas incidencias.", "en", count="7") == (
        "7 quality issues are pending review. You may download the reports for review, but their generation does not resolve those issues."
    )


def test_dataframe_translation_is_presentation_only_and_translates_categories():
    source = pd.DataFrame(
        {
            "Tipo de incidencia": ["Posible duplicado"],
            "Estado de revisión": ["En revisión"],
            "Nota de revisión": ["Pendiente de validar"],
        }
    )
    translated = i18n.translate_dataframe(source, "en")
    assert translated.columns.tolist() == ["Issue type", "Review status", "Review note"]
    assert translated.iloc[0]["Issue type"] == "Possible duplicate"
    assert translated.iloc[0]["Review status"] == "In review"
    assert source.iloc[0]["Tipo de incidencia"] == "Posible duplicado"
    assert source.columns.tolist() == ["Tipo de incidencia", "Estado de revisión", "Nota de revisión"]


def test_translation_dictionary_has_no_duplicate_literal_keys():
    assert len(i18n.ENGLISH) == len(set(i18n.ENGLISH))
