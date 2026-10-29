from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "TES" / "Salariales" / "Nueva Escala 3210" / "app_escala.py"


def _contrast_ratio(foreground: str, background: str) -> float:
    def luminance(color: str) -> float:
        channels = [int(color[index:index + 2], 16) / 255 for index in (1, 3, 5)]
        linear = [
            channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4
            for channel in channels
        ]
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

    lighter, darker = sorted((luminance(foreground), luminance(background)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)


def test_accessibility_styles_include_focus_reflow_and_reduced_motion_support():
    app_source = APP.read_text(encoding="utf-8")
    style_match = re.search(r"<style>(.*?)</style>", app_source, flags=re.DOTALL)
    assert style_match is not None
    styles = style_match.group(1)

    assert ":focus-visible" in styles
    assert "@media(max-width:640px)" in styles
    assert "@media(prefers-reduced-motion:reduce)" in styles
    assert "color-scheme:light" in styles


def test_accessibility_palette_meets_wcag_aa_for_small_text_pairs():
    app_source = APP.read_text(encoding="utf-8")
    style_match = re.search(r"<style>(.*?)</style>", app_source, flags=re.DOTALL)
    assert style_match is not None
    styles = style_match.group(1)
    variables = dict(re.findall(r"--([\w-]+):(#[0-9a-fA-F]{6})", styles))

    assert _contrast_ratio(variables["ch-muted"], "#ffffff") >= 4.5
    assert _contrast_ratio("#b8ddd4", "#174b59") >= 4.5
    assert _contrast_ratio("#ffffff", variables["ch-teal"]) >= 4.5
    assert _contrast_ratio("#236645", "#e8f5ee") >= 4.5
