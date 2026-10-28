#!/usr/bin/env python3
"""Actualiza huellas solo después de revisar y aprobar cambios de fuente/código."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "TES" / "Salariales" / "Nueva Escala 3210"
sys.path.insert(0, str(APP_DIR))

import generar_escala as base  # noqa: E402


if __name__ == "__main__":
    manifest = base.crear_manifiesto_fuentes(ROOT)
    destination = ROOT / "docs" / "source_manifest.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Manifiesto actualizado: {destination.relative_to(ROOT)}")
    print("Revisa los cambios de contenido, normativa y cálculo antes de confirmar el archivo.")
