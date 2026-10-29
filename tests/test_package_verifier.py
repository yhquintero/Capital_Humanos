from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import warnings
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tools.verificar_paquete_exportacion import verificar_paquete  # noqa: E402

VERIFIER = ROOT / "tools" / "verificar_paquete_exportacion.py"
SENSITIVE_MARKER = "NOMINA-PRIVADA-987654321"


def _package_files() -> dict[str, bytes]:
    return {
        "Propuesta_Escala_Resolucion87_1.25x.xlsx": f"contenido {SENSITIVE_MARKER}".encode(),
        "Propuesta_Escala_Resolucion87_1.25x.pdf": b"pdf content",
        "Informe_Aprobacion_Capital_Humano_1.25x.pdf": b"draft content",
        "Parametros_escenario.json": b'{"factor": 1.25}',
        "Incidencias_calidad_revisadas.csv": b"ID incidencia,Nota de revision\n",
        "manifiesto_fuentes.json": b'{"integrity_ok": true}',
        "LEEME.txt": b"El inventario detecta cambios; no es una firma digital.\n",
    }


def _hash_manifest(files: dict[str, bytes]) -> bytes:
    data = {
        "algorithm": "SHA-256",
        "coverage": "All package files except this hashes_salidas.json file.",
        "files": [
            {"file": name, "bytes": len(content), "sha256": hashlib.sha256(content).hexdigest()}
            for name, content in sorted(files.items())
        ],
    }
    return json.dumps(data).encode("utf-8")


def _write_package(
    path: Path,
    *,
    tampered: bool = False,
    omit: str | None = None,
    extra: tuple[str, bytes] | None = None,
    malformed_manifest: bool = False,
    duplicate: bool = False,
) -> Path:
    files = _package_files()
    hashes = _hash_manifest(files)
    if tampered:
        files["Propuesta_Escala_Resolucion87_1.25x.xlsx"] += b" altered"
    if omit:
        files.pop(omit, None)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)
        if extra:
            archive.writestr(*extra)
        archive.writestr("hashes_salidas.json", b"not-json" if malformed_manifest else hashes)
        if duplicate:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                archive.writestr("LEEME.txt", b"second copy")
    return path


def test_verifies_valid_package_offline_without_printing_salary_content(tmp_path):
    package = _write_package(tmp_path / "bundle.zip")
    result = verificar_paquete(package)
    assert result == {"ok": True, "checked_files": 7, "errors": []}

    completed = subprocess.run(
        [sys.executable, str(VERIFIER), str(package)],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
    )
    assert completed.returncode == 0
    assert "7 archivos verificados" in completed.stdout
    assert SENSITIVE_MARKER not in completed.stdout
    assert SENSITIVE_MARKER not in completed.stderr


def test_detects_changed_package_payload_without_exposing_contents(tmp_path):
    package = _write_package(tmp_path / "altered.zip", tampered=True)
    result = verificar_paquete(package)
    assert result["ok"] is False
    assert any("SHA-256" in error for error in result["errors"])

    completed = subprocess.run(
        [sys.executable, str(VERIFIER), str(package)],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
    )
    assert completed.returncode == 1
    assert SENSITIVE_MARKER not in completed.stdout
    assert SENSITIVE_MARKER not in completed.stderr


def test_rejects_missing_or_unexpected_package_entries(tmp_path):
    missing = _write_package(
        tmp_path / "missing.zip", omit="Informe_Aprobacion_Capital_Humano_1.25x.pdf"
    )
    missing_result = verificar_paquete(missing)
    assert missing_result["ok"] is False
    assert any("requeridos" in error for error in missing_result["errors"])

    extra = _write_package(tmp_path / "extra.zip", extra=("not-listed.txt", b"extra"))
    extra_result = verificar_paquete(extra)
    assert extra_result["ok"] is False
    assert any("inventario" in error or "estructura" in error for error in extra_result["errors"])


def test_rejects_malformed_hash_inventory_unsafe_paths_and_duplicates(tmp_path):
    malformed = _write_package(tmp_path / "malformed.zip", malformed_manifest=True)
    assert verificar_paquete(malformed)["ok"] is False

    unsafe = _write_package(tmp_path / "unsafe.zip", extra=("../payroll.txt", b"hidden"))
    assert any("ruta" in error for error in verificar_paquete(unsafe)["errors"])

    duplicate = _write_package(tmp_path / "duplicate.zip", duplicate=True)
    duplicate_result = verificar_paquete(duplicate)
    assert duplicate_result["ok"] is False
    assert any("duplicados" in error for error in duplicate_result["errors"])


def test_rejects_non_zip_input_without_echoing_path(tmp_path):
    not_a_zip = tmp_path / f"{SENSITIVE_MARKER}.txt"
    not_a_zip.write_text("not a zip", encoding="utf-8")
    result = verificar_paquete(not_a_zip)
    assert result["ok"] is False
    assert result["errors"] == ["No se pudo abrir un archivo ZIP válido."]
