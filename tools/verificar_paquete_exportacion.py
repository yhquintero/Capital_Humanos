#!/usr/bin/env python3
"""Verifica offline la integridad del ZIP de exportación de Capital Humano.

La herramienta comprueba el inventario SHA-256 sin extraer los archivos ni
mostrar su contenido. Devuelve código 0 si todos los controles pasan y 1 si
el paquete es inválido.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

HASH_FILENAME = "hashes_salidas.json"
EXPECTED_COVERAGE = "All package files except this hashes_salidas.json file."
MAX_ARCHIVE_MEMBERS = 32
MAX_MEMBER_BYTES = 64 * 1024 * 1024
MAX_TOTAL_BYTES = 128 * 1024 * 1024
CHUNK_BYTES = 64 * 1024
SHA256_PATTERN = re.compile(r"^[0-9a-f]{64}$")
REQUIRED_FIXED_FILES = {
    "Parametros_escenario.json",
    "Incidencias_calidad_revisadas.csv",
    "manifiesto_fuentes.json",
    "LEEME.txt",
    HASH_FILENAME,
}
REQUIRED_ARTIFACT_GROUPS = (
    ("Excel salarial", "Propuesta_Escala_Resolucion87_", ".xlsx"),
    ("PDF salarial", "Propuesta_Escala_Resolucion87_", ".pdf"),
    ("borrador de aprobación", "Informe_Aprobacion_Capital_Humano_", ".pdf"),
)


def _is_flat_safe_name(name: str) -> bool:
    """Generated packages are flat; reject traversal, path separators and NUL."""
    return bool(
        name
        and name not in {".", ".."}
        and "\x00" not in name
        and "/" not in name
        and "\\" not in name
        and not name.startswith(".")
    )


def _valid_required_artifacts(names: set[str]) -> bool:
    if not REQUIRED_FIXED_FILES.issubset(names):
        return False
    return len(names) == len(REQUIRED_FIXED_FILES) + len(REQUIRED_ARTIFACT_GROUPS) and all(
        sum(name.startswith(prefix) and name.endswith(suffix) for name in names) == 1
        for _, prefix, suffix in REQUIRED_ARTIFACT_GROUPS
    )


def _read_hash_manifest(archive: zipfile.ZipFile) -> tuple[dict[str, dict[str, Any]] | None, str | None]:
    try:
        if archive.getinfo(HASH_FILENAME).file_size > 1024 * 1024:
            return None, "El inventario SHA-256 supera el tamaño permitido."
        payload = json.loads(archive.read(HASH_FILENAME).decode("utf-8"))
    except (KeyError, UnicodeDecodeError, json.JSONDecodeError, OSError, RuntimeError, zipfile.BadZipFile):
        return None, "El inventario SHA-256 falta o no es JSON UTF-8 válido."

    if not isinstance(payload, dict) or payload.get("algorithm") != "SHA-256":
        return None, "El inventario no declara el algoritmo SHA-256 esperado."
    if payload.get("coverage") != EXPECTED_COVERAGE:
        return None, "La cobertura declarada por el inventario no coincide con el formato del paquete."
    records = payload.get("files")
    if not isinstance(records, list):
        return None, "El inventario no contiene una lista válida de archivos."

    result: dict[str, dict[str, Any]] = {}
    for record in records:
        if not isinstance(record, dict):
            return None, "El inventario contiene una entrada con estructura inválida."
        name = record.get("file")
        size = record.get("bytes")
        digest = record.get("sha256")
        if not isinstance(name, str) or not _is_flat_safe_name(name) or name == HASH_FILENAME:
            return None, "El inventario contiene un nombre de archivo no permitido."
        if name in result:
            return None, "El inventario contiene nombres de archivo duplicados."
        if isinstance(size, bool) or not isinstance(size, int) or size < 0:
            return None, "El inventario contiene un tamaño inválido."
        if not isinstance(digest, str) or not SHA256_PATTERN.fullmatch(digest):
            return None, "El inventario contiene una huella SHA-256 inválida."
        result[name] = record
    return result, None


def _hash_member(archive: zipfile.ZipFile, info: zipfile.ZipInfo) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    with archive.open(info, "r") as stream:
        while chunk := stream.read(CHUNK_BYTES):
            size += len(chunk)
            if size > MAX_MEMBER_BYTES:
                raise ValueError("member-size-limit")
            digest.update(chunk)
    return digest.hexdigest(), size


def verificar_paquete(path: str | Path) -> dict[str, Any]:
    """Return a content-safe summary after verifying package structure and hashes."""
    errors: list[str] = []
    checked_files = 0
    try:
        archive = zipfile.ZipFile(path, "r")
    except (OSError, zipfile.BadZipFile, ValueError):
        return {"ok": False, "checked_files": 0, "errors": ["No se pudo abrir un archivo ZIP válido."]}

    with archive:
        infos = archive.infolist()
        if not infos:
            return {"ok": False, "checked_files": 0, "errors": ["El ZIP está vacío."]}
        if len(infos) > MAX_ARCHIVE_MEMBERS:
            return {"ok": False, "checked_files": 0, "errors": ["El ZIP supera el límite de archivos permitido."]}

        names: list[str] = []
        total_declared_bytes = 0
        for info in infos:
            if info.is_dir() or not _is_flat_safe_name(info.filename):
                errors.append("El ZIP contiene una ruta o entrada no permitida.")
            names.append(info.filename)
            if info.file_size < 0 or info.file_size > MAX_MEMBER_BYTES:
                errors.append("Un archivo del ZIP supera el tamaño individual permitido.")
            total_declared_bytes += max(info.file_size, 0)
        if total_declared_bytes > MAX_TOTAL_BYTES:
            errors.append("El ZIP supera el tamaño total descomprimido permitido.")
        if len(names) != len(set(names)):
            errors.append("El ZIP contiene nombres de archivo duplicados.")
        if errors:
            return {"ok": False, "checked_files": 0, "errors": errors}

        name_set = set(names)
        if not _valid_required_artifacts(name_set):
            errors.append("Faltan artefactos requeridos o el paquete contiene una estructura no compatible.")

        hash_records, manifest_error = _read_hash_manifest(archive)
        if manifest_error:
            errors.append(manifest_error)
        elif hash_records is not None:
            actual_names = name_set - {HASH_FILENAME}
            if set(hash_records) != actual_names:
                errors.append("La lista del inventario no coincide con los archivos contenidos en el ZIP.")
            records = hash_records
            actual_total_bytes = 0
            for index, info in enumerate(infos, start=1):
                if info.filename == HASH_FILENAME:
                    continue
                try:
                    actual_hash, actual_size = _hash_member(archive, info)
                except (OSError, RuntimeError, ValueError, zipfile.BadZipFile, NotImplementedError):
                    errors.append(f"No se pudo leer de forma segura la entrada {index} del ZIP.")
                    continue
                actual_total_bytes += actual_size
                if actual_total_bytes > MAX_TOTAL_BYTES:
                    errors.append("El contenido descomprimido excede el límite total permitido.")
                    break
                record = records.get(info.filename)
                if record is None:
                    continue
                if actual_size != record["bytes"] or actual_hash != record["sha256"]:
                    errors.append(f"La entrada {index} no coincide con el tamaño o SHA-256 declarado.")
                else:
                    checked_files += 1

    return {"ok": not errors, "checked_files": checked_files, "errors": errors}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Verifica offline el inventario SHA-256 de un ZIP de exportación, sin extraer ni mostrar su contenido."
    )
    parser.add_argument("paquete", type=Path, help="Ruta del archivo ZIP que se desea verificar")
    args = parser.parse_args(argv)

    resultado = verificar_paquete(args.paquete)
    if resultado["ok"]:
        print(f"Paquete íntegro: {resultado['checked_files']} archivos verificados con SHA-256.")
        return 0
    print(f"Verificación fallida: {len(resultado['errors'])} control(es) no superados.")
    for error in resultado["errors"]:
        print(f"- {error}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
