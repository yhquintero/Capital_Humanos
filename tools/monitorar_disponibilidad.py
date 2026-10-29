#!/usr/bin/env python3
"""Sonda externa de disponibilidad, sin leer ni registrar contenido de la app."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import time
import urllib.error
import urllib.request
from urllib.parse import urlsplit

HEALTH_PATH = "/healthz"
USER_AGENT = "Capital-Humano-Availability-Monitor/1.0"


class _SinRedirecciones(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        return None


def validar_url(url: str) -> str:
    """Permite HTTPS público o HTTP loopback solo para pruebas locales."""
    parts = urlsplit(url)
    hostname = (parts.hostname or "").casefold()
    loopback_hosts = {"localhost", "127.0.0.1", "::1"}
    try:
        _ = parts.port
    except ValueError as exc:
        raise ValueError("La URL contiene un puerto inválido.") from exc
    if any(character.isspace() for character in url):
        raise ValueError("La URL no puede contener espacios ni saltos de línea.")
    if parts.username or parts.password:
        raise ValueError("No incluyas credenciales en la URL de monitorización.")
    if parts.query or parts.fragment:
        raise ValueError("La URL no puede incluir parámetros ni fragmentos.")
    if parts.path != HEALTH_PATH:
        raise ValueError(f"La ruta debe ser exactamente {HEALTH_PATH}.")
    if parts.scheme != "https" and not (parts.scheme == "http" and hostname in loopback_hosts):
        raise ValueError("Usa HTTPS público; HTTP solo se permite para pruebas en loopback.")
    if not hostname:
        raise ValueError("La URL debe incluir un host.")
    return url


def comprobar_endpoint(url: str, timeout: float = 5.0) -> dict:
    """Comprueba HTTP 200 + cuerpo `ok`; nunca devuelve ni registra el cuerpo."""
    validar_url(url)
    request = urllib.request.Request(
        url,
        method="GET",
        headers={"Accept": "text/plain", "User-Agent": USER_AGENT},
    )
    started = time.monotonic()
    http_status = None
    try:
        opener = urllib.request.build_opener(_SinRedirecciones)
        with opener.open(request, timeout=timeout) as response:
            http_status = response.status
            body_ok = response.read(16).strip().lower() == b"ok"
            is_up = http_status == 200 and body_ok
    except urllib.error.HTTPError as exc:
        http_status = exc.code
        is_up = False
    except (urllib.error.URLError, TimeoutError, OSError):
        is_up = False
    duration_ms = round((time.monotonic() - started) * 1000, 1)
    return {
        "service": "capital-humano",
        "checked_at_utc": dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat(),
        "status": "UP" if is_up else "DOWN",
        "http_status": http_status,
        "duration_ms": duration_ms,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True, help="URL HTTPS pública terminada en /healthz")
    parser.add_argument("--timeout", type=float, default=5.0, help="Tiempo máximo de espera (segundos)")
    args = parser.parse_args(argv)
    if not 0.1 <= args.timeout <= 30:
        parser.error("--timeout debe estar entre 0.1 y 30 segundos")
    try:
        result = comprobar_endpoint(args.url, args.timeout)
    except ValueError as exc:
        parser.error(str(exc))
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result["status"] == "UP" else 1


if __name__ == "__main__":
    sys.exit(main())
