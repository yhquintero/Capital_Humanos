from __future__ import annotations

import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tools.monitorar_disponibilidad import comprobar_endpoint, main, validar_url  # noqa: E402


class _HealthHandler(BaseHTTPRequestHandler):
    status_code = 200
    response_body = b"ok"
    redirect = False

    def do_GET(self):
        if self.path != "/healthz":
            self.send_response(404)
            self.end_headers()
            return
        if self.redirect:
            self.send_response(302)
            self.send_header("Location", "/healthz")
            self.end_headers()
            return
        self.send_response(self.status_code)
        self.end_headers()
        self.wfile.write(self.response_body)

    def log_message(self, _format, *_args):
        pass


@pytest.fixture
def health_server():
    server = ThreadingHTTPServer(("127.0.0.1", 0), _HealthHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/healthz"
    finally:
        server.shutdown()
        thread.join(timeout=2)
        server.server_close()


def test_monitor_reports_only_status_for_healthy_endpoint(health_server):
    _HealthHandler.status_code = 200
    _HealthHandler.response_body = b"ok"
    _HealthHandler.redirect = False

    result = comprobar_endpoint(health_server, timeout=1)

    assert result["status"] == "UP"
    assert result["http_status"] == 200
    assert result["duration_ms"] >= 0
    assert "body" not in result and "url" not in result


def test_monitor_reports_synthetic_outage_without_exposing_response_body(health_server, capsys):
    _HealthHandler.status_code = 503
    _HealthHandler.response_body = b"private salary response must not be logged"
    _HealthHandler.redirect = False

    exit_code = main(["--url", health_server, "--timeout", "1"])
    output = capsys.readouterr().out
    result = json.loads(output)

    assert exit_code == 1
    assert result["status"] == "DOWN"
    assert result["http_status"] == 503
    assert "private salary" not in output
    assert "url" not in result


def test_monitor_does_not_follow_redirects(health_server):
    _HealthHandler.status_code = 200
    _HealthHandler.response_body = b"ok"
    _HealthHandler.redirect = True
    try:
        result = comprobar_endpoint(health_server, timeout=1)
    finally:
        _HealthHandler.redirect = False
    assert result["status"] == "DOWN"
    assert result["http_status"] == 302


@pytest.mark.parametrize(
    "url",
    [
        "http://example.invalid/healthz",
        "https://user:password@example.com/healthz",
        "https://example.com/healthz?token=secret",
        "https://example.com/other",
        "https://example.com/healthz#fragment",
    ],
)
def test_monitor_rejects_insecure_or_ambiguous_urls(url):
    with pytest.raises(ValueError):
        validar_url(url)


def test_caddy_exposes_only_the_minimal_health_route_without_authentication():
    caddyfile = (ROOT / "deploy" / "Caddyfile").read_text(encoding="utf-8")
    health = caddyfile.index("handle @health")
    fallback = caddyfile.index("handle {", health)
    auth = caddyfile.index("basic_auth {", fallback)
    rewrite = caddyfile.index("rewrite * /_stcore/health", health)

    assert health < rewrite < fallback < auth
    assert "header Cache-Control \"no-store\"" in caddyfile
    assert "reverse_proxy app:8501" in caddyfile
