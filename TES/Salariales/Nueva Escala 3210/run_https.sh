#!/usr/bin/env bash
set -euo pipefail

PORT="${PORT:-8501}"
LAN_MODE=false
if [[ "${1:-}" == "--lan" ]]; then LAN_MODE=true; fi
HERE="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(cd -- "$HERE/../../.." && pwd)"
CERT_DIR="$ROOT/.certs"
CERT="$CERT_DIR/capital-humano-local.crt"
KEY="$CERT_DIR/capital-humano-local.key"
APP="$HERE/app_escala.py"

command -v mkcert >/dev/null 2>&1 || { echo "Instala mkcert desde https://github.com/FiloSottile/mkcert e inténtalo de nuevo." >&2; exit 1; }
if [[ -x "$HERE/.venv/bin/python" ]]; then
  PYTHON="$HERE/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON="$(command -v python3)"
else
  echo "Se requiere Python 3.10 o posterior." >&2
  exit 1
fi
"$PYTHON" -c 'import struct,sys,streamlit,pyarrow; assert struct.calcsize("P")*8 == 64 and sys.version_info >= (3,10)' || {
  echo "Usa un Python 3.10+ de 64 bits e instala requirements.txt antes de iniciar." >&2
  exit 1
}
mkdir -p "$CERT_DIR"

# El paso -install agrega la CA de mkcert al almacén de confianza de ESTE equipo.
mkcert -install
HOST="$(hostname -s 2>/dev/null || hostname)"
NAMES=(localhost 127.0.0.1 ::1 "$HOST")
if $LAN_MODE; then
  while IFS= read -r ip; do
    [[ -n "$ip" && "$ip" != 127.* && "$ip" != 169.254.* ]] && NAMES+=("$ip")
  done < <(python3 - <<'PY'
import socket
for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
    print(info[4][0])
PY
)
  UNIQUE_NAMES=()
  while IFS= read -r name; do UNIQUE_NAMES+=("$name"); done < <(printf '%s\n' "${NAMES[@]}" | awk '!seen[$0]++')
  NAMES=("${UNIQUE_NAMES[@]}")
fi

mkcert -cert-file "$CERT" -key-file "$KEY" "${NAMES[@]}"
if $LAN_MODE; then
  LISTEN=0.0.0.0
  echo "Modo LAN: permite el puerto $PORT únicamente en redes privadas/firewall."
  echo "Los demás dispositivos deben confiar en la CA de mkcert. Comparte solo rootCA.pem; nunca rootCA-key.pem."
else
  LISTEN=127.0.0.1
fi
exec "$PYTHON" -m streamlit run "$APP" \
  --server.address "$LISTEN" \
  --server.port "$PORT" \
  "--server.sslCertFile=$CERT" \
  "--server.sslKeyFile=$KEY" \
  --browser.gatherUsageStats=false
