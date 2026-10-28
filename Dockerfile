FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    HOME=/home/streamlit \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

RUN apt-get update \
    && apt-get install -y --no-install-recommends fonts-dejavu-core ca-certificates \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --create-home --shell /usr/sbin/nologin streamlit

WORKDIR /app
COPY ["TES/Salariales/Nueva Escala 3210/requirements.txt", "/app/requirements.txt"]
RUN python -m pip install --no-cache-dir -r /app/requirements.txt

COPY ["TES/Salariales/Nueva Escala 3210/app_escala.py", "/app/TES/Salariales/Nueva Escala 3210/app_escala.py"]
COPY ["TES/Salariales/Nueva Escala 3210/generar_escala.py", "/app/TES/Salariales/Nueva Escala 3210/generar_escala.py"]
COPY ["TES/Salariales/Escala 3210_OK.xlsx", "/app/TES/Salariales/Escala 3210_OK.xlsx"]
COPY ["TES/Salariales/goc-2026-o60.pdf", "/app/TES/Salariales/goc-2026-o60.pdf"]
COPY ["Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls", "/app/Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls"]
COPY ["docs/source_manifest.json", "/app/docs/source_manifest.json"]
COPY ["tools/actualizar_manifiesto_fuentes.py", "/app/tools/actualizar_manifiesto_fuentes.py"]

RUN chown -R streamlit:streamlit /app /home/streamlit
USER streamlit
EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=3)"]

CMD ["python", "-m", "streamlit", "run", "/app/TES/Salariales/Nueva Escala 3210/app_escala.py", "--server.address=0.0.0.0", "--server.port=8501", "--server.headless=true", "--browser.gatherUsageStats=false"]
