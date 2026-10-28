<div align="center">

<img src="docs/capital_humano_banner.svg" width="100%" alt="Capital Humano — escala salarial interactiva, mínimo $3,210 CUP, ajuste en vivo 1 a 6 veces y exportación PDF y Excel" />

### Escala salarial interactiva · propuesta de plantilla · exportación PDF y Excel

Herramienta para revisar y ajustar la propuesta salarial de la plantilla, partiendo del salario mínimo de **$3,210.00 CUP** y respetando un máximo de **6× el salario original** de cada cargo.

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](TES/Salariales/Nueva%20Escala%203210/README.md)
[![Excel](https://img.shields.io/badge/Exportaci%C3%B3n-Excel%20%7C%20PDF-16866B?style=flat-square)](TES/Salariales/Nueva%20Escala%203210/)
[![HTTPS](https://img.shields.io/badge/HTTPS-Caddy%20%2B%20Let's%20Encrypt-1C7C54?style=flat-square)](docs/HTTPS.md)
[![CI](https://github.com/yhquintero/Capital_Humanos/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/yhquintero/Capital_Humanos/actions/workflows/ci.yml)
[![Estado](https://img.shields.io/badge/estado-en%20desarrollo-E6B85C?style=flat-square)](#-alcance-y-validación)

**[Abrir la guía de uso](TES/Salariales/Nueva%20Escala%203210/README.md) · [Ver el Excel de ejemplo](TES/Salariales/Nueva%20Escala%203210/Escala_Salarial_3210_Nueva.xlsx) · [Ver el PDF de ejemplo](TES/Salariales/Nueva%20Escala%203210/Escala_Salarial_3210_Nueva.pdf)**

</div>

---

## 🎯 Qué resuelve

- Lee los salarios y grupos de la plantilla heredada `propuesta de plantilla por la resolucion 87.xls`.
- Permite cambiar en **tiempo real** un multiplicador común de **1.00× a 6.00×**.
- Calcula y muestra para cada cargo el salario original, el multiplicador, el tope y el importe ajustado.
- Aplica un control explícito: **el salario ajustado no puede superar 6× el salario base del cargo en la plantilla fuente**.
- Mantiene anclado el **grupo I en $3,210.00**; los grupos sin propuesta de cargos conservan la escala oficial de 44 horas.
- Incluye un panel ejecutivo, auditoría filtrable por cargo, comparación de dos escenarios (por grupo, CSV/Excel), diagnóstico de fuentes/HTTPS y priorización de futuras mejoras.
- Exporta escala, detalle y trazabilidad SHA-256 a **Excel (.xlsx)** y **PDF A4 horizontal**; las exportaciones CSV incluyen un manifiesto JSON en el ZIP.

## 🚀 Inicio rápido

### Windows PowerShell

Se requiere **Python 3.10+ de 64 bits**: Streamlit depende de PyArrow y no existen wheels de PyArrow para Windows de 32 bits. Desde la raíz del repositorio:

```powershell
Set-Location -LiteralPath '.\TES\Salariales\Nueva Escala 3210'
# Elimina el entorno anterior solo si se creó con Python 32-bit:
if (Test-Path .\.venv) { Remove-Item -LiteralPath .\.venv -Recurse -Force }
.\instalar_windows.ps1
```

El instalador selecciona un intérprete x64, actualiza `pip`, instala las dependencias y verifica PyArrow. Si informa que no hay Python x64, instala Python 3.12 con `winget install Python.Python.3.12`, abre una consola nueva y ejecuta de nuevo `.\instalar_windows.ps1`. El error `Not supported on 32-bit Windows` no se soluciona actualizando pip: hay que usar Python de 64 bits.

### Linux / macOS

```bash
cd "TES/Salariales/Nueva Escala 3210"
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Abrir la aplicación

En Windows PowerShell, desde la carpeta del proyecto, ejecuta `python .\app_escala.py` o, de forma recomendada, `.\.venv\Scripts\python.exe -m streamlit run .\app_escala.py`. El primer comando se relanza ahora mediante Streamlit, evitando los avisos de “Missing ScriptRunContext” que aparecen al ejecutar una app Streamlit como un script Python normal.

Para HTTPS local, instala [mkcert](https://github.com/FiloSottile/mkcert) y ejecuta `run_https.ps1` en Windows o `./run_https.sh` en Linux/macOS. Se abre `https://localhost:8501`. La CA de mkcert se confía solo en los dispositivos donde se instale explícitamente; para LAN, usa la opción correspondiente y limita el Firewall a redes privadas.

**Autoarranque de Windows:** el script `control_app.ps1` administra el estado y puede configurar el inicio HTTPS al iniciar sesión:

```powershell
.\control_app.ps1 -Accion estado
.\control_app.ps1 -Accion instalar-autoinicio
```

Para una **URL pública con certificado reconocido**, dominio propio, renovación automática TLS y autenticación adicional por contraseña, consulta [HTTPS seguro con Caddy + Let's Encrypt](docs/HTTPS.md). Se requieren DNS apuntando al servidor y puertos 80/443 disponibles. No publiques directamente el puerto 8501.

El control **Multiplicador salarial** recalcula la propuesta en la interfaz; los botones generan Excel y PDF imprimibles.

Para regenerar los archivos de referencia originales desde la fuente:

```bash
python "TES/Salariales/Nueva Escala 3210/generar_escala.py"
```

## 🧭 Cómo usar el panel

1. En **Panel de control**, revisa los indicadores, comparación visual y controles de fuente.
2. Mueve el multiplicador de **1.00× a 6.00×**; los resultados y el detalle se recalculan de inmediato.
3. En **Escala salarial**, revisa los 32 grupos, sus importes y la metodología.
4. En **Comparar escenarios**, contrasta dos multiplicadores en paralelo y exporta la tabla por grupo; no la interpretes como costo total de nómina.
5. En **Plantilla por cargo**, filtra cargos y grupos y revisa las filas de origen.
6. En **Exportaciones**, descarga el libro Excel o el PDF imprimible del escenario activo.
7. En **Sistema y HTTPS**, comprueba fuentes/TLS y consulta el autoarranque HTTPS local de Windows; en **Mejoras futuras**, prioriza sugerencias que permanecen en el equipo.

El control se aplica a cada salario propuesto de la plantilla; el tope se verifica por fila, no solo sobre un promedio o el total. Los grupos que no aparecen en la plantilla mantienen el valor oficial vigente, para no inventar salarios fuente inexistentes.

## 📚 Fuentes y metodología

| Dato | Fuente |
|---|---|
| Salario mínimo mensual de $3,210.00 | Resolución 14/2026, GOC-2026-393-O60 |
| Escala oficial de 32 grupos, jornada de 44 horas | Resolución 15/2026, GOC-2026-394-O60 |
| Salario histórico | `TES/Salariales/Escala 3210_OK.xlsx` (Res. 29/2020) |
| Grupo, cargo y salario propuesto | `Nueva Estructura de Plantilla/Organigrama/propuesta de plantilla por la resolucion 87.xls` |

La plantilla incluye 183 registros distribuidos entre los grupos IV–XVIII. Cuando una misma categoría de grupo contiene importes distintos, el generador identifica y documenta las discrepancias; no modifica el archivo fuente.

## 📦 Archivos

```text
TES/Salariales/Nueva Escala 3210/
├── app_escala.py                     # interfaz web interactiva y exportadores
├── instalar_windows.ps1              # instala dependencias sobre Python x64
├── control_app.ps1                   # panel local y autoarranque HTTPS
├── run_https.ps1 / run_https.sh      # certificado local mkcert
├── generar_escala.py                 # extracción y generador de referencia
├── requirements.txt                  # dependencias
└── README.md                         # guía detallada
Dockerfile                            # contenedor Streamlit no-root
deploy/                               # Caddy + Let's Encrypt + acceso protegido
docs/HTTPS.md                         # despliegue TLS y gestión segura
docs/source_manifest.json             # huellas aprobadas de fuentes y código
tools/actualizar_manifiesto_fuentes.py # regenerar huellas tras revisión humana
docs/TRAZABILIDAD_FUENTES.md          # política y límites de SHA-256
```

## 🧰 Tecnologías

- **Python** — extracción, validaciones y cálculos reproducibles.
- **Streamlit** — interfaz interactiva con actualización inmediata.
- **xlrd** — lectura del libro antiguo `.xls`.
- **openpyxl** — generación de `.xlsx` con hojas, formato, filtros y configuración de impresión.
- **fpdf2** — informes PDF multipágina listos para imprimir.

## 🗺️ Hoja de ruta y contribuciones

Las mejoras se priorizan con beneficio, esfuerzo y criterios de aceptación explícitos; no se proclama una mejora como terminada sin pruebas. Revisa la [hoja de ruta](docs/ROADMAP.md), el [tablero de propuestas](docs/MEJORAS_PROPUESTAS.md), la [guía de trazabilidad SHA-256](docs/TRAZABILIDAD_FUENTES.md) y el [registro de cambios](CHANGELOG.md). GitHub ofrece formularios para reportar fallos y sugerir opciones sin adjuntar datos personales ni salarios reales.

Para ejecutar la suite localmente:

```bash
python -m pip install -r "TES/Salariales/Nueva Escala 3210/requirements.txt" -r requirements-dev.txt
python -m pytest -q
```

GitHub Actions repite las pruebas, analiza los scripts PowerShell en Windows y valida/compila el contenedor sin publicarlo.

## ✅ Alcance y validación

La herramienta facilita el análisis y la preparación de una propuesta; no sustituye la revisión de Recursos Humanos, Finanzas ni la validación jurídica que corresponda. Los importes calculados deben aprobarse por las autoridades competentes antes de su aplicación oficial.

---

<div align="center"><sub>Capital Humano · Escala salarial y plantilla · Cuba</sub></div>
