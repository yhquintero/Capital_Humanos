# Escala Salarial 3210 · interfaz interactiva

Herramienta Python para recalcular la escala salarial y preparar una propuesta imprimible a partir de `propuesta de plantilla por la resolucion 87.xls`.

## Funciones

- Slider en vivo entre **1× y 6×**, en pasos de 0.05.
- Aplicación del factor por cargo a cada salario extraído de la plantilla.
- Validación estricta por fila: `salario ajustado ≤ salario original × 6`.
- Mínimo del grupo I fijado en **$3,210.00 CUP** (Res. 14/2026).
- Grupos con propuesta recalculados; los demás conservan la escala oficial Res. 15/2026 (44 h).
- Comparación de dos escenarios simultáneos, por grupo, con exportación CSV/Excel.
- Exportación de la escala activa a **Excel `.xlsx`** y **PDF A4 horizontal** con página de trazabilidad.
- SHA-256 y versión de app en los informes; las descargas CSV incluyen el manifiesto JSON dentro del ZIP.
- Validación de integridad de plantilla, escala histórica, Gaceta Oficial y módulos de cálculo.
- Trazabilidad por hoja y fila de origen del libro `.xls`.
- Informe de calidad en **Plantilla por cargo**: salarios ausentes/no válidos, cargos sin descripción, posibles duplicados y discrepancias por grupo con hoja/fila; exportable a CSV + manifiesto y sin correcciones automáticas.
- Pruebas automáticas de reglas, fuentes y navegación Streamlit en GitHub Actions.

## Windows PowerShell · instalación recomendada

**Requisito: Python 3.10 o posterior, de 64 bits.** Streamlit usa PyArrow y PyArrow no tiene paquetes binarios para Python de 32 bits en Windows. El error `RuntimeError: Not supported on 32-bit Windows` significa que el intérprete usado creó un entorno virtual de 32 bits; actualizar `pip` no lo corrige.

En PowerShell, pon entre comillas la ruta que contiene el espacio. Por ejemplo, desde la ubicación mostrada en tu consola:

```powershell
Set-Location -LiteralPath 'D:\Proyectos\Github\Capital_Humanos\TES\Salariales\Nueva Escala 3210'
```

Si ya estás en la raíz del repositorio, también sirve la ruta relativa:

```powershell
Set-Location -LiteralPath '.\TES\Salariales\Nueva Escala 3210'
```

Si ya habías creado `.venv` con Python de 32 bits, elimínalo primero (esto elimina solo el entorno virtual, no los archivos de la escala):

```powershell
Remove-Item -LiteralPath .\.venv -Recurse -Force
```

Instala Python 3.12 de 64 bits si aún no está instalado y vuelve a abrir PowerShell:

```powershell
winget install Python.Python.3.12
```

Después ejecuta el instalador del proyecto. Este selecciona automáticamente Python x64 3.10+, actualiza pip, instala los paquetes y verifica Streamlit/PyArrow:

```powershell
.\instalar_windows.ps1
```

Si PowerShell bloquea la ejecución del script, permite scripts solo en la ventana actual:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\instalar_windows.ps1
```

El lanzador `py` puede listar los intérpretes instalados con `py -0p`. Para comprobar el entorno creado:

```powershell
.\.venv\Scripts\python.exe -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8, 'bits')"
```

Debe indicar **64 bits** y Python **3.10+**.

## Iniciar la aplicación

Forma recomendada desde PowerShell, en esta carpeta:

```powershell
.\.venv\Scripts\python.exe -m streamlit run .\app_escala.py
```

Con el entorno virtual activado también funciona el comando corto:

```powershell
python .\app_escala.py
```

El script detecta que no se está ejecutando dentro de Streamlit y se relanza con su servidor web. Ya no debe mostrar la repetición de `Missing ScriptRunContext` ni el aviso de “run it with streamlit run”. Alternativamente, el modo HTTPS local se inicia con el script descrito abajo.

## Panel de control y autoarranque HTTPS en Windows

Instala [mkcert](https://github.com/FiloSottile/mkcert) una vez y abre PowerShell en esta carpeta. El panel de consola permite iniciar, comprobar, abrir y detener la app, además de programarla para que arranque por HTTPS al iniciar sesión:

```powershell
.\control_app.ps1 -Accion estado
.\control_app.ps1 -Accion iniciar
.\control_app.ps1 -Accion abrir
.\control_app.ps1 -Accion instalar-autoinicio
```

El primer inicio crea/confía el certificado local y valida `https://localhost:8501/_stcore/health`. El autoarranque se ejecuta bajo el mismo usuario de Windows al iniciar sesión; no se recomienda ejecutarlo como servicio de sistema porque la CA mkcert pertenece al perfil del usuario. Para quitarlo: ` .\control_app.ps1 -Accion quitar-autoinicio `. Para uso en LAN, configura el Firewall en perfil Privado y confía en la raíz mkcert desde cada dispositivo.

La app web incluye un **Panel de Control**, vistas de escala y comparación A/B, auditoría filtrable, exportaciones, diagnóstico del entorno y un formulario local para priorizar mejoras futuras. La comparación muestra importes por grupo; no estima el costo total de nómina porque no dispone de cantidades autorizadas de plazas.

Para obtener un certificado público validado por los navegadores y renovación automática, se necesita un dominio propio. Sigue [`docs/HTTPS.md`](../../../docs/HTTPS.md) para el despliegue Caddy + Let's Encrypt con autenticación adicional.

## Linux / macOS

Desde la carpeta del proyecto:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
./run_https.sh
```

## Salidas y fuentes

- `Escala_Salarial_3210_Nueva.xlsx` y `.pdf`: ejemplos generados por `generar_escala.py`.
- Las descargas Excel/PDF y CSV de detalle anteponen un Id. correlativo a la hoja de origen, conservado al filtrar para contar y rastrear cada ítem; los valores faltantes aparecen en blanco, no como `None`.
- Las descargas Excel/PDF de la app contienen la escala y los cargos con el multiplicador seleccionado; el informe de aprobación separado contiene un resumen por grupo, metodología, hashes y campos para revisión humana.
- El informe de aprobación se marca como borrador hasta completar manualmente las revisiones y firmas; no contiene una aprobación digital ni estima el gasto total de nómina.
- Res. 14/2026, GOC-2026-393-O60: salario mínimo mensual de $3,210.00.
- Res. 15/2026, GOC-2026-394-O60: escala oficial de 32 grupos, jornada de 44 horas.
- `Escala 3210_OK.xlsx`: escala histórica Res. 29/2020.
- `propuesta de plantilla por la resolucion 87.xls`: cargos y salarios de origen.

El tope se verifica por cargo: `salario ajustado ≤ salario original × 6`. Los cargos sin salario propuesto conservan el salario oficial y el grupo I permanece anclado en $3,210.00. Dos discrepancias de origen aparecen en los grupos XII y XV; la interfaz mantiene el detalle por fila para su revisión. Validar cualquier propuesta antes de aprobarla.

Para entender o actualizar las huellas SHA-256 de las fuentes, sigue la [guía de trazabilidad](../../../docs/TRAZABILIDAD_FUENTES.md). No regeneres el manifiesto para silenciar una discrepancia que no haya sido revisada.
