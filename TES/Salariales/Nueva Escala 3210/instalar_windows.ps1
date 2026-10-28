$ErrorActionPreference = 'Stop'
$proyecto = $PSScriptRoot
$venv = Join-Path $proyecto '.venv'
$pythonVenv = Join-Path $venv 'Scripts\python.exe'
$requirements = Join-Path $proyecto 'requirements.txt'

$launcher = Get-Command py.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $launcher) {
    Write-Error 'No se encontró el Python Launcher (py.exe). Instala Python 3.12 de 64 bits desde python.org o winget install Python.Python.3.12 y vuelve a abrir PowerShell.'
    exit 1
}

# Streamlit depende de PyArrow, que no ofrece wheels para Windows de 32 bits.
# Buscar primero los runtimes de 64 bits más recientes compatibles.
$runtime = $null
foreach ($candidate in @('-3.13-64', '-3.12-64', '-3.11-64', '-3.10-64')) {
    & $launcher.Source $candidate -c "import sys,struct; assert struct.calcsize('P')*8 == 64 and sys.version_info >= (3,10)" *> $null
    if ($LASTEXITCODE -eq 0) { $runtime = $candidate; break }
}
if (-not $runtime) {
    Write-Error 'No hay Python 3.10+ de 64 bits instalado. Ejecuta: winget install Python.Python.3.12 ; luego abre una nueva consola y vuelve a correr este script.'
    exit 1
}
Write-Host "Usando runtime: py $runtime" -ForegroundColor Cyan

if (Test-Path $pythonVenv) {
    & $pythonVenv -c "import struct,sys; assert struct.calcsize('P')*8 == 64 and sys.version_info >= (3,10)" *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Error 'El .venv existente no es Python 3.10+ de 64 bits. Elimínalo (solo es el entorno virtual, no tus documentos): Remove-Item -LiteralPath .\.venv -Recurse -Force ; después vuelve a correr este script.'
        exit 1
    }
    Write-Host 'Usando el entorno virtual existente de 64 bits.' -ForegroundColor Green
} else {
    & $launcher.Source $runtime -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw 'No se pudo crear el entorno virtual.' }
}

& $pythonVenv -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw 'No se pudo actualizar pip.' }
& $pythonVenv -m pip install -r $requirements
if ($LASTEXITCODE -ne 0) { throw 'Falló la instalación. Verifica que Python sea de 64 bits y que haya conexión a PyPI.' }
& $pythonVenv -c "import struct,sys,pandas,pyarrow,streamlit; print('Python',sys.version.split()[0], '| bits:',struct.calcsize('P')*8,'| Streamlit',streamlit.__version__)"
if ($LASTEXITCODE -ne 0) { throw 'La comprobación de dependencias falló.' }
Write-Host ''
Write-Host 'Instalación lista. Panel de control HTTPS (después de instalar mkcert):' -ForegroundColor Green
Write-Host '.\control_app.ps1 -Accion estado'
Write-Host '.\control_app.ps1 -Accion abrir'
Write-Host 'Para habilitar el arranque HTTPS al iniciar sesión:'
Write-Host '.\control_app.ps1 -Accion instalar-autoinicio'
Write-Host 'Arranque HTTPS manual alternativo:'
Write-Host '.\run_https.ps1'
