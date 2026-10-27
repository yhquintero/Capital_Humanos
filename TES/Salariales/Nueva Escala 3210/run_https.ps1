[CmdletBinding()]
param(
    [int]$Puerto = 8501,
    [switch]$EnRedLocal
)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$certDir = Join-Path $raiz '.certs'
$cert = Join-Path $certDir 'capital-humano-local.crt'
$key = Join-Path $certDir 'capital-humano-local.key'
$app = Join-Path $PSScriptRoot 'app_escala.py'

$mkcert = Get-Command mkcert -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $mkcert) {
    Write-Error 'Falta mkcert. Instálalo con: winget install FiloSottile.mkcert. Luego abre una nueva consola.'
    exit 1
}
$pythonVenv = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (Test-Path $pythonVenv) {
    $pythonExe = $pythonVenv
} else {
    $python = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) {
        Write-Error 'No se encontró Python. Ejecuta primero instalar_windows.ps1.'
        exit 1
    }
    $pythonExe = $python.Source
}
& $pythonExe -c "import struct,sys,streamlit,pyarrow; assert struct.calcsize('P')*8 == 64 and sys.version_info >= (3,10)" *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Error 'El entorno no es Python 3.10+ de 64 bits o faltan dependencias. Ejecuta primero .\instalar_windows.ps1.'
    exit 1
}

New-Item -ItemType Directory -Path $certDir -Force | Out-Null
Write-Host '[1/3] Instalando/validando la CA local de mkcert en este equipo…' -ForegroundColor Cyan
& $mkcert.Source -install
if ($LASTEXITCODE -ne 0) { throw 'mkcert no pudo instalar su autoridad local de confianza.' }

$names = @('localhost', '127.0.0.1', '::1', $env:COMPUTERNAME)
if ($EnRedLocal) {
    $ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254.*' } |
        Select-Object -ExpandProperty IPAddress -Unique)
    $names += $ips
}
$names = @($names | Where-Object { $_ } | Select-Object -Unique)
Write-Host "[2/3] Creando certificado local con SAN: $($names -join ', ')" -ForegroundColor Cyan
& $mkcert.Source -cert-file $cert -key-file $key @names
if ($LASTEXITCODE -ne 0) { throw 'No se pudo crear el certificado TLS local.' }

$listen = if ($EnRedLocal) { '0.0.0.0' } else { '127.0.0.1' }
Write-Host "[3/3] Iniciando la aplicación en https://localhost:$Puerto" -ForegroundColor Green
if ($EnRedLocal) {
    Write-Warning 'Modo LAN: el puerto queda accesible en interfaces de red. Usa solo redes privadas y permite el puerto en Firewall únicamente para perfil Privado.'
    foreach ($ip in $ips) { Write-Host "  URL de esta red: https://${ip}:$Puerto" }
    Write-Warning 'Los demás dispositivos deben confiar en la CA pública (rootCA.pem) de mkcert; nunca copies rootCA-key.pem.'
}
& $pythonExe -m streamlit run $app --server.address $listen --server.port $Puerto "--server.sslCertFile=$cert" "--server.sslKeyFile=$key" --server.headless=true --browser.gatherUsageStats=false
