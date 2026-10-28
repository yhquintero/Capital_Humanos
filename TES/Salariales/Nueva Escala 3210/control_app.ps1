[CmdletBinding()]
param(
    [ValidateSet('iniciar','detener','estado','abrir','instalar-autoinicio','quitar-autoinicio')]
    [string]$Accion = 'estado',
    [ValidateRange(1024,65535)]
    [int]$Puerto = 8501
)

$ErrorActionPreference = 'Stop'
$NombreTarea = 'Capital Humano - Escala Salarial HTTPS'
$RunHttps = Join-Path $PSScriptRoot 'run_https.ps1'
$App = Join-Path $PSScriptRoot 'app_escala.py'
$Url = "https://localhost:$Puerto"
$DatosControl = Join-Path $env:LOCALAPPDATA 'CapitalHumano'
$StdoutLog = Join-Path $DatosControl 'inicio-https.log'
$StderrLog = Join-Path $DatosControl 'inicio-https-error.log'

function Test-WebLocal {
    try {
        $r = Invoke-WebRequest -Uri "$Url/_stcore/health" -SkipCertificateCheck -TimeoutSec 3
        return ($r.StatusCode -eq 200 -and $r.Content.Trim() -eq 'ok')
    } catch { return $false }
}

function Start-WebLocal {
    if (Test-WebLocal) {
        Write-Host "La aplicación ya está funcionando en $Url" -ForegroundColor Green
        return
    }
    if (-not (Test-Path $RunHttps)) { throw "No se encontró el lanzador HTTPS: $RunHttps" }
    if (-not (Get-Command mkcert -ErrorAction SilentlyContinue)) {
        throw 'Falta mkcert. Instálalo con: winget install FiloSottile.mkcert'
    }
    New-Item -ItemType Directory -Path $DatosControl -Force | Out-Null
    $PowerShellExe = (Get-Process -Id $PID).Path
    $Argumentos = "-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$RunHttps`" -Puerto $Puerto"
    Start-Process -FilePath $PowerShellExe -ArgumentList $Argumentos -WindowStyle Hidden `
        -RedirectStandardOutput $StdoutLog -RedirectStandardError $StderrLog | Out-Null

    Write-Host 'Iniciando el servidor TLS; el primer arranque puede tardar un poco…' -ForegroundColor Cyan
    for ($i = 0; $i -lt 35; $i++) {
        Start-Sleep -Seconds 1
        if (Test-WebLocal) {
            Write-Host "HTTPS activo: $Url" -ForegroundColor Green
            return
        }
    }
    $detalle = if (Test-Path $StderrLog) { (Get-Content $StderrLog -Tail 10 -ErrorAction SilentlyContinue) -join "`n" } else { '' }
    throw "La app no respondió por HTTPS dentro del tiempo esperado. Revisa $StderrLog. $detalle"
}

function Stop-WebLocal {
    $basePath = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
    $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $_.CommandLine -and $_.CommandLine.IndexOf('app_escala.py', [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
        $_.CommandLine.IndexOf($basePath, [StringComparison]::OrdinalIgnoreCase) -ge 0
    }
    if (-not $processes) {
        Write-Host 'No se encontró un proceso de esta aplicación en ejecución.' -ForegroundColor Yellow
        return
    }
    foreach ($process in $processes) {
        & taskkill.exe /PID $process.ProcessId /T /F *> $null
    }
    Write-Host 'Aplicación detenida.' -ForegroundColor Green
}

switch ($Accion) {
    'iniciar' { Start-WebLocal }
    'detener' { Stop-WebLocal }
    'abrir' {
        if (-not (Test-WebLocal)) { Start-WebLocal }
        Start-Process $Url
    }
    'estado' {
        $running = Test-WebLocal
        $task = Get-ScheduledTask -TaskName $NombreTarea -ErrorAction SilentlyContinue
        Write-Host 'CAPITAL HUMANO · PANEL DE CONTROL' -ForegroundColor Cyan
        Write-Host ("Web HTTPS:       " + $(if ($running) { "ACTIVA · $Url" } else { 'DETENIDA' }))
        Write-Host ("Autoarranque:    " + $(if ($task) { 'CONFIGURADO al iniciar sesión' } else { 'NO configurado' }))
        Write-Host "Puerto local:    $Puerto"
        if (-not $running) { Write-Host "Para iniciar:    .\control_app.ps1 -Accion iniciar" }
    }
    'instalar-autoinicio' {
        # El primer inicio instala la CA local de mkcert y crea el certificado SAN.
        Start-WebLocal
        $PowerShellExe = (Get-Process -Id $PID).Path
        $Argumentos = "-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Accion iniciar -Puerto $Puerto"
        $action = New-ScheduledTaskAction -Execute $PowerShellExe -Argument $Argumentos
        $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 2)
        $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
        Register-ScheduledTask -TaskName $NombreTarea -Action $action -Trigger $trigger `
            -Settings $settings -Principal $principal -Description 'Inicia Capital Humano por HTTPS al iniciar sesión.' -Force | Out-Null
        Write-Host "Autoarranque HTTPS instalado para $user. Se iniciará al iniciar sesión en Windows." -ForegroundColor Green
    }
    'quitar-autoinicio' {
        Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host 'Autoarranque quitado. La aplicación actual no se detuvo.' -ForegroundColor Green
    }
}
