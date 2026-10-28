$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$scripts = Get-ChildItem -Path (Join-Path $root 'TES/Salariales/Nueva Escala 3210') -Filter '*.ps1' -File
if (-not $scripts) { throw 'No se encontraron scripts PowerShell.' }
foreach ($script in $scripts) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile(
        $script.FullName,
        [ref]$tokens,
        [ref]$errors
    ) | Out-Null
    if ($errors.Count -gt 0) {
        $messages = ($errors | ForEach-Object { $_.Message }) -join '; '
        throw "Error de sintaxis en $($script.Name): $messages"
    }
    Write-Host "Sintaxis PowerShell correcta: $($script.Name)"
}
