$ErrorActionPreference = "Stop"

$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dllPath = Join-Path $PSScriptRoot "bin\Debug\net48\FitsPreviewHandler.dll"

if (-not (Test-Path $dllPath)) {
    Write-Host "Error: No se encontro $dllPath. Asegurate de compilar el proyecto primero." -ForegroundColor Red
    Pause
    exit
}

# Verificamos si somos admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Elevando privilegios para ejecutar RegAsm..." -ForegroundColor Yellow
    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell -ArgumentList $args -Verb RunAs
    exit
}

Write-Host "Registrando extensión COM..." -ForegroundColor Cyan
& $regasm /codebase $dllPath

Write-Host "`nRegistro completado. Por favor, reinicia el explorador de Windows o simplemente abre una ventana nueva y prueba el Panel de Vista Previa." -ForegroundColor Green
Pause
