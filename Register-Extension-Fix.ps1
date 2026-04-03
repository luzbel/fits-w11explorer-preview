$ErrorActionPreference = "Stop"

$dllPath = Join-Path $PSScriptRoot "bin\Debug\net48\FitsPreviewHandler.dll"
if (-not (Test-Path $dllPath)) {
    Write-Host "Error: No se encontro la DLL: $dllPath" -ForegroundColor Red
    Pause
    exit
}

# Verificamos si somos admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell -ArgumentList $args -Verb RunAs
    exit
}

$dllUrl = "file:///" + $dllPath.Replace('\', '/')
$guid = "{AF1C3D6A-81E9-4F5B-9A8C-2D9E71F04B3E}"
$previewGuid = "{8895b1c6-b41f-4c1c-a562-0d564250836f}"

# Instalar SharpShell en la GAC (requerido para que prevhost.exe encuentre la dependencia)
Write-Host "Instalando dependencias de SharpShell en el sistema..." -ForegroundColor Yellow
try {
    $sharpShellPath = Join-Path $PSScriptRoot "bin\Debug\net48\SharpShell.dll"
    [System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a") | Out-Null
    $publish = New-Object System.EnterpriseServices.Internal.Publish
    $publish.GacInstall($sharpShellPath)
} catch {
    Write-Host "Advertencia al instalar en GAC: $_" -ForegroundColor Red
}

# Crear el AppID en el sistema si Windows no lo tiene registrado (OBLIGATORIO para prevhost.exe)
$appIdPath = "Registry::HKEY_CLASSES_ROOT\AppID\{6d2b5079-6799-4d53-a192-ee068cbff056}"
if (-not (Test-Path $appIdPath)) { New-Item $appIdPath -Force | Out-Null }
Set-ItemProperty $appIdPath -Name "(default)" -Value "Preview Handler Surrogate Host"
Set-ItemProperty $appIdPath -Name "DllSurrogate" -Value "%SystemRoot%\system32\prevhost.exe" -Type ExpandString

# Registro del COM Server (equivalente a regasm /codebase pero infalible)
$clsidPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$guid"
if (-not (Test-Path $clsidPath)) { New-Item $clsidPath -Force | Out-Null }
Set-ItemProperty $clsidPath -Name "(default)" -Value "FitsPreviewHandler.FitsPreviewHandlerExtension"
Set-ItemProperty $clsidPath -Name "AppID" -Value "{6d2b5079-6799-4d53-a192-ee068cbff056}"
Set-ItemProperty $clsidPath -Name "DisableLowILProcessIsolation" -Value 1 -Type DWord

$inprocPath = "$clsidPath\InprocServer32"
if (-not (Test-Path $inprocPath)) { New-Item $inprocPath -Force | Out-Null }
Set-ItemProperty $inprocPath -Name "(default)" -Value "mscoree.dll"
Set-ItemProperty $inprocPath -Name "ThreadingModel" -Value "Apartment"
Set-ItemProperty $inprocPath -Name "Class" -Value "FitsPreviewHandler.FitsPreviewHandlerExtension"
Set-ItemProperty $inprocPath -Name "Assembly" -Value "FitsPreviewHandler, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty $inprocPath -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty $inprocPath -Name "CodeBase" -Value $dllUrl

$versionPath = "$inprocPath\1.0.0.0"
if (-not (Test-Path $versionPath)) { New-Item $versionPath -Force | Out-Null }
Set-ItemProperty $versionPath -Name "Class" -Value "FitsPreviewHandler.FitsPreviewHandlerExtension"
Set-ItemProperty $versionPath -Name "Assembly" -Value "FitsPreviewHandler, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty $versionPath -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty $versionPath -Name "CodeBase" -Value $dllUrl

# Registro de la extensión (.fits -> Preview Handler) en HKCR\.fits
$extPath = "Registry::HKEY_CLASSES_ROOT\.fits\shellex\$previewGuid"
if (-not (Test-Path $extPath)) { New-Item $extPath -Force | Out-Null }
Set-ItemProperty $extPath -Name "(default)" -Value $guid

# Registro FUERZA BRUTA en Fits.File (ProgID) y SystemFileAssociations para evitar sobreescritura por ASIFitsView
$progIdPath = "Registry::HKEY_CLASSES_ROOT\Fits.File\shellex\$previewGuid"
if (-not (Test-Path $progIdPath)) { New-Item $progIdPath -Force | Out-Null }
Set-ItemProperty $progIdPath -Name "(default)" -Value $guid

$sysAssocPath = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.fits\shellex\$previewGuid"
if (-not (Test-Path "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.fits\shellex")) { New-Item "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.fits\shellex" -Force | Out-Null }
if (-not (Test-Path $sysAssocPath)) { New-Item $sysAssocPath -Force | Out-Null }
Set-ItemProperty $sysAssocPath -Name "(default)" -Value $guid

# Registro en la lista de vistas previas autorizadas de Windows
$previewHandlersList = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers"
if (Test-Path $previewHandlersList) {
    Set-ItemProperty $previewHandlersList -Name $guid -Value "FITS File Preview Handler"
}

Write-Host "Registro completado con escritura en el registro manual y a prueba de fallos." -ForegroundColor Green
Write-Host "Abre el Explorador, activa el Panel de Vista Previa y revisa un archivo .FITS" -ForegroundColor Cyan
Pause
