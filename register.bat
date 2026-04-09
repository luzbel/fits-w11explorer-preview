@echo off
:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator privileges detected.
) else (
    echo [ERROR] This script must be run as Administrator.
    pause
    exit /b 1
)

:: Find regasm.exe. 
:: Note: The v4.0.30319 directory is a fixed path for ALL .NET 4.0-4.8.x versions.
set REGASM="%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
if not exist %REGASM% (
    set REGASM="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
)

if not exist %REGASM% (
    echo [ERROR] .NET Framework 4.0 or higher was not found on this system.
    pause
    exit /b 1
)

echo Registering FitsPreviewHandler.dll using: %REGASM%
%REGASM% /codebase "%~dp0FitsPreviewHandler.dll"
pause
