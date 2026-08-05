@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
:: =========================================================
:: Reactivar OneDrive - CityPC
:: Restaura OneDrive a su funcionamiento normal
:: =========================================================

set "LOCAL_VER=2"
set "CITYPC_TOOL_ID=onedrive"
set "CITYPC_RELEASE_TAG=preparacion-v49"
set "CITYPC_UPDATER_SHA256=02b05836503e43fb61a7ac1b60267a79374ee77e7653c13765c0cf0da798ed38"

:: Verificar permisos de administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Este script requiere permisos de administrador.
    echo Haz clic derecho y selecciona "Ejecutar como administrador".
    pause
    exit /b 1
)

echo   Verificando paquete de preparacion...
call :citypc_update "%~1" "%~2"
if "!CITYPC_UPDATE_RC!"=="20" exit
if "!CITYPC_UPDATE_RC!"=="30" exit
if "!CITYPC_UPDATE_RC!"=="31" exit

echo =========================================================
echo   Reactivar OneDrive - CityPC
echo =========================================================
echo.

:: 1. Eliminar la politica que deshabilita la sincronizacion
echo   Eliminando politica de bloqueo de sincronizacion...
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v DisableFileSyncNGSC /f >nul 2>&1
if %errorlevel% equ 0 (
    echo   [OK] Politica de bloqueo eliminada
) else (
    echo   [INFO] No existia politica de bloqueo
)

:: 2. Restaurar OneDrive en el inicio automatico del usuario actual
echo   Restaurando OneDrive en el inicio automatico...
set "ONEDRIVE_PATH=%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
if exist "%ONEDRIVE_PATH%" (
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "OneDrive" /t REG_SZ /d "\"%ONEDRIVE_PATH%\" /background" /f >nul 2>&1
    if %errorlevel% equ 0 (
        echo   [OK] OneDrive restaurado en inicio automatico
    ) else (
        echo   [ERROR] No se pudo restaurar en inicio automatico
    )
) else (
    echo   [AVISO] No se encontro OneDrive.exe en la ruta esperada.
    echo           Puede que OneDrive no este instalado.
    echo           Ruta buscada: %ONEDRIVE_PATH%
)

:: 3. Iniciar OneDrive
echo   Iniciando OneDrive...
if exist "%ONEDRIVE_PATH%" (
    start "" "%ONEDRIVE_PATH%"
    echo   [OK] OneDrive iniciado
) else (
    echo   [AVISO] No se puede iniciar OneDrive, el ejecutable no existe.
)

echo.
echo =========================================================
echo   OneDrive ha sido reactivado correctamente.
echo   Se iniciara automaticamente con Windows.
echo =========================================================
echo.
pause
exit

:: ACTUALIZADOR COMUN CITYPC - BUNDLE INMUTABLE Y TRANSACCIONAL
:: =========================================================
:citypc_update
set "CITYPC_UPDATE_RC=0"
set "CITYPC_UPDATE_RESUME="
if /I "%~1"=="--citypc-update-resume" set "CITYPC_UPDATE_RESUME=%~2"
set "CITYPC_UPDATER_PATH=%~dp0CityPC_Updater.ps1"
set "CITYPC_UPDATER_TMP=%TEMP%\CityPC_Updater_%RANDOM%_%RANDOM%.new"
set "CITYPC_UPDATER_BAK=%~dp0CityPC_Updater.bootstrap.bak"
set "CITYPC_UPDATER_OK=0"
if exist "!CITYPC_UPDATER_PATH!" for /f "usebackq delims=" %%H in (`powershell -NoProfile -Command "try{(Get-FileHash -LiteralPath $env:CITYPC_UPDATER_PATH -Algorithm SHA256).Hash.ToLowerInvariant()}catch{}"`) do if /I "%%H"=="!CITYPC_UPDATER_SHA256!" set "CITYPC_UPDATER_OK=1"
if "!CITYPC_UPDATER_OK!"=="0" (
    echo   [INFO] Recuperando actualizador verificado...
    del /F /Q "!CITYPC_UPDATER_TMP!" >nul 2>&1
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;$u=@('https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/!CITYPC_RELEASE_TAG!/CityPC_Updater.ps1','https://github.com/rodrigofufer/CityPC-Installer/raw/refs/tags/!CITYPC_RELEASE_TAG!/CityPC_Updater.ps1');$ok=$false;foreach($x in $u){1..2|ForEach-Object{if(-not $ok){try{Invoke-WebRequest -Uri ($x+'?bootstrap='+[DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) -OutFile $env:CITYPC_UPDATER_TMP -UseBasicParsing -TimeoutSec 12 -MaximumRedirection 3;if((Get-FileHash -LiteralPath $env:CITYPC_UPDATER_TMP -Algorithm SHA256).Hash.ToLowerInvariant() -eq $env:CITYPC_UPDATER_SHA256){$ok=$true}}catch{};if(-not $ok){Start-Sleep -Seconds 1}}}};if(-not $ok){exit 1}" >nul 2>&1
    if exist "!CITYPC_UPDATER_TMP!" (
        if exist "!CITYPC_UPDATER_PATH!" copy /Y "!CITYPC_UPDATER_PATH!" "!CITYPC_UPDATER_BAK!" >nul 2>&1
        move /Y "!CITYPC_UPDATER_TMP!" "!CITYPC_UPDATER_PATH!" >nul 2>&1
    )
    if exist "!CITYPC_UPDATER_PATH!" for /f "usebackq delims=" %%H in (`powershell -NoProfile -Command "try{(Get-FileHash -LiteralPath $env:CITYPC_UPDATER_PATH -Algorithm SHA256).Hash.ToLowerInvariant()}catch{}"`) do if /I "%%H"=="!CITYPC_UPDATER_SHA256!" set "CITYPC_UPDATER_OK=1"
    if "!CITYPC_UPDATER_OK!"=="0" (
        if exist "!CITYPC_UPDATER_BAK!" move /Y "!CITYPC_UPDATER_BAK!" "!CITYPC_UPDATER_PATH!" >nul 2>&1
        if defined CITYPC_UPDATE_RESUME (
            echo   [ERROR] Reinicio sin actualizador valido. Se detiene para rollback.
            set "CITYPC_UPDATE_RC=30"
            exit /b 0
        )
        echo   [AVISO] No se pudo recuperar el actualizador. Se usa V!LOCAL_VER!.
        exit /b 0
    )
)
powershell -NoProfile -ExecutionPolicy Bypass -File "!CITYPC_UPDATER_PATH!" -Mode Check -ToolId "!CITYPC_TOOL_ID!" -CurrentVersion !LOCAL_VER! -TargetPath "%~f0" -ResumeToken "!CITYPC_UPDATE_RESUME!"
set "CITYPC_UPDATE_RC=!errorlevel!"
exit /b 0
