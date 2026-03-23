@echo off
:: =========================================================
:: Reactivar OneDrive - CityPC
:: Restaura OneDrive a su funcionamiento normal
:: =========================================================

:: Verificar permisos de administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Este script requiere permisos de administrador.
    echo Haz clic derecho y selecciona "Ejecutar como administrador".
    pause
    exit /b 1
)

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
