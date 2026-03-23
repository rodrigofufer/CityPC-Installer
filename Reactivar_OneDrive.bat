@echo off
setlocal EnableDelayedExpansion
:: =========================================================
:: Reactivar OneDrive - CityPC
:: Restaura OneDrive completamente a su funcionamiento normal
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

:: Buscar el ejecutable de OneDrive en las rutas conocidas
set "ONEDRIVE_EXE="
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    set "ONEDRIVE_EXE=%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
) else if exist "%ProgramFiles%\Microsoft OneDrive\OneDrive.exe" (
    set "ONEDRIVE_EXE=%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"
) else if exist "%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe" (
    set "ONEDRIVE_EXE=%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe"
) else if exist "%SystemRoot%\SysWOW64\OneDriveSetup.exe" (
    set "ONEDRIVE_SETUP=%SystemRoot%\SysWOW64\OneDriveSetup.exe"
) else if exist "%SystemRoot%\System32\OneDriveSetup.exe" (
    set "ONEDRIVE_SETUP=%SystemRoot%\System32\OneDriveSetup.exe"
)

if not defined ONEDRIVE_EXE (
    if defined ONEDRIVE_SETUP (
        echo   [INFO] OneDrive no esta instalado pero se encontro el instalador.
        echo   Reinstalando OneDrive...
        start /wait "" "!ONEDRIVE_SETUP!" /install
        timeout /t 3 /nobreak >nul
        if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
            set "ONEDRIVE_EXE=%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
            echo   [OK] OneDrive reinstalado correctamente
        ) else (
            echo   [ERROR] No se pudo reinstalar OneDrive
            pause
            exit /b 1
        )
    ) else (
        echo   [ERROR] No se encontro OneDrive en el sistema.
        echo           Descargalo desde https://www.microsoft.com/en-us/microsoft-365/onedrive/download
        pause
        exit /b 1
    )
)

echo   Ejecutable encontrado: !ONEDRIVE_EXE!
echo.

:: =========================================================
:: 1. Eliminar TODAS las politicas de grupo que bloquean OneDrive
:: =========================================================
echo   [1/5] Eliminando politicas de grupo que bloquean OneDrive...
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v DisableFileSyncNGSC /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v DisableFileSync /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /v DisableLibrariesDefaultSaveToOneDrive /f >nul 2>&1
:: Eliminar la clave completa si quedo vacia
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive" /f >nul 2>&1
echo   [OK] Politicas de bloqueo eliminadas

:: =========================================================
:: 2. Restaurar OneDrive en el registro de inicio (Run)
:: =========================================================
echo   [2/5] Restaurando entrada de inicio en el registro...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "OneDrive" /t REG_SZ /d "\"!ONEDRIVE_EXE!\" /background" /f >nul 2>&1
echo   [OK] Entrada de inicio restaurada en HKCU\...\Run

:: =========================================================
:: 3. Habilitar OneDrive en el Administrador de Tareas (StartupApproved)
::    Esta es la clave que controla si aparece habilitado en
::    Administrador de Tareas > Inicio
:: =========================================================
echo   [3/5] Habilitando OneDrive en Administrador de Tareas...
:: Valor 02000000... = Habilitado en el Administrador de Tareas
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "OneDrive" /t REG_BINARY /d 020000000000000000000000 /f >nul 2>&1
if !errorlevel! equ 0 (
    echo   [OK] OneDrive habilitado en Administrador de Tareas ^> Inicio
) else (
    echo   [AVISO] No se pudo escribir en StartupApproved\Run
)

:: =========================================================
:: 4. Rehabilitar OneDrive en Windows (quitar bloqueos adicionales)
:: =========================================================
echo   [4/5] Eliminando bloqueos adicionales del sistema...

:: Quitar prevencion de instalacion de OneDrive
reg delete "HKLM\SOFTWARE\Microsoft\OneDrive" /v PreventNetworkTrafficPreUserSignIn /f >nul 2>&1

:: Rehabilitar la integracion con el shell de Windows
reg delete "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v System.IsPinnedToNameSpaceTree /f >nul 2>&1
reg add "HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f >nul 2>&1

:: Para sistemas 64-bit, tambien en WOW6432Node
reg delete "HKCR\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v System.IsPinnedToNameSpaceTree /f >nul 2>&1
reg add "HKCR\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f >nul 2>&1

echo   [OK] Bloqueos adicionales eliminados

:: =========================================================
:: 5. Iniciar OneDrive
:: =========================================================
echo   [5/5] Iniciando OneDrive...
taskkill /f /im OneDrive.exe >nul 2>&1
timeout /t 2 /nobreak >nul
start "" "!ONEDRIVE_EXE!" /background
echo   [OK] OneDrive iniciado

echo.
echo =========================================================
echo   OneDrive reactivado correctamente.
echo.
echo   - Aparecera en Administrador de Tareas ^> Inicio
echo   - Se sincronizara con normalidad
echo   - Se iniciara automaticamente con Windows
echo.
echo   Si no ves cambios en el Administrador de Tareas,
echo   reinicia el equipo.
echo =========================================================
echo.
pause
